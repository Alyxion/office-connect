"""Content guard — keeps internal names out of this public repository.

The blocked terms are stored only as salted SHA-256 hashes (scripts/content_guard.sha256), so the
guard never publishes what it protects. Every word of the checked text is hashed and compared:

    s <length> <hash>   blocked anywhere, also inside longer words and domains (compound-safe)
    t <hash>            blocked as a whole token; separators are ignored, so a hashed id
                        also matches when written with dashes or dots

A local, never-committed file ~/.config/office-connect/content-guard.txt can add plaintext
regexes (one per line), e.g. personal numbers that should not even be hashed here.

    python3 scripts/content_guard.py --staged          # pre-commit: staged blobs + file names
    python3 scripts/content_guard.py --message FILE    # commit-msg
    python3 scripts/content_guard.py --tree            # whole working tree (tests / CI)
    python3 scripts/content_guard.py --range A..B      # pre-push: every commit (added lines, names, message)
    python3 scripts/content_guard.py --history         # every commit ever (CI) — history leaks are forever
    python3 scripts/content_guard.py --archive dist/*  # wheel / sdist before a PyPI upload
    python3 scripts/content_guard.py --hash TERM [--substring]   # line to add to the hash list

Letter-class obfuscation like `[Ab][Cd]…` is decoded and checked as well — a regex that spells a
name letter by letter publishes that name just the same.

Hooks: `git config core.hooksPath .githooks` once per clone (see CLAUDE.md).
"""
from __future__ import annotations

import argparse
import hashlib
import io
import re
import subprocess
import sys
import tarfile
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
HASHES = Path(__file__).with_suffix(".sha256")
LOCAL_TERMS = Path.home() / ".config" / "office-connect" / "content-guard.txt"
SALT = b"office-connect/content-guard/v1:"
SKIP_EXT = {".woff", ".woff2", ".png", ".jpg", ".jpeg", ".gif", ".ico", ".zip", ".whl", ".gz", ".pdf", ".pyc"}
SKIP_FILES = {"scripts/content_guard.sha256"}

_WORD = re.compile(r"[0-9a-zäöüß]+")
_RUN = re.compile(r"[0-9a-zäöüß][0-9a-zäöüß._@/-]*[0-9a-zäöüß]")  # word incl. separators: domains, ids
_CLASS_RUN = re.compile(r"(?:\[([A-Za-z])[A-Za-z]\](?:\{\d(?:,\d)?\})?){3,}")  # [Ll][Ee][Cc]… spelled letter by letter


def digest(term: str) -> str:
    return hashlib.sha256(SALT + term.encode()).hexdigest()


def _squash(run: str) -> str:
    return re.sub(r"[._@/-]", "", run)


def load_rules() -> tuple[dict[int, set[str]], set[str], list[re.Pattern]]:
    substrings: dict[int, set[str]] = {}
    tokens: set[str] = set()
    for line in HASHES.read_text().splitlines():
        parts = line.split()
        if not parts or parts[0].startswith("#"):
            continue
        if parts[0] == "s":
            substrings.setdefault(int(parts[1]), set()).add(parts[2])
        elif parts[0] == "t":
            tokens.add(parts[1])
    local = []
    if LOCAL_TERMS.is_file():
        local = [re.compile(ln.strip(), re.I) for ln in LOCAL_TERMS.read_text().splitlines()
                 if ln.strip() and not ln.startswith("#")]
    return substrings, tokens, local


def scan(text: str, rules) -> bool:
    """True when *text* contains a blocked term."""
    substrings, tokens, local = rules
    decoded = ["".join(re.findall(r"\[([A-Za-z])[A-Za-z]\]", m.group(0))) for m in _CLASS_RUN.finditer(text)]
    if decoded:
        text = text + " " + " ".join(decoded)
    lowered = text.lower()
    candidates = set(_WORD.findall(lowered)) | {_squash(r) for r in _RUN.findall(lowered)}
    for word in candidates:
        if digest(word) in tokens:
            return True
        for length, hashes in substrings.items():
            for i in range(len(word) - length + 1):
                if digest(word[i:i + length]) in hashes:
                    return True
    return any(p.search(text) for p in local)


def _git(*args: str) -> bytes:
    return subprocess.run(["git", *args], cwd=ROOT, capture_output=True, check=True).stdout


def _checkable(path: str, data: bytes) -> bool:
    return path not in SKIP_FILES and Path(path).suffix.lower() not in SKIP_EXT and b"\0" not in data[:8192]


def check_files(items: list[tuple[str, bytes]], rules) -> list[str]:
    """Returns 'path:line' locations (never the matched text — the output may end up in CI logs)."""
    found = []
    for path, data in items:
        if scan(path, rules):
            found.append(f"{path} (file name)")
        if not _checkable(path, data):
            continue
        for lineno, line in enumerate(data.decode("utf-8", errors="ignore").splitlines(), 1):
            if scan(line, rules):
                found.append(f"{path}:{lineno}")
    return found


def staged_items() -> list[tuple[str, bytes]]:
    names = _git("diff", "--cached", "--name-only", "--diff-filter=ACMR", "-z").decode().split("\0")
    return [(n, _git("show", f":{n}")) for n in names if n]


def tree_items() -> list[tuple[str, bytes]]:
    names = _git("ls-files", "-co", "--exclude-standard", "-z").decode().split("\0")
    return [(n, (ROOT / n).read_bytes()) for n in names if n and (ROOT / n).is_file()]


def _commit_items(commit: str) -> list[tuple[str, bytes]]:
    """Added lines of one commit (per file) plus its message — what a push publishes."""
    items = [(f"{commit[:9]} (message)", _git("log", "-1", "--format=%B", commit))]
    diff = _git("show", "--format=", "--unified=0", "--no-color", "--no-renames", commit).decode("utf-8", "ignore")
    path, added = "", []
    for line in diff.splitlines():
        if line.startswith("+++ "):
            if path:
                items.append((f"{commit[:9]} {path}", "\n".join(added).encode()))
            path, added = line[6:] if line.startswith("+++ b/") else "", []
        elif line.startswith("+") and path:
            added.append(line[1:])
    if path:
        items.append((f"{commit[:9]} {path}", "\n".join(added).encode()))
    return items


def range_items(spec: str) -> list[tuple[str, bytes]]:
    commits = _git("rev-list", *spec.split()).decode().split()  # "A..B" or "SHA --not --remotes"
    return [item for c in commits for item in _commit_items(c)]


def archive_items(paths: list[str]) -> list[tuple[str, bytes]]:
    items = []
    for archive in paths:
        if archive.endswith(".whl") or archive.endswith(".zip"):
            with zipfile.ZipFile(archive) as z:
                items += [(f"{archive}!{n}", z.read(n)) for n in z.namelist() if not n.endswith("/")]
        else:
            with tarfile.open(archive) as t:
                items += [(f"{archive}!{m.name}", t.extractfile(m).read()) for m in t.getmembers() if m.isfile()]
    return items


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    mode = parser.add_mutually_exclusive_group(required=True)
    mode.add_argument("--staged", action="store_true")
    mode.add_argument("--tree", action="store_true")
    mode.add_argument("--message")
    mode.add_argument("--range", metavar="A..B")
    mode.add_argument("--history", action="store_true")
    mode.add_argument("--archive", nargs="+", metavar="FILE")
    mode.add_argument("--hash", metavar="TERM")
    parser.add_argument("--substring", action="store_true", help="with --hash: block inside longer words too")
    args = parser.parse_args()

    if args.hash:
        term = args.hash.lower() if args.substring else _squash(args.hash.lower())
        print(f"s {len(term)} {digest(term)}" if args.substring else f"t {digest(term)}")
        return 0
    rules = load_rules()
    if args.message:
        text = Path(args.message).read_text(errors="ignore")
        found = ["commit message"] if scan(text, rules) else []
    elif args.range or args.history:
        found = check_files(range_items(args.range or "--all"), rules)
    elif args.archive:
        found = check_files(archive_items(args.archive), rules)
    else:
        found = check_files(staged_items() if args.staged else tree_items(), rules)
    if found:
        print("\n[CONTENT GUARD] internal terms found — this repository is public:\n", file=sys.stderr)
        for location in found:
            print(f"  {location}", file=sys.stderr)
        print(f"\n  {len(found)} location(s) — remove them (the guard never prints the term).\n", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
