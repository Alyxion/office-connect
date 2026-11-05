"""The content guard keeps internal names out of this public repository (explicit rule).

The real blocked terms exist only as salted hashes in scripts/content_guard.sha256, so these
tests use made-up terms. `test_working_tree_is_clean` runs the guard over every tracked and
untracked file, which makes the rule hold in CI and on clones without the git hooks too.
"""
from __future__ import annotations

import importlib.util
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
_spec = importlib.util.spec_from_file_location("content_guard", ROOT / "scripts" / "content_guard.py")
guard = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(guard)


def _rules(substrings: list[str] = (), tokens: list[str] = ()):
    subs: dict[int, set[str]] = {}
    for term in substrings:
        subs.setdefault(len(term), set()).add(guard.digest(term))
    return subs, {guard.digest(t) for t in tokens}, []


def test_substring_terms_match_inside_words_and_domains():
    rules = _rules(substrings=["examplecorp"])
    assert guard.scan("see https://ai.ExampleCorp.com/chat", rules)
    assert guard.scan("ExampleCorpConnect", rules)
    assert guard.scan("example-corp is fine?", rules)  # separators are squashed inside a run
    assert not guard.scan("an example of a corp", rules)


def test_token_terms_match_whole_ids_with_separators():
    rules = _rules(tokens=["12345678abcd"])
    assert guard.scan("tenant 1234-5678-abcd", rules)
    assert guard.scan("id=12345678ABCD", rules)
    assert not guard.scan("12345678abcde", rules)


def test_hash_list_contains_no_plaintext():
    for line in (ROOT / "scripts" / "content_guard.sha256").read_text().splitlines():
        parts = line.split()
        if parts and not parts[0].startswith("#"):
            assert parts[0] in ("s", "t") and len(parts[-1]) == 64


def test_working_tree_is_clean():
    assert guard.check_files(guard.tree_items(), guard.load_rules()) == []
