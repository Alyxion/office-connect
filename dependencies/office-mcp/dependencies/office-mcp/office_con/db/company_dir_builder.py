from pathlib import Path
from typing import List, Optional, Dict, Any
import json
from pydantic import BaseModel, Field
import shutil
import logging
from office_con.msgraph.ms_graph_handler import MsGraphInstance
from office_con.db.company_dir import CompanyUserList, CompanyUser, CompanyDirData


class UserReplacements(BaseModel):
    """
    User replacements patterns.

    See user_replacements.md for an example.
    """
    guesses: List[Dict[str, Any]] = Field(default_factory=list, description="Guess patterns for user replacement")
    data: List[Dict[str, Any]] = Field(default_factory=list, description="Definitive user replacement data")


def apply_user_replacements(user_list: CompanyUserList, replacements: UserReplacements, *, add_genders: bool = False) -> None:
    """Apply user_replacements rules to a CompanyUserList in-place (sync).

    Runs guesses first, strips empty-id users, optionally guesses genders,
    then applies definitive data rules.
    """
    if not user_list.users:
        return

    def _apply(rule_list: list, guess: bool = False) -> None:
        for rule in rule_list:
            for user in user_list.users:
                if user.matches_rule(rule, use_guessed=guess):
                    for key, value in rule.items():
                        if key.startswith("_"):
                            continue
                        if guess:
                            if getattr(user, key) != value:
                                user.guessed_fields[key] = True
                        else:
                            if key in user.guessed_fields:
                                del user.guessed_fields[key]
                        setattr(user, key, value)

    _apply(replacements.guesses, guess=True)
    user_list.users = [u for u in user_list.users if len(u.id) > 0]
    if add_genders:
        _guess_genders_sync(user_list)
    _apply(replacements.data)


def _guess_genders_sync(user_list: CompanyUserList) -> None:
    """Synchronous gender guessing (mirrors CompanyDirBuilder.guess_genders)."""
    try:
        from names_dataset import NameDataset, NameWrapper
    except ImportError:
        return
    nd = NameDataset()
    for user in user_list.users:
        if user.gender == "undefined" and user.object_type == "user":
            if not user.given_name:
                continue
            first_name = user.given_name
            display_name = user.display_name or ""
            first_name = first_name.strip("Prof. ").strip("Dr. ").strip(" ")
            first_name = first_name.split("-")[0]
            if "(" in display_name and ")" in display_name:
                european_name = display_name.split("(", 1)[1].split(")", 1)[0].strip()
                if european_name:
                    first_name = european_name
                else:
                    first_name = first_name.split(" ")[0]
            else:
                first_name = first_name.split(" ")[0]
            gender = NameWrapper(nd.search(first_name)).describe
            user.guessed_fields["gender"] = True
            if gender.startswith("Male"):
                user.gender = "male"
            elif gender.startswith("Female"):
                user.gender = "female"
            else:
                user.gender = "other"
                if user.given_name:
                    gender = NameWrapper(nd.search(user.given_name)).describe
                    if gender.startswith("Male"):
                        user.gender = "male"
                    elif gender.startswith("Female"):
                        user.gender = "female"


class CompanyDirBuilder:
    """
    Builds and manages a local company directory from Microsoft Graph data.

    Downloads all users and images, applies replacements, and stores them in a directory structure:

        <target_dir>/users.json
        <target_dir>/photos/<user_id>.jpg

    :param target_dir: Directory to store the users.json and photos
    :param msgraph: Microsoft Graph instance to fetch users and images
    :param logger: Logger for status and error messages
    :param clear_images: Whether to clear existing images before loading
    :param user_replacements: UserReplacements patterns to apply
    :param add_genders: Guess gender for users based on first names (requires names_dataset)
    """
    USERS_JSON = "users.json"
    PHOTO_DIR = "photos"

    def __init__(self, *, target_dir: Path, msgraph: MsGraphInstance, logger: logging.Logger, clear_images: bool = True,
                user_replacements: 'UserReplacements', add_genders: bool = False) -> None:
        self.target_dir = Path(target_dir)
        self.msgraph: MsGraphInstance = msgraph
        self.logger: logging.Logger | None = logger
        self.photos_dir = self.target_dir / self.PHOTO_DIR
        self.user_list: Optional[CompanyUserList] = None
        self.user_replacements: UserReplacements = user_replacements
        if clear_images:
            self.clear()
        self.target_dir.mkdir(parents=True, exist_ok=True)
        self.photos_dir.mkdir(exist_ok=True)
        self.add_genders = add_genders
        self.data = CompanyDirData()

    async def apply_replacements(self) -> None:
        if len(self.user_list.users) == 0:
            return
        try:
            apply_user_replacements(self.user_list, self.user_replacements, add_genders=self.add_genders)
        except Exception as e:
            if self.logger:
                self.logger.error(f"Error applying replacements: {e}")
            raise

    async def guess_genders(self):
        from names_dataset import NameDataset, NameWrapper
        nd = NameDataset()
        for user in self.user_list.users:
            user: CompanyUser
            if user.gender == "undefined" and user.object_type == "user":
                if not user.given_name:
                    continue
                first_name = user.given_name
                display_name = user.display_name
                # remove prof, dr. etc.
                first_name = first_name.strip("Prof. ").strip("Dr. ").strip(" ")
                first_name = first_name.split("-")[0]
                # European names in parentheses in Chinese names
                if "(" in display_name and ")" in display_name:
                    european_name = display_name.split("(", 1)[1].split(")", 1)[0].strip()
                    if european_name:
                        first_name = european_name
                    else:
                        first_name = first_name.split(" ")[0]
                else:
                    first_name = first_name.split(" ")[0]
                gender = NameWrapper(nd.search(first_name)).describe
                user.guessed_fields["gender"] = True
                if gender.startswith("Male"):
                    user.gender = "male"
                elif gender.startswith("Female"):
                    user.gender = "female"
                else:
                    user.gender = "other"
                    if user.given_name:
                        gender = NameWrapper(nd.search(user.given_name)).describe
                        if gender.startswith("Male"):
                            user.gender = "male"
                        elif gender.startswith("Female"):
                            user.gender = "female"
                
    async def build(self, *args: object, **kwargs: object) -> None:
        """
        Fetch all users and their images, and store them locally.
        fetch_users: returns list of user dicts
        fetch_user_image: takes user_id, returns image bytes or None
        """
        if self.logger:
            self.logger.info("Fetching users from directory...")
        dir_list = await self.msgraph.directory.get_all_users_async()
        users = [CompanyUser(**u.model_dump()) for u in dir_list.users]
        if self.logger:
            self.logger.info(f"Fetched {len(users)} users from directory.")
        self.user_list = CompanyUserList(users=users)
        self.data.users = self.user_list
        if self.logger:
            self.logger.info("Saved users.json.")
            self.logger.info("Fetching user images...")
        # scan image dir
        all_files_in_image_dir = set([f.name for f in self.photos_dir.iterdir()])
        total_count = 0
        # Save images
        for index, user in enumerate(users):
            user: CompanyUser
            if index%100 == 0 and self.logger:
                self.logger.info(f"Fetching {index+1} of {len(users)} user images...")
            # check if image exists
            img_path = self.photos_dir / f"{user.id}.jpg"
            placeholder_path = self.photos_dir / f"__{user.id}.ph"
            if img_path.name in all_files_in_image_dir or placeholder_path.name in all_files_in_image_dir:
                user.has_image = img_path.name in all_files_in_image_dir
                continue
            img_bytes = await self.msgraph.directory.get_user_photo_async(user.id)
            if img_bytes:
                with open(img_path, "wb") as imgf:
                    imgf.write(img_bytes)
                user.has_image = True
            else:
                # put placeholder there to flag it as non-existent
                with open(placeholder_path, "wb") as imgf:
                    imgf.write(b'')
        if self.logger:
            self.logger.info(f"Saved {total_count} user images.")
        await self.apply_replacements()
        with (self.target_dir / self.USERS_JSON).open("w", encoding="utf-8") as f:
            json.dump(self.user_list.model_dump(), f, ensure_ascii=False, indent=2)

    def clear(self):
        """Remove all data in the target directory."""
        if (self.target_dir / self.USERS_JSON).exists():
            (self.target_dir / self.USERS_JSON).unlink()
        if self.photos_dir.exists():
            shutil.rmtree(self.photos_dir)
        self.photos_dir.mkdir(exist_ok=True)
