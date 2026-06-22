from office_con.msgraph.files_handler import DriveItem, FilesHandler
from office_con.msgraph.mail_handler import OfficeMail, OfficeMailHandler
from office_con.privacy import (
    OfficeFileFilterRules,
    OfficeFilterRules,
    OfficeFolderBlock,
    OfficePrivacyConfig,
)


class DummyGraph:
    def __init__(self, privacy_settings: OfficePrivacyConfig) -> None:
        self.privacy_settings = privacy_settings


def test_file_metadata_privacy_hides_and_blocks_content() -> None:
    handler = FilesHandler(DummyGraph(OfficePrivacyConfig(
        files=OfficeFileFilterRules(
            hidden_name_terms=["secret"],
            blocked_content_name_terms=["confidential"],
        ),
    )))

    visible = handler._filter_items([
        DriveItem(id="1", name="public.docx"),
        DriveItem(id="2", name="secret-plan.docx"),
        DriveItem(id="3", name="confidential-roadmap.docx"),
    ])

    assert [item.id for item in visible] == ["1", "3"]
    assert visible[1].access_status == "content_blocked"
    assert "privacy" in visible[1].access_reason


def test_file_privacy_marks_blocked_folder_descendants() -> None:
    handler = FilesHandler(DummyGraph(OfficePrivacyConfig(
        files=OfficeFileFilterRules(
            blocked_folders=[
                OfficeFolderBlock(
                    item_id="folder-1",
                    name="HR",
                    drive_id="drive-1",
                    parent_path="/drive/root:",
                )
            ],
        ),
    )))

    visible = handler._filter_items([
        DriveItem(id="folder-1", name="HR", drive_id="drive-1", is_folder=True, parent_path="/drive/root:"),
        DriveItem(id="child-1", name="salary.xlsx", drive_id="drive-1", parent_id="folder-1", parent_path="/drive/root:/HR"),
    ])

    assert visible[0].access_status == "folder_blocked"
    assert visible[1].access_status == "folder_blocked"


def test_file_privacy_blocks_individual_item() -> None:
    handler = FilesHandler(DummyGraph(OfficePrivacyConfig(
        files=OfficeFileFilterRules(
            blocked_items=[OfficeFolderBlock(item_id="2", name="budget.xlsx", drive_id="drive-1")],
        ),
    )))

    visible = handler._filter_items([
        DriveItem(id="1", name="public.docx", drive_id="drive-1"),
        DriveItem(id="2", name="budget.xlsx", drive_id="drive-1"),
        DriveItem(id="2", name="budget.xlsx", drive_id="other-drive"),
    ])

    assert [item.access_status for item in visible] == ["ok", "content_blocked", "ok"]


def test_mail_privacy_hides_or_blocks_content() -> None:
    handler = OfficeMailHandler(DummyGraph(OfficePrivacyConfig(
        mail=OfficeFilterRules(
            hidden_name_terms=["password"],
            blocked_content_terms=["do not forward"],
        ),
    )))

    visible = handler._filter_mail_list([
        OfficeMail(email_id="1", email_type="mail", subject="Weekly update", body_preview="All good"),
        OfficeMail(email_id="2", email_type="mail", subject="Password reset", body_preview="Reset link"),
        OfficeMail(email_id="3", email_type="mail", subject="Plan", body_preview="Do not forward this"),
    ])

    assert [mail.email_id for mail in visible] == ["1", "3"]
    assert visible[1].access_status == "content_blocked"
    assert visible[1].body_preview is None


def test_mail_and_file_rules_are_independent() -> None:
    config = OfficePrivacyConfig(
        mail=OfficeFilterRules(hidden_name_terms=["password"]),
        files=OfficeFileFilterRules(hidden_name_terms=["secret"]),
    )

    files = FilesHandler(DummyGraph(config))._filter_items([
        DriveItem(id="1", name="password.txt"),  # only a mail rule → stays
        DriveItem(id="2", name="secret.txt"),
    ])
    assert [item.id for item in files] == ["1"]

    mails = OfficeMailHandler(DummyGraph(config))._filter_mail_list([
        OfficeMail(email_id="1", email_type="mail", subject="secret memo"),  # only a file rule → stays
        OfficeMail(email_id="2", email_type="mail", subject="password reset"),
    ])
    assert [mail.email_id for mail in mails] == ["1"]


def test_disabled_surface_skips_filtering() -> None:
    config = OfficePrivacyConfig(
        mail=OfficeFilterRules(enabled=False, hidden_name_terms=["password"]),
        files=OfficeFileFilterRules(enabled=False, hidden_name_terms=["secret"]),
    )

    files = FilesHandler(DummyGraph(config))._filter_items([DriveItem(id="2", name="secret.txt")])
    assert [item.id for item in files] == ["2"]

    mails = OfficeMailHandler(DummyGraph(config))._filter_mail_list(
        [OfficeMail(email_id="2", email_type="mail", subject="password reset")]
    )
    assert [mail.email_id for mail in mails] == ["2"]
