"""Exceptions shared by the Graph handlers."""


class GraphRequestError(RuntimeError):
    """Raised by operations that promise a COMPLETE result (paged listings,
    conditional writes) when Graph answers with an unexpected status.

    Most handlers return an empty result on failure, which is fine for an
    interactive lookup but fatal for a full listing: a rejected query once
    surfaced as "no mails with attachments" and silently hid every receipt.
    """

    def __init__(self, message: str, status_code: int | None = None, body: str = "") -> None:
        super().__init__(f"{message} (HTTP {status_code}){': ' + body[:300] if body else ''}")
        self.status_code = status_code
        self.body = body


class DriveItemConflictError(GraphRequestError):
    """A conditional drive write lost: the item changed since its eTag was read
    (HTTP 412), or it already exists although ``create_only`` was requested (409)."""
