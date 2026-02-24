"""Outlook email draft creation via COM automation.

Creates draft emails in the user's Outlook Drafts folder without
opening a compose window or sending. Preserves the default Outlook
signature by reading it from the freshly created mail item and
inserting body content before it.

Requires: pywin32 (win32com)
"""

from __future__ import annotations

from pathlib import Path

import win32com.client

# Outlook constants
OL_MAIL_ITEM = 0
OL_FORMAT_HTML = 2


def create_draft(
    to: str,
    subject: str,
    body_html: str,
    cc: str | None = None,
    attachment_paths: list[str | Path] | None = None,
) -> dict:
    """Create an Outlook draft email and save it to the Drafts folder.

    Parameters
    ----------
    to : str
        Recipient email address(es), semicolon-separated.
    subject : str
        Email subject line.
    body_html : str
        HTML body content (inserted before the default signature).
    cc : str | None
        CC email address(es), semicolon-separated.
    attachment_paths : list | None
        Paths to files to attach (typically PDFs).

    Returns
    -------
    dict
        Metadata: ``{"entry_id": str, "subject": str, "to": str, "attachments": int}``

    Raises
    ------
    OSError
        If Outlook is not running or COM connection fails.
    FileNotFoundError
        If an attachment path does not exist.
    """
    # Validate attachments exist before touching Outlook
    resolved: list[Path] = []
    for p in attachment_paths or []:
        path = Path(p)
        if not path.exists():
            raise FileNotFoundError(f"Attachment not found: {path}")
        resolved.append(path)

    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
    except Exception as exc:
        raise OSError(
            "Could not connect to Outlook. Is it running?\n"
            f"  Detail: {exc}"
        ) from exc

    mail = outlook.CreateItem(OL_MAIL_ITEM)
    mail.To = to
    mail.Subject = subject
    mail.BodyFormat = OL_FORMAT_HTML

    if cc:
        mail.CC = cc

    # Preserve default signature: Outlook populates HTMLBody with the
    # signature when a new mail item is created. We read it, then
    # insert our body content before the signature.
    signature_html = mail.HTMLBody or ""

    if signature_html and "<body" in signature_html.lower():
        # Insert our content right after the <body...> tag
        import re
        body_tag_match = re.search(r"(<body[^>]*>)", signature_html, re.IGNORECASE)
        if body_tag_match:
            insert_pos = body_tag_match.end()
            mail.HTMLBody = (
                signature_html[:insert_pos]
                + body_html
                + signature_html[insert_pos:]
            )
        else:
            mail.HTMLBody = body_html + signature_html
    else:
        mail.HTMLBody = body_html

    # Add attachments
    for path in resolved:
        mail.Attachments.Add(str(path))

    # Save to Drafts (do NOT send)
    mail.Save()

    entry_id = ""
    try:
        entry_id = mail.EntryID
    except Exception:
        pass  # EntryID may not be available in all Outlook versions

    return {
        "entry_id": entry_id,
        "subject": subject,
        "to": to,
        "attachments": len(resolved),
    }
