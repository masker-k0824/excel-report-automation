"""pywin32 を使って Outlook からメールを送信するモジュール。"""

from __future__ import annotations

import logging
from pathlib import Path

import pythoncom
import win32com.client

logger = logging.getLogger(__name__)


def send_mail(
    to: str | list[str],
    subject: str,
    body: str,
    attachments: list[str | Path] | None = None,
    cc: str | list[str] | None = None,
) -> None:
    """Outlook を使ってメールを送信する。

    Args:
        to: 宛先メールアドレス（複数の場合はリストまたはセミコロン区切り文字列）
        subject: 件名
        body: 本文
        attachments: 添付ファイルのパスリスト（省略可）
        cc: CCメールアドレス（省略可）
    """
    pythoncom.CoInitialize()
    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)  # 0 = olMailItem

    mail.To = to if isinstance(to, str) else "; ".join(to)
    mail.Subject = subject
    mail.Body = body

    if cc:
        mail.CC = cc if isinstance(cc, str) else "; ".join(cc)

    if attachments:
        for path in attachments:
            resolved = str(Path(path).resolve())
            mail.Attachments.Add(resolved)
            logger.info("添付ファイル追加: %s", resolved)

    to_addr = mail.To
    mail.Send()
    logger.info("メール送信完了: %s", to_addr)
