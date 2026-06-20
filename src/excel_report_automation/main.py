"""スクリプトのエントリポイント。タスクスケジューラから直接呼び出す。"""

from __future__ import annotations

import logging
import sys
from pathlib import Path

import yaml

from excel_report_automation.config_reader import load_tasks
from excel_report_automation.macro_runner import run_macro
from excel_report_automation.mail_sender import send_mail

# プロジェクト直下の config.yaml を読み込む
_PROJECT_ROOT = Path(__file__).parent.parent.parent
_CONFIG_PATH = _PROJECT_ROOT / "config.yaml"

with _CONFIG_PATH.open(encoding="utf-8") as _f:
    _cfg = yaml.safe_load(_f)

MANAGER_FILE = Path(_cfg["manager_file"])
MAIL_TO: str = _cfg["mail"]["to"]
MAIL_CC: str = _cfg["mail"].get("cc", "")
MAIL_SUBJECT: str = _cfg["mail"]["subject"]

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler(_PROJECT_ROOT / "run.log", encoding="utf-8"),
    ],
)

logger = logging.getLogger(__name__)


def main() -> None:
    logger.info("=== Excel レポート自動化 開始 ===")

    tasks = load_tasks(MANAGER_FILE)
    if not tasks:
        logger.warning("有効なタスクがありません。処理を終了します。")
        return

    results: list[str] = []
    has_error = False

    for task in tasks:
        if not task.enabled:
            logger.info("[%d] スキップ (無効): %s", task.order, task.notes)
            continue

        logger.info("[%d] 開始: %s / %s", task.order, task.file_path.name, task.macro_name)
        try:
            run_macro(task.file_path, task.macro_name, timeout_min=task.timeout_min)
            logger.info("[%d] 完了: %s", task.order, task.notes)
            results.append(f"[成功] order={task.order} {task.file_path.name} ({task.notes})")
        except Exception as e:
            has_error = True
            logger.error("[%d] エラー: %s", task.order, e)
            results.append(f"[失敗] order={task.order} {task.file_path.name} ({task.notes}) - {e}")
            if task.on_error == "stop":
                logger.error("on_error=stop のため処理を中断します")
                break

    status = "一部エラーあり" if has_error else "正常完了"
    body = f"Excelマクロの実行が完了しました。\n\nステータス: {status}\n\n" + "\n".join(results)

    try:
        send_mail(
            to=MAIL_TO,
            subject=f"{MAIL_SUBJECT} [{status}]",
            body=body,
            cc=MAIL_CC or None,
        )
    except Exception:
        logger.exception("メール送信に失敗しました")

    logger.info("=== 終了: %s ===", status)
    if has_error:
        sys.exit(1)


if __name__ == "__main__":
    main()
