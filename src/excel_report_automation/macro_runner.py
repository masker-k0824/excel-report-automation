"""xlwings を使ってExcelファイルのマクロを実行するモジュール。"""

from __future__ import annotations

import logging
import threading
from pathlib import Path

import pythoncom
import xlwings as xw

logger = logging.getLogger(__name__)


def run_macro(
    excel_path: str | Path,
    macro_name: str,
    *macro_args: object,
    timeout_min: int | None = None,
    visible: bool = False,
) -> None:
    """Excelファイルを開いて指定マクロを実行し、保存して閉じる。

    タイムアウト時は Excel を強制終了して TimeoutError を送出する。

    Args:
        excel_path: 実行対象の .xlsm / .xlsx ファイルパス
        macro_name: 実行するVBAマクロ名（例: "Module1.RunReport"）
        *macro_args: マクロに渡す引数（省略可）
        timeout_min: タイムアウト時間（分）。None の場合は無制限。
        visible: Excelウィンドウを表示するか（デバッグ時に True）
    """
    excel_path = Path(excel_path).resolve()
    if not excel_path.exists():
        raise FileNotFoundError(f"Excelファイルが見つかりません: {excel_path}")

    logger.info("Excel を開きます: %s", excel_path)

    app_holder: list[xw.App | None] = [None]
    error_holder: list[BaseException | None] = [None]

    def _execute() -> None:
        pythoncom.CoInitialize()
        try:
            app = xw.App(visible=visible, add_book=False)
            app_holder[0] = app
            try:
                wb = app.books.open(str(excel_path))
                try:
                    logger.info("マクロを実行します: %s", macro_name)
                    macro = app.macro(macro_name)
                    if macro_args:
                        macro(*macro_args)
                    else:
                        macro()
                    wb.save()
                    logger.info("マクロ完了・保存しました")
                finally:
                    wb.close()
                app.quit()
            except Exception:
                try:
                    app.kill()
                except Exception:
                    pass
                raise
        except Exception as e:
            error_holder[0] = e
        finally:
            pythoncom.CoUninitialize()

    thread = threading.Thread(target=_execute, daemon=True)
    thread.start()
    timeout_sec = timeout_min * 60 if timeout_min is not None else None
    thread.join(timeout=timeout_sec)

    if thread.is_alive():
        logger.error("タイムアウト (%d分)：Excel を強制終了します", timeout_min)
        app = app_holder[0]
        if app is not None:
            try:
                app.kill()
            except Exception:
                pass
        raise TimeoutError(f"マクロがタイムアウトしました ({timeout_min}分): {macro_name}")

    if error_holder[0] is not None:
        raise error_holder[0]
