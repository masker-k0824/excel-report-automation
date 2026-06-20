# Excel Report Automation

Excel のマクロを自動実行し、結果を Outlook でメール送信するツールです。

## 構成

```
excel-report-automation/
├── config.yaml              # 実行設定（ファイルパス・メール設定）
├── macro_manager.xlsx       # マクロ実行リスト（macro_settings シート）
├── run.bat                  # 実行用バッチファイル
├── run.log                  # 実行ログ（自動生成）
├── src/
│   └── excel_report_automation/
│       ├── main.py          # エントリポイント
│       ├── config_reader.py # macro_manager.xlsx の読み込み
│       ├── macro_runner.py  # xlwings によるマクロ実行
│       └── mail_sender.py   # Outlook メール送信
└── tests/
```

## セットアップ

```bat
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
pip install -e .
```

## 設定

### config.yaml

```yaml
manager_file: "C:/path/to/macro_manager.xlsx"

mail:
  to: "送信先@example.com"
  cc: ""
  subject: "【自動送信】Excelレポート実行完了"
```

### macro_manager.xlsx（macro_settings シート）

| order | enabled | file_path | macro_name | timeout_min | on_error | notes |
|-------|---------|-----------|------------|-------------|----------|-------|
| 1 | 1 | C:/path/to/file.xlsm | Module1.Main | 5 | continue | 説明 |
| 2 | 0 | C:/path/to/file.xlsm | Module1.Main | 10 | stop | 無効中 |

- **enabled**: `1` で実行、それ以外はスキップ
- **on_error**: `stop` でエラー時に中断、`continue` で次のタスクへ進む
- **timeout_min**: 指定時間を超えると Excel を強制終了してタイムアウト扱い

## 実行

```bat
run.bat
```

## タスクスケジューラへの登録

1. タスクスケジューラを開く
2. 「基本タスクの作成」→ トリガー: 毎日 9:00
3. 操作: プログラムの開始
   - プログラム: `C:\path\to\excel-report-automation\run.bat`
   - 開始(オプション): `C:\path\to\excel-report-automation`
