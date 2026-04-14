"""顧客管理タブから顧客情報を読み取る

使い方:
  python scripts/read_customers.py                    # 全顧客を表示
  python scripts/read_customers.py --name "田中様"     # 特定の顧客を表示
  python scripts/read_customers.py --active            # 物件探し中の顧客のみ
"""

import argparse
import json
import sys
from pathlib import Path

import gspread
import yaml
from google.oauth2.service_account import Credentials

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive.readonly",
]


def load_config() -> dict:
    config_path = Path(__file__).parent.parent / "config.yaml"
    with open(config_path, encoding="utf-8") as f:
        return yaml.safe_load(f)


def get_customers(config: dict, name: str = None, active_only: bool = False) -> list[dict]:
    sheets_config = config["output"]["google_sheets"]
    creds = Credentials.from_service_account_file(
        sheets_config["credentials_path"],
        scopes=SCOPES,
    )
    gc = gspread.authorize(creds)
    spreadsheet = gc.open_by_key(sheets_config["spreadsheet_id"])

    try:
        sheet = spreadsheet.worksheet("顧客管理")
    except gspread.WorksheetNotFound:
        print("「顧客管理」タブが見つかりません。--init で初期化してください。", file=sys.stderr)
        sys.exit(1)

    records = sheet.get_all_records()

    if name:
        records = [r for r in records if r.get("お客様名") == name]

    if active_only:
        records = [r for r in records if r.get("ステータス") in ("物件探し中", "検討中")]

    return records


def main():
    parser = argparse.ArgumentParser(description="顧客管理タブから顧客情報を読み取る")
    parser.add_argument("--name", help="顧客名で絞り込み")
    parser.add_argument("--active", action="store_true", help="物件探し中・検討中の顧客のみ")
    args = parser.parse_args()

    config = load_config()
    customers = get_customers(config, name=args.name, active_only=args.active)

    if not customers:
        print("該当する顧客が見つかりません。", file=sys.stderr)
        sys.exit(1)

    print(json.dumps(customers, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
