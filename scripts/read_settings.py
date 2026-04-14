"""スプレッドシートの「設定」タブから検索条件を読み取る

使い方:
  python scripts/read_settings.py          # JSON で出力
  python scripts/read_settings.py --pretty  # 見やすく出力
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


def read_settings_from_sheet(config: dict) -> dict:
    sheets_config = config["output"]["google_sheets"]
    creds = Credentials.from_service_account_file(
        sheets_config["credentials_path"],
        scopes=SCOPES,
    )
    gc = gspread.authorize(creds)
    spreadsheet = gc.open_by_key(sheets_config["spreadsheet_id"])

    try:
        sheet = spreadsheet.worksheet("設定")
    except gspread.WorksheetNotFound:
        print("「設定」タブが見つかりません。", file=sys.stderr)
        sys.exit(1)

    rows = sheet.get_all_values()

    settings = {}
    property_types = {}
    platforms = {}
    section = "main"

    for row in rows:
        if not row or not row[0]:
            continue

        key = str(row[0]).strip()
        value = str(row[1]).strip() if len(row) > 1 else ""

        if key == "項目":
            continue

        if key == "物件種別":
            section = "property_types"
            continue
        elif key == "対象プラットフォーム":
            section = "platforms"
            continue

        if section == "main":
            settings[key] = value
        elif section == "property_types":
            property_types[key] = value
        elif section == "platforms":
            platforms[key] = value

    settings["物件種別"] = property_types
    settings["platforms"] = platforms

    return settings


def main():
    parser = argparse.ArgumentParser(description="設定タブから検索条件を読み取る")
    parser.add_argument("--pretty", action="store_true", help="見やすく出力")
    args = parser.parse_args()

    config = load_config()
    settings = read_settings_from_sheet(config)

    if args.pretty:
        print("━━━ スプレッドシート「設定」タブの内容 ━━━")
        for key, value in settings.items():
            if isinstance(value, dict):
                print(f"\n{key}:")
                for k, v in value.items():
                    print(f"  {k}: {v}")
            else:
                print(f"{key}: {value}")
    else:
        print(json.dumps(settings, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
