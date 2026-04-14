"""Google スプレッドシートへの物件データ書き込み

スプレッドシートに4つのタブを作成・管理する:
  - 「顧客管理」タブ: お客様情報と検索条件
  - 「案件一覧」タブ: 最新の探索結果（毎回リセット）
  - 「保管庫」タブ: 過去の提案をすべて蓄積
  - 「設定」タブ: config.yaml の検索条件（確認用）

使い方:
  python scripts/sheets_writer.py --init              # タブ初期化（初回のみ）
  python scripts/sheets_writer.py --csv results/properties.csv  # CSV→スプシ
  python scripts/sheets_writer.py --data '[{...}]'    # JSON→スプシ
"""

import argparse
import csv
import json
import sys
from datetime import datetime
from pathlib import Path

import gspread
import yaml
from google.oauth2.service_account import Credentials

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# 「案件一覧」「保管庫」タブのヘッダー
PROPERTY_HEADERS = [
    "取得日時",
    "おすすめ度",
    "おすすめ理由",
    "物件名",
    "会社名",
    "プラットフォーム",
    "価格（万円）",
    "面積",
    "用途地域",
    "最寄駅",
    "駅徒歩（分）",
    "周辺環境",
    "築年数",
    "間取り",
    "物件URL",
    "ステータス",
    "メモ",
]

STATUS_OPTIONS = ["未確認", "要検討", "見送り", "応募済み"]

# 「顧客管理」タブのヘッダー
CUSTOMER_HEADERS = [
    "お客様名",
    "ステータス",
    "予算上限",
    "予算下限",
    "希望エリア",
    "物件種別",
    "最小面積",
    "最大面積",
    "駅徒歩",
    "築年数",
    "除外条件",
    "現在のお住まい",
    "親族の住所",
    "競合",
    "追加希望",
    "メモ",
]


def load_config() -> dict:
    config_path = Path(__file__).parent.parent / "config.yaml"
    with open(config_path, encoding="utf-8") as f:
        return yaml.safe_load(f)


def get_spreadsheet(config: dict):
    sheets_config = config["output"]["google_sheets"]
    creds = Credentials.from_service_account_file(
        sheets_config["credentials_path"],
        scopes=SCOPES,
    )
    gc = gspread.authorize(creds)
    return gc.open_by_key(sheets_config["spreadsheet_id"])


# ─── 「案件一覧」タブ ───

def _apply_property_sheet_format(sheet) -> None:
    """案件一覧・保管庫の共通フォーマットを適用"""
    last_col = chr(ord("A") + len(PROPERTY_HEADERS) - 1)

    sheet.format(f"A1:{last_col}1", {
        "backgroundColor": {"red": 0.24, "green": 0.52, "blue": 0.78},
        "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}},
    })

    reason_col_idx = PROPERTY_HEADERS.index("おすすめ理由")
    reason_col = chr(ord("A") + reason_col_idx)
    sheet.format(f"{reason_col}2:{reason_col}200", {"wrapStrategy": "WRAP"})

    body = {
        "requests": [{
            "updateDimensionProperties": {
                "range": {
                    "sheetId": sheet.id,
                    "dimension": "COLUMNS",
                    "startIndex": reason_col_idx,
                    "endIndex": reason_col_idx + 1,
                },
                "properties": {"pixelSize": 300},
                "fields": "pixelSize",
            }
        }]
    }
    sheet.spreadsheet.batch_update(body)

    status_col_idx = PROPERTY_HEADERS.index("ステータス")
    status_col = chr(ord("A") + status_col_idx)
    validation_body = {
        "requests": [{
            "setDataValidation": {
                "range": {
                    "sheetId": sheet.id,
                    "startRowIndex": 1,
                    "endRowIndex": 200,
                    "startColumnIndex": status_col_idx,
                    "endColumnIndex": status_col_idx + 1,
                },
                "rule": {
                    "condition": {
                        "type": "ONE_OF_LIST",
                        "values": [{"userEnteredValue": v} for v in STATUS_OPTIONS],
                    },
                    "showCustomUi": True,
                    "strict": True,
                },
            }
        }]
    }
    sheet.spreadsheet.batch_update(validation_body)

    sheet.freeze(rows=1)
    sheet.columns_auto_resize(0, len(PROPERTY_HEADERS))


def init_property_sheet(spreadsheet) -> None:
    """「案件一覧」タブを作成してヘッダーを書き込む"""
    try:
        sheet = spreadsheet.worksheet("案件一覧")
        sheet.clear()
    except gspread.WorksheetNotFound:
        sheet = spreadsheet.add_worksheet(title="案件一覧", rows=200, cols=20)

    sheet.append_row(PROPERTY_HEADERS)
    _apply_property_sheet_format(sheet)

    print("「案件一覧」タブを初期化しました。")


def write_properties(spreadsheet, properties: list[dict]) -> None:
    """「案件一覧」タブに物件データを追記"""
    try:
        sheet = spreadsheet.worksheet("案件一覧")
    except gspread.WorksheetNotFound:
        sheet = spreadsheet.add_worksheet(title="案件一覧", rows=200, cols=20)
        sheet.append_row(PROPERTY_HEADERS)

    existing = sheet.get_all_values()
    if not existing:
        sheet.append_row(PROPERTY_HEADERS)

    today = datetime.now().strftime("%Y-%m-%d")

    rows = []
    for p in properties:
        rows.append(_property_to_row(p, today))

    if rows:
        sheet.append_rows(rows)

    print(f"「案件一覧」タブに {len(rows)} 件を書き込みました。")


def _property_to_row(p: dict, today: str) -> list:
    return [
        p.get("取得日時", today),
        p.get("おすすめ度", ""),
        p.get("おすすめ理由", ""),
        p.get("物件名", ""),
        p.get("会社名", ""),
        p.get("プラットフォーム", ""),
        p.get("価格（万円）", ""),
        p.get("面積", ""),
        p.get("用途地域", ""),
        p.get("最寄駅", ""),
        p.get("駅徒歩（分）", ""),
        p.get("周辺環境", ""),
        p.get("築年数", ""),
        p.get("間取り", ""),
        p.get("物件URL", ""),
        p.get("ステータス", "未確認"),
        p.get("メモ", ""),
    ]


# ─── 「保管庫」タブ ───

def init_archive_sheet(spreadsheet) -> None:
    """「保管庫」タブを作成してヘッダーを書き込む（既存データは残す）"""
    try:
        sheet = spreadsheet.worksheet("保管庫")
        existing = sheet.get_all_values()
        if existing:
            return
    except gspread.WorksheetNotFound:
        sheet = spreadsheet.add_worksheet(title="保管庫", rows=500, cols=20)

    sheet.append_row(PROPERTY_HEADERS)
    _apply_property_sheet_format(sheet)

    print("「保管庫」タブを初期化しました。")


def archive_properties(spreadsheet, properties: list[dict]) -> None:
    """「保管庫」タブに物件データを追記（削除せず蓄積）"""
    try:
        sheet = spreadsheet.worksheet("保管庫")
    except gspread.WorksheetNotFound:
        sheet = spreadsheet.add_worksheet(title="保管庫", rows=500, cols=20)
        sheet.append_row(PROPERTY_HEADERS)

    existing = sheet.get_all_values()
    if not existing:
        sheet.append_row(PROPERTY_HEADERS)

    today = datetime.now().strftime("%Y-%m-%d")

    rows = []
    for p in properties:
        rows.append(_property_to_row(p, today))

    if rows:
        sheet.append_rows(rows)

    print(f"「保管庫」タブに {len(rows)} 件を追記しました。")


# ─── 「顧客管理」タブ ───

def init_customer_sheet(spreadsheet) -> None:
    """「顧客管理」タブを作成してヘッダーを書き込む（既存データは残す）"""
    try:
        sheet = spreadsheet.worksheet("顧客管理")
        existing = sheet.get_all_values()
        if existing:
            return
    except gspread.WorksheetNotFound:
        sheet = spreadsheet.add_worksheet(title="顧客管理", rows=100, cols=20)

    sheet.append_row(CUSTOMER_HEADERS)

    last_col = chr(ord("A") + len(CUSTOMER_HEADERS) - 1)
    sheet.format(f"A1:{last_col}1", {
        "backgroundColor": {"red": 0.24, "green": 0.52, "blue": 0.78},
        "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}},
    })
    sheet.freeze(rows=1)
    sheet.columns_auto_resize(0, len(CUSTOMER_HEADERS))

    print("「顧客管理」タブを初期化しました。")


# ─── 「設定」タブ ───

def _to_yes_no(value) -> str:
    """bool や文字列を「はい」「いいえ」に統一"""
    if isinstance(value, bool):
        return "はい" if value else "いいえ"
    s = str(value).strip().lower()
    if s in ("true", "yes", "on", "1", "はい"):
        return "はい"
    return "いいえ"


def _format_header(sheet, row: int, cols: str = "A:C") -> None:
    """ヘッダー行に背景色と太字を適用"""
    cell_range = f"{cols.split(':')[0]}{row}:{cols.split(':')[1]}{row}"
    sheet.format(cell_range, {
        "backgroundColor": {"red": 0.24, "green": 0.52, "blue": 0.78},
        "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}},
    })


def _format_section_header(sheet, row: int, cols: str = "A:C") -> None:
    """セクション見出し行に背景色と太字を適用"""
    cell_range = f"{cols.split(':')[0]}{row}:{cols.split(':')[1]}{row}"
    sheet.format(cell_range, {
        "backgroundColor": {"red": 0.24, "green": 0.52, "blue": 0.78},
        "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}},
    })


def write_settings(spreadsheet, config: dict) -> None:
    """「設定」タブを作成して config.yaml の内容を反映"""
    try:
        sheet = spreadsheet.worksheet("設定")
        sheet.clear()
    except gspread.WorksheetNotFound:
        sheet = spreadsheet.add_worksheet(title="設定", rows=30, cols=3)

    rows = [
        ["項目", "値", "説明"],
        ["取得件数", config.get("取得件数", 5), "この件数だけ探索して案件一覧に保存"],
        ["エリア", config.get("エリア", ""), "探索対象エリア"],
        ["上限価格", config.get("上限価格", ""), "予算上限"],
        ["下限価格", config.get("下限価格", ""), "予算下限"],
        ["最小面積", config.get("最小面積", ""), "土地の最小面積"],
        ["最大面積", config.get("最大面積", ""), "土地の最大面積"],
        ["駅徒歩", config.get("駅徒歩", ""), "最寄り駅からの距離"],
        ["築年数", config.get("築年数", ""), "中古の場合の築年数上限"],
        ["除外条件", config.get("除外条件", ""), "除外するキーワード"],
        ["追加指示", config.get("追加指示", ""), "自然文で追加条件"],
        [],
        ["物件種別", "利用", "探すものだけ「はい」"],
    ]

    property_types = config.get("物件種別", {})
    for name, enabled in property_types.items():
        rows.append([name, _to_yes_no(enabled), ""])

    rows.append([])

    rows.append(["対象プラットフォーム", "利用", "使うものだけ「はい」"])

    platforms = config.get("platforms", {})
    for name, enabled in platforms.items():
        rows.append([name, _to_yes_no(enabled), ""])

    sheet.append_rows(rows)

    header_row = 1
    type_header_row = len(rows) - len(platforms) - 1 - len(property_types)
    platform_header_row = type_header_row + len(property_types) + 1

    _format_header(sheet, header_row)
    _format_section_header(sheet, type_header_row)
    _format_section_header(sheet, platform_header_row)

    sheet.columns_auto_resize(0, 3)

    print("「設定」タブを更新しました。")


# ─── デフォルトの Sheet1 を削除 ───

def remove_default_sheet(spreadsheet) -> None:
    """初期状態の「シート1」があれば削除（案件一覧・設定の2タブだけにする）"""
    for name in ["シート1", "Sheet1"]:
        try:
            default = spreadsheet.worksheet(name)
            spreadsheet.del_worksheet(default)
        except gspread.WorksheetNotFound:
            pass


def main():
    parser = argparse.ArgumentParser(description="物件データをスプレッドシートに書き込む")
    parser.add_argument("--data", help="物件データ（JSON文字列）")
    parser.add_argument("--csv", help="物件データCSVファイルパス")
    parser.add_argument("--init", action="store_true", help="「案件一覧」+「設定」タブを初期化")
    args = parser.parse_args()

    config = load_config()
    spreadsheet = get_spreadsheet(config)

    if args.init:
        init_customer_sheet(spreadsheet)
        init_property_sheet(spreadsheet)
        init_archive_sheet(spreadsheet)
        try:
            settings_sheet = spreadsheet.worksheet("設定")
            existing = settings_sheet.get_all_values()
            if not existing or len(existing) <= 1:
                write_settings(spreadsheet, config)
            else:
                print("「設定」タブは既にデータがあるためスキップしました。")
        except gspread.WorksheetNotFound:
            write_settings(spreadsheet, config)
        remove_default_sheet(spreadsheet)
        print("\n初期化完了。スプレッドシートに「顧客管理」「案件一覧」「保管庫」「設定」の4タブを作成しました。")
        return

    properties = []

    if args.data:
        properties = json.loads(args.data)
    elif args.csv:
        csv_path = Path(args.csv)
        if csv_path.exists():
            with open(csv_path, encoding="utf-8") as f:
                reader = csv.DictReader(f)
                properties = list(reader)

    if not properties:
        print("書き込むデータがありません。", file=sys.stderr)
        sys.exit(1)

    write_properties(spreadsheet, properties)
    archive_properties(spreadsheet, properties)


if __name__ == "__main__":
    main()
