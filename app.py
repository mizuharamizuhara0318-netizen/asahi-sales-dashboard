import os
import io
import time
import json
import glob
import re
import base64
import warnings
from datetime import datetime, timedelta
from concurrent.futures import ThreadPoolExecutor
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.drawing.image import Image as OpenpyxlImage
from openpyxl.utils import get_column_letter

warnings.filterwarnings("ignore")
st.set_page_config(page_title="売上管理＆営業分析表", layout="wide")

st.markdown("""
<style>
    .main-header { background-color: #2a75d3; color: white; padding: 8px 15px; font-size: 19px; font-weight: bold; border-radius: 4px; margin-bottom: 10px; }
    .total-card { background-color: #f0f4f8; padding: 10px 14px; border-radius: 6px; border-left: 5px solid #2a75d3; margin-bottom: 8px; }
    .side-panel-header { background: linear-gradient(135deg, #1e3a8a, #2563eb); color: white; padding: 8px 14px; border-radius: 6px; font-size: 15px; font-weight: bold; margin-bottom: 8px; }
    .est-card { border: 1px solid #cbd5e1; border-radius: 6px; background-color: #ffffff; margin-bottom: 16px; overflow: hidden; }
    .est-header { background-color: #f8fafc; padding: 10px 14px; border-bottom: 1px solid #cbd5e1; display: flex; justify-content: space-between; align-items: center; }
    .badge-green { background-color: #16a34a; color: white; padding: 2px 6px; border-radius: 3px; font-size: 11px; font-weight: bold; }
    @media (max-width: 768px) {
        .main-header { font-size: 16px; padding: 6px 10px; }
        .total-card { font-size: 12px; padding: 8px 10px; }
    }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-header">売上管理＆営業分析表</div>', unsafe_allow_html=True)

BASE_SHARE = r"G:\マイドライブ\売上ダッシュボード"
OTHER_DATA_DIR = os.path.join(BASE_SHARE, "その他データ")
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
GDRIVE_SYNC_DIR = os.path.join(BASE_DIR, "_gdrive_sync_cache")
os.makedirs(GDRIVE_SYNC_DIR, exist_ok=True)
AMZ_DAILY_STORE_DIR = os.path.join(BASE_DIR, "amazon_daily_csvs")
os.makedirs(AMZ_DAILY_STORE_DIR, exist_ok=True)

CACHE_FILE = os.path.join(BASE_DIR, "sales_cache.parquet")
SAVED_QUOTES_FILE = os.path.join(BASE_SHARE, "saved_estimates.json")
LOCAL_QUOTES_FILE = os.path.join(BASE_DIR, "saved_estimates.json")
SAVED_MEETING_FILE = os.path.join(BASE_SHARE, "saved_meeting_sheet.json")
LOCAL_MEETING_FILE = os.path.join(BASE_DIR, "saved_meeting_sheet.json")

STAFF_MAP = {
    "1": "小泉", "2": "河端", "3": "安部", "4": "原",
    "5": "西山", "6": "萩原", "7": "田中", "8": "水原",
    "9": "川村", "10": "穴山", "11": "鈴木", "12": "松平"
}

DEFAULT_STAFF_TARGETS = {
    "小泉": {"monthly": {7: 8000000, 8: 8000000, 9: 7000000, 10: 7000000, 11: 7040000, 12: 8000000, 1: 6500000, 2: 6500000, 3: 6500000, 4: 5000000, 5: 5000000, 6: 5460000}, "h1": 45040000, "h2": 34960000, "annual": 80000000},
    "水原": {"monthly": {7: 15500000, 8: 15000000, 9: 16000000, 10: 16000000, 11: 15500000, 12: 17710000, 1: 13000000, 2: 13000000, 3: 13000000, 4: 12000000, 5: 11290000, 6: 12000000}, "h1": 95710000, "h2": 74290000, "annual": 170000000},
    "西山": {"monthly": {7: 14450000, 8: 15000000, 9: 15040000, 10: 15040000, 11: 15040000, 12: 15040000, 1: 11847000, 2: 11840000, 3: 12390000, 4: 11260000, 5: 11260000, 6: 11760000}, "h1": 89610000, "h2": 70357000, "annual": 159967000},
    "松平": {"monthly": {7: 11000000, 8: 15000000, 9: 13160000, 10: 13160000, 11: 14160000, 12: 16090000, 1: 13388000, 2: 13360000, 3: 12360000, 4: 9040000, 5: 9040000, 6: 10110000}, "h1": 82570000, "h2": 67298000, "annual": 149868000},
    "河端": {"monthly": {7: 5165000, 8: 0, 9: 0, 10: 0, 11: 0, 12: 0, 1: 0, 2: 0, 3: 0, 4: 0, 5: 0, 6: 0}, "h1": 5165000, "h2": 0, "annual": 5165000},
    "原":   {"monthly": {7: 4800000, 8: 4550000, 9: 4800000, 10: 4600000, 11: 4800000, 12: 4600000, 1: 3800000, 2: 3750000, 3: 3800000, 4: 3500000, 5: 3500000, 6: 3500000}, "h1": 28150000, "h2": 21850000, "annual": 50000000},
    "鈴木": {"monthly": {7: 12190000, 8: 12200000, 9: 12200000, 10: 12200000, 11: 12200000, 12: 12200000, 1: 10000000, 2: 10000000, 3: 9000000, 4: 9000000, 5: 8810000, 6: 10000000}, "h1": 73190000, "h2": 56810000, "annual": 130000000},
    "穴山": {"monthly": {7: 5000000, 8: 7000000, 9: 7000000, 10: 8000000, 11: 8000000, 12: 40000000, 1: 12000000, 2: 13000000, 3: 15000000, 4: 5000000, 5: 5000000, 6: 5000000}, "h1": 75000000, "h2": 55000000, "annual": 130000000},
    "川村": {"monthly": {7: 12190000, 8: 12200000, 9: 12200000, 10: 12200000, 11: 12200000, 12: 12200000, 1: 10000000, 2: 10000000, 3: 9000000, 4: 9000000, 5: 8810000, 6: 10000000}, "h1": 73190000, "h2": 56810000, "annual": 130000000},
    "田中": {"monthly": {7: 5165000, 8: 5000000, 9: 5200000, 10: 5200000, 11: 5200000, 12: 5200000, 1: 4235000, 2: 4200000, 3: 4200000, 4: 3800000, 5: 3800000, 6: 3800000}, "h1": 30965000, "h2": 24035000, "annual": 55000000},
    "安部": {"monthly": {m: 0 for m in range(1, 13)}, "h1": 0, "h2": 0, "annual": 0},
    "萩原": {"monthly": {m: 0 for m in range(1, 13)}, "h1": 0, "h2": 0, "annual": 0}
}

ORDER_CSV_COLUMNS = [
    '受注ID', '顧客ID', '要望等', '注文者お名前(姓)', '注文者お名前(名)', '注文者お名前(フリガナ・姓)', '注文者お名前(フリガナ名)',
    '注文者会社名', '注文者メールアドレス', '注文者メールアドレスタイプ', '注文者電話番号', '注文者FAX', '注文者国', '注文者郵便番号',
    '注文者都道府県', '注文者住所1', '注文者住所2', '注文者性別ID', '注文者性別', '注文者生年月日', '注文者職種ID', '注文者職種',
    '商品合計', '調整額', '調整額備考', 'うち割引額１', 'うち割引額２', 'うち割引額３', 'うちその他調整額', '合計額', '送料（税別）',
    '手数料（税別）', 'クーポンコード', 'クーポン名', 'クーポン利用', 'ポイント利用', '請求額', '消費税額', '加算ポイント', '入金額',
    '配送コスト', 'EC種類', 'EC種類名', 'EC注文番号', 'EC注文番号枝番', 'EC店舗ID', 'EC店舗名', '販売ルートID', '受注ステータス',
    '支払方法ID', '決済ステータス', '配送方法ID', '配送区分', '申込番号種類', '注文日時', '出荷予定日', '出荷日', 'ひとことメモ',
    '伝達事項', '直接ＡＤコード', '初回ＡＤコード', '再計算フラグ', '送付先コード', '送付先お名前(姓)', '送付先お名前(名)',
    '送付先お名前(フリガナ・姓)', '送付先お名前(フリガナ名)', '送付先会社名', '送付先電話番号', '送付先FAX', '送付先国',
    '送付先郵便番号', '送付先都道府県', '送付先住所1', '送付先住所2', 'お届け日', 'お届け時間', 'マーク１', 'マーク２', 'マーク３',
    '出力フラグ(受注伝票)', '出力フラグ(お手紙)', '出力フラグ(納品書)', '出力フラグ(領収書)', '出力フラグ(請求書)', '出力フラグ(EMS)',
    '出力フラグ(送り状)', '出力済(受注伝票)', '出力済(お手紙)', '出力済(納品書)', '出力済(領収書)', '出力済(請求書)', '出力済(EMS)',
    '出力済(送り状)', '出力済(後払い取引登録)', '出力済(後払い出荷登録)', '出力済(配送指示伝票)', '出力済(モール連携出荷)',
    '出力済(モール連携処理済)', 'メール送信済(サンクス)', 'メール送信済(入金確認)', 'メール送信済(発送完了)', 'メール送信済(フォロー)',
    'メール送信済(その他１)', 'メール送信済(その他２)', 'ご購入明細商品コード', 'ご購入明細商品名', 'ご購入明細消費税区分',
    'ご購入明細規格名(横軸選択肢)', 'ご購入明細規格名(縦軸選択肢)', 'ご購入明細選択肢名', 'ご購入明細消費税課税規則', 'ご購入明細送料区分',
    'ご購入明細手数料区分', 'ご購入明細税抜単価', 'ご購入明細税込単価', 'ご購入明細消費税', 'ご購入明細消費税率', 'ご購入明細数量',
    'ご購入明細金額', 'ご購入明細定期区分', 'ご購入明細標準販売価格', 'ご購入明細JANコード', 'ご購入明細ロットNO', '消費税計算区分'
]

LOTMEMBER_CSV_COLUMNS = [
    'コントロールカラム', '商品規格ID', '商品ID', '商品管理番号', '商品番号',
    'グループ連番', '加盟店コード', '連番', 'ロット数', 'ロット単価'
]

ALLOWED_EXACT_RANKS = [
    "1.1倍", "1.15倍", "1.2倍", "1.25倍", "1.3倍", "1.35倍",
    "1.4倍", "1.45倍", "1.5倍", "0.9倍", "1.05倍", "1.07倍"
]

def generate_excel_bytes(data_df, title_str):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        data_df.to_excel(writer, index=False, sheet_name=str(title_str)[:30])
    return output.getvalue()

def clean_cust_name(val):
    if pd.isna(val) or not isinstance(val, str) or not val.strip(): return ""
    raw = str(val).strip()
    if raw.lower() == "nan": return ""
    norm = re.sub(r'[\s\u3000\xa0]+', '', raw)
    if "アマゾン" in norm:
        if "セラー" in norm: return "アマゾンセラー"
        elif "ホーム" in norm or "home" in norm.lower(): return "アマゾンベンダー（ホーム）"
        elif "アクセ" in norm or "acc" in norm.lower(): return "アマゾンベンダー（アクセサリー）"
        elif "ベンダー" in norm: return "アマゾンベンダー（ホーム）"
        else: return "アマゾンセラー"
    return norm

def clean_num(val):
    if pd.isna(val): return 0.0
    s = re.sub(r'[^\d.-]', '', str(val))
    try: return float(s) if s else 0.0
    except Exception: return 0.0

def clean_code_str(val):
    if pd.isna(val): return ""
    s = str(val).replace("_x000D_", "").strip()
    s = re.sub(r'\.0$', '', s).strip()
    return "" if s.lower() in ["nan", "none"] else s

def convert_staff_name(val):
    if pd.isna(val): return ""
    v = str(val).strip().replace(".0", "")
    if v.lower() == "nan": return ""
    return STAFF_MAP.get(v, v)

def normalize_target_staff_name(raw_name):
    s = str(raw_name).strip()
    for k in ["小泉", "水原", "西山", "松平", "河端", "鈴木", "穴山", "川村", "田中", "安部", "萩原"]:
        if k in s: return k
    if "社長" in s or s == "原": return "原"
    return s

def clean_product_name(val):
    if pd.isna(val) or not isinstance(val, str) or not val.strip(): return ""
    if "】" in val: val = val.split("】", 1)[-1]
    return re.sub(r'[\(（]\s*[\d０-９]{7,16}\s*[\)）]', '', val).strip()

def find_file_multidir(filename_pattern, exclude_substr=None):
    search_dirs = [GDRIVE_SYNC_DIR, BASE_DIR, OTHER_DATA_DIR, BASE_SHARE]
    best_file, best_mtime = None, 0
    for d in search_dirs:
        if os.path.exists(d):
            fl = glob.glob(os.path.join(d, filename_pattern))
            if not fl:
                fl = glob.glob(os.path.join(d, "**", filename_pattern), recursive=True)
            if exclude_substr and fl:
                fl = [f for f in fl if exclude_substr not in os.path.basename(f)]
            for f in fl:
                mt = os.path.getmtime(f)
                if mt > best_mtime:
                    best_file, best_mtime = f, mt
    return best_file, best_mtime

# ========================================================
# ★ ☁️ Googleドライブ 双方向・高速差分同期エンジン
# ========================================================
@st.cache_data(show_spinner=False)
def _get_gdrive_creds_info():
    try:
        folder_id = st.secrets.get("folder_id", "").strip()
        raw_json = st.secrets.get("gcp_service_account_json", "")
        if not folder_id:
            return None, None
        if raw_json:
            info = json.loads(raw_json, strict=False)
        elif "gcp_service_account" in st.secrets:
            info = dict(st.secrets["gcp_service_account"])
        else:
            return None, None
        return info, folder_id
    except Exception:
        return None, None

def get_gdrive_service_and_folder():
    info, folder_id = _get_gdrive_creds_info()
    if not info or not folder_id:
        return None, None
    try:
        from google.oauth2 import service_account
        from googleapiclient.discovery import build
        creds = service_account.Credentials.from_service_account_info(
            info, scopes=["https://www.googleapis.com/auth/drive"]
        )
        service = build("drive", "v3", credentials=creds, cache_discovery=False)
        return service, folder_id
    except Exception:
        return None, None

def sync_all_from_gdrive(force=False):
    last_sync = st.session_state.get("_last_gdrive_sync_ts", 0)
    if not force and (time.time() - last_sync < 30):
        return True

    service, folder_id = get_gdrive_service_and_folder()
    if not service or not folder_id:
        st.session_state["_gdrive_diag"] = "❌ Googleドライブ未接続（st.secrets の設定を確認してください）"
        return False

    try:
        from googleapiclient.http import MediaIoBaseDownload
        os.makedirs(GDRIVE_SYNC_DIR, exist_ok=True)

        meta_path = os.path.join(GDRIVE_SYNC_DIR, "_sync_meta.json")
        local_meta = {}
        if os.path.exists(meta_path):
            try:
                with open(meta_path, "r", encoding="utf-8") as f:
                    local_meta = json.load(f)
            except Exception:
                local_meta = {}

        def _download_single(fid, target_path, copy_to_base=False, is_gsheet=False):
            if is_gsheet:
                req = service.files().export_media(fileId=fid, mimeType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            else:
                req = service.files().get_media(fileId=fid, supportsAllDrives=True)
            fh = io.BytesIO()
            downloader = MediaIoBaseDownload(fh, req)
            done = False
            while not done:
                _, done = downloader.next_chunk()
            data_bytes = fh.getvalue()
            with open(target_path, "wb") as out_f:
                out_f.write(data_bytes)
            if copy_to_base:
                try:
                    with open(os.path.join(BASE_DIR, os.path.basename(target_path)), "wb") as out_b:
                        out_b.write(data_bytes)
                except Exception:
                    pass

        # メインフォルダ直下のファイル一覧を取得
        q = f"'{folder_id}' in parents and trashed = false"
        res = service.files().list(
            q=q,
            fields="files(id, name, mimeType, modifiedTime, shortcutDetails)",
            pageSize=300,
            supportsAllDrives=True,
            includeItemsFromAllDrives=True
        ).execute()
        items = res.get("files", [])

        # ローカルに残っている古い yayoi_*.csv のうち、今ドライブ直下に無い（『済』に移動された）ものを掃除
        current_drive_names = {it.get("name", "") for it in items}
        for old_f in glob.glob(os.path.join(GDRIVE_SYNC_DIR, "yayoi_*.csv")):
            if os.path.basename(old_f) not in current_drive_names:
                try: os.remove(old_f)
                except Exception: pass

        downloaded_files = []

        # ① sales_cache.parquet の同期
        for item in items:
            if item.get("name") == "sales_cache.parquet":
                fid, mtime = item["id"], item.get("modifiedTime", "")
                t_path = os.path.join(GDRIVE_SYNC_DIR, "sales_cache.parquet")
                if force or not os.path.exists(t_path) or local_meta.get(fid) != mtime:
                    _download_single(fid, t_path, copy_to_base=False)
                    local_meta[fid] = mtime
                    downloaded_files.append("sales_cache.parquet")
                break

        # ② ドライブ直下の売上CSV・マスタ類（ショートカット含む）を同期
        ignore_prefixes = ("app", "requirements", "run_all", "silent_start", "item_class_lotmember")
        drive_csv_names = []
        master_detected_name = "未検出"

        for item in items:
            mtype = item.get("mimeType", "")
            if mtype == "application/vnd.google-apps.folder":
                continue

            fname = item.get("name", "")
            fid = item["id"]
            mtime = item.get("modifiedTime", "")

            # ★重要：ショートカットの場合は、リンク先（本体ファイル）のIDと「本体の最新更新日時」を取得する！
            if mtype == "application/vnd.google-apps.shortcut":
                sc_info = item.get("shortcutDetails", {})
                target_id = sc_info.get("targetId")
                if target_id:
                    fid = target_id
                    mtype = sc_info.get("targetMimeType", "")
                    try:
                        t_meta = service.files().get(
                            fileId=fid,
                            fields="name, mimeType, modifiedTime",
                            supportsAllDrives=True
                        ).execute()
                        mtime = t_meta.get("modifiedTime", mtime)
                        mtype = t_meta.get("mimeType", mtype)
                    except Exception:
                        pass

            if fname == "sales_cache.parquet":
                continue

            is_gsheet = (mtype == "application/vnd.google-apps.spreadsheet")
            if is_gsheet:
                if any(k in fname.lower() for k in ["order_r", "マスター", "在庫確保", "発注検討"]):
                    if not fname.lower().endswith(".xlsx"):
                        fname = fname + ".xlsx"
                else:
                    continue

            # ショートカット名に拡張子(.xlsx)が付いていない場合も自動で補完
            if ("order_r" in fname.lower() or "マスター" in fname) and not fname.lower().endswith((".xlsx", ".xls")):
                fname = fname + ".xlsx"

            if fname.lower().startswith(ignore_prefixes):
                continue
            if fname.lower().startswith("order_") and "order_r" not in fname.lower() and "マスター" not in fname:
                continue
            if not fname.lower().endswith((".csv", ".xlsx", ".xls", ".json", ".png")):
                continue

            if "order_r" in fname.lower() or "マスター" in fname:
                master_detected_name = fname

            target_path = os.path.join(GDRIVE_SYNC_DIR, fname)
            is_inventory_or_master = ("ロジザード" in fname or "在庫" in fname or "ランク" in fname or fname.lower().endswith((".xlsx", ".xls", ".json", ".png")))

            if fname.lower().endswith(".csv") and not is_inventory_or_master:
                drive_csv_names.append(fname)
                if force or not os.path.exists(target_path) or local_meta.get(fid) != mtime:
                    _download_single(fid, target_path, copy_to_base=False, is_gsheet=False)
                    local_meta[fid] = mtime
                    downloaded_files.append(fname)
            else:
                if force or not os.path.exists(target_path) or local_meta.get(fid) != mtime:
                    _download_single(fid, target_path, copy_to_base=True, is_gsheet=is_gsheet)
                    local_meta[fid] = mtime
                    downloaded_files.append(fname)

        with open(meta_path, "w", encoding="utf-8") as f:
            json.dump(local_meta, f, ensure_ascii=False, indent=2)

    st.session_state["_gdrive_diag"] = (
            f"☁️ 同期OK（{datetime.now().strftime('%H:%M:%S')}）｜"
            f"マスターExcel: {master_detected_name} ｜ "
            f"直下の売上CSV: {sorted(drive_csv_names)} ｜ "
            f"今回DL: {downloaded_files if downloaded_files else '変更なし'}"
        )
        st.session_state["_last_gdrive_sync_ts"] = time.time()
        if downloaded_files:
            st.cache_data.clear()
        return True
    except Exception as e:
        st.session_state["_gdrive_diag"] = f"❌ ドライブ同期エラー: {str(e)}"
        return False 　

def upload_file_to_gdrive(local_path, drive_filename=None):
    service, folder_id = get_gdrive_service_and_folder()
    if not service or not folder_id or not os.path.exists(local_path):
       return False
    try:
        from googleapiclient.http import MediaFileUpload
        fname = drive_filename or os.path.basename(local_path)
        q = f"'{folder_id}' in parents and name = '{fname}' and trashed = false"
        res = service.files().list(q=q, fields="files(id, name)").execute()
        existing = res.get("files", [])
        media = MediaFileUpload(local_path, resumable=True)
        if existing:
            up_res = service.files().update(fileId=existing[0]["id"], media_body=media, fields="id, modifiedTime").execute()
        else:
            meta = {"name": fname, "parents": [folder_id]}
            up_res = service.files().create(body=meta, media_body=media, fields="id, modifiedTime").execute()

        meta_path = os.path.join(GDRIVE_SYNC_DIR, "_sync_meta.json")
        if os.path.exists(meta_path) and up_res.get("id"):
            try:
                with open(meta_path, "r", encoding="utf-8") as f:
                    local_meta = json.load(f)
                local_meta[up_res["id"]] = up_res.get("modifiedTime", "")
                with open(meta_path, "w", encoding="utf-8") as f:
                    json.dump(local_meta, f, ensure_ascii=False, indent=2)
            except Exception:
                pass
        return True
    except Exception:
        return False

def save_parquet_and_sync(target_df, local_path, drive_filename="sales_cache.parquet"):
    try:
        target_df.to_parquet(local_path, index=False)
        sync_copy = os.path.join(GDRIVE_SYNC_DIR, drive_filename)
        target_df.to_parquet(sync_copy, index=False)
        upload_file_to_gdrive(local_path, drive_filename)
    except Exception:
        pass

_gdrive_ok = sync_all_from_gdrive(force=False)

# ========================================================
# ★ 在庫・確保数・発注検討表 データ読込エンジン
# ========================================================
def get_inventory_files_info():
    p_ao, m_ao = find_file_multidir("*ロジザード在庫*青山*.csv")
    p_ng, m_ng = find_file_multidir("*ロジザード在庫*名古屋*.csv")
    p_kk, m_kk = find_file_multidir("*在庫確保*.xls*")
    p_hc, m_hc = find_file_multidir("*発注検討表*.xls*")
    return (p_ao, m_ao), (p_ng, m_ng), (p_kk, m_kk), (p_hc, m_hc)

@st.cache_data(show_spinner=False)
def load_inventory_and_order_data(ao_info, ng_info, kk_info, hc_info):
    logi_map = {}
    for p, _ in [ao_info, ng_info]:
        if p and os.path.exists(p):
            for enc in ["cp932", "shift_jis", "utf-8-sig", "utf-8"]:
                try:
                    ldf = pd.read_csv(p, encoding=enc, low_memory=False)
                    for _, r in ldf.iterrows():
                        sku = str(r.get("商品ID", "")).strip()
                        if not sku or sku in ["nan", "None"]: continue
                        wh = str(r.get("倉庫名", "")).strip()
                        wh_short = "青山" if ("三重" in wh or "青山" in wh) else ("名古屋" if "名古屋" in wh else wh)
                        q_id = int(clean_num(r.get("品質区分ID", 0)))
                        qty = int(clean_num(r.get("在庫数(引当数を含む)", 0)))
                        avail = int(clean_num(r.get("出荷可能数", 0)))
                        inc = int(clean_num(r.get("入荷予定数", 0)))
                        if sku not in logi_map:
                            logi_map[sku] = {
                                "name": str(r.get("商品名", "")).strip(),
                                "wh_set": set(),
                                "good_qty": 0, "avail_qty": 0, "uninspected_qty": 0,
                                "reserved_logi": 0, "defective_qty": 0, "incoming_qty": 0,
                                "aoyama_good": 0, "nagoya_good": 0
                            }
                        if wh_short: logi_map[sku]["wh_set"].add(wh_short)
                        logi_map[sku]["incoming_qty"] += inc
                        if q_id == 1:
                            logi_map[sku]["good_qty"] += qty
                            logi_map[sku]["avail_qty"] += avail
                            if wh_short == "青山": logi_map[sku]["aoyama_good"] += qty
                            elif wh_short == "名古屋": logi_map[sku]["nagoya_good"] += qty
                        elif q_id == 2:
                            logi_map[sku]["uninspected_qty"] += qty
                        elif q_id == 13:
                            logi_map[sku]["reserved_logi"] += qty
                        elif q_id in [3, 4, 7]:
                            logi_map[sku]["defective_qty"] += qty
                    break
                except Exception: continue

    kakuho_sum_map = {}
    kakuho_details_list = []
    p_kk, _ = kk_info
    if p_kk and os.path.exists(p_kk):
        try:
            raw_k = pd.read_excel(p_kk, header=None)
            h_idx = 3
            for i in range(min(15, len(raw_k))):
                row_vals = [str(x).strip() for x in raw_k.iloc[i].values if pd.notna(x)]
                if "商品コード" in row_vals and ("受注残数" in row_vals or "得意先名" in row_vals):
                    h_idx = i
                    break
            kdf = pd.read_excel(p_kk, header=h_idx)
            for _, r in kdf.iterrows():
                sku = str(r.get("商品コード", "")).strip()
                if not sku or sku in ["nan", "None", ""]: continue
                rem_qty = int(clean_num(r.get("受注残数", r.get("数量", 0))))
                if rem_qty == 0 and int(clean_num(r.get("数量", 0))) == 0: continue
                kakuho_sum_map[sku] = kakuho_sum_map.get(sku, 0) + rem_qty
                dt_v = r.get("受注日")
                dt_str = pd.to_datetime(dt_v).strftime("%Y/%m/%d") if pd.notna(dt_v) else ""
                due_v = r.get("納入期日")
                due_str = pd.to_datetime(due_v).strftime("%Y/%m/%d") if pd.notna(due_v) else ""
                kakuho_details_list.append({
                    "商品コード": sku,
                    "商品名": str(r.get("商品名/摘要", "")).strip() if pd.notna(r.get("商品名/摘要")) else "",
                    "受注日": dt_str,
                    "受注番号": clean_code_str(r.get("受注番号", "")),
                    "得意先名": str(r.get("得意先名", "")).strip() if pd.notna(r.get("得意先名")) else "",
                    "担当者": str(r.get("担当者名", "")).strip() if pd.notna(r.get("担当者名")) else "",
                    "倉庫": str(r.get("倉庫名", "")).strip() if pd.notna(r.get("倉庫名")) else "",
                    "確保残数": rem_qty,
                    "納入期日": due_str,
                    "備考": str(r.get("備考", "")).strip() if pd.notna(r.get("備考")) else ""
                })
        except Exception: pass

    hacchu_map = {}
    p_hc, _ = hc_info
    if p_hc and os.path.exists(p_hc):
        try:
            hdf = pd.read_excel(p_hc)
            c_o = next((c for c in hdf.columns if "発注点" in str(c) and "5ヶ月" in str(c)), None)
            c_r = next((c for c in hdf.columns if "在庫数" in str(c) and "ロジザード" in str(c)), None)
            c_s = next((c for c in hdf.columns if "入荷予定" in str(c) and "日" not in str(c)), None)
            c_u = next((c for c in hdf.columns if "入荷予定日" in str(c)), None)
            c_v = next((c for c in hdf.columns if "在庫確保" in str(c) and "ロジザード" in str(c)), None)
            c_w = next((c for c in hdf.columns if "在庫確保" in str(c) and "弥生" in str(c)), None)
            c_y = next((c for c in hdf.columns if "丸藤" in str(c)), None)
            c_ac = next((c for c in hdf.columns if "ﾘﾝﾍﾞﾙ" in str(c) or "リンベル" in str(c)), None)
            c_ad = next((c for c in hdf.columns if "前回決定事項" in str(c)), None)
            c_ae = next((c for c in hdf.columns if str(c).strip() == "決定事項"), None)
            c_l = next((c for c in hdf.columns if "77期" in str(c)), None)

            for _, r in hdf.iterrows():
                sku = str(r.get("品番", "")).strip()
                if not sku or sku in ["nan", "None"]: continue
                hacchu_map[sku] = {
                    "center": str(r.get("商品センター", "")).strip() if pd.notna(r.get("商品センター")) else "",
                    "category": str(r.get("商品区分", "")).strip() if pd.notna(r.get("商品区分")) else "",
                    "color": str(r.get("カラー", "")).strip() if pd.notna(r.get("カラー")) else "",
                    "discontinued": str(r.get("処分品", "")).strip() if pd.notna(r.get("処分品")) else "",
                    "order_point": round(clean_num(r.get(c_o, 0)), 1) if c_o else 0.0,
                    "sheet_logi_stock": int(clean_num(r.get(c_r, 0))) if c_r else 0,
                    "sheet_incoming": int(clean_num(r.get(c_s, 0))) if c_s else 0,
                    "incoming_date": str(r.get(c_u, "")).strip() if (c_u and pd.notna(r.get(c_u))) else "",
                    "sheet_logi_res": int(clean_num(r.get(c_v, 0))) if c_v else 0,
                    "sheet_yayoi_res": int(clean_num(r.get(c_w, 0))) if c_w else 0,
                    "marufuji_res": int(clean_num(r.get(c_y, 0))) if c_y else 0,
                    "ringbell_res": int(clean_num(r.get(c_ac, 0))) if c_ac else 0,
                    "prev_note": str(r.get(c_ad, "")).strip() if (c_ad and pd.notna(r.get(c_ad))) else "",
                    "decision": str(r.get(c_ae, "")).strip() if (c_ae and pd.notna(r.get(c_ae))) else "",
                    "term77_qty": int(clean_num(r.get(c_l, 0))) if c_l else 0
                }
        except Exception: pass

    return logi_map, kakuho_sum_map, pd.DataFrame(kakuho_details_list), hacchu_map

def get_target_file_info():
    return find_file_multidir("*今期目標*.xls*")

@st.cache_data(show_spinner=False)
def load_staff_target_master(file_path, file_mtime):
    targets = {k: dict(v) for k, v in DEFAULT_STAFF_TARGETS.items()}
    if file_path and os.path.exists(file_path):
        try:
            wb = openpyxl.load_workbook(file_path, data_only=True)
            ws = wb.active
            col_to_month = {3: 7, 4: 8, 5: 9, 6: 10, 7: 11, 8: 12, 10: 1, 11: 2, 12: 3, 13: 4, 14: 5, 15: 6}
            curr_person = None
            for r in range(1, ws.max_row + 1):
                p = ws.cell(row=r, column=1).value
                if p is not None and str(p).strip() != "":
                    curr_person = normalize_target_staff_name(str(p).strip())
                metric = str(ws.cell(row=r, column=2).value or "").strip()
                if curr_person and metric == "今期目標":
                    m_dict = {}
                    for c_idx, m_num in col_to_month.items():
                        v = ws.cell(row=r, column=c_idx).value
                        try: m_dict[m_num] = int(float(v) * 1000) if v is not None else 0
                        except Exception: m_dict[m_num] = 0
                    v_h1 = ws.cell(row=r, column=9).value
                    v_h2 = ws.cell(row=r, column=16).value
                    v_fy = ws.cell(row=r, column=17).value
                    try: h1_val = int(float(v_h1) * 1000) if v_h1 is not None else sum(m_dict[m] for m in [7,8,9,10,11,12])
                    except Exception: h1_val = sum(m_dict[m] for m in [7,8,9,10,11,12])
                    try: h2_val = int(float(v_h2) * 1000) if v_h2 is not None else sum(m_dict[m] for m in [1,2,3,4,5,6])
                    except Exception: h2_val = sum(m_dict[m] for m in [1,2,3,4,5,6])
                    try: fy_val = int(float(v_fy) * 1000) if v_fy is not None else (h1_val + h2_val)
                    except Exception: fy_val = h1_val + h2_val
                    targets[curr_person] = {"monthly": m_dict, "h1": h1_val, "h2": h2_val, "annual": fy_val}
        except Exception:
            pass
    return targets

# ========================================================
# ★ 👑 order_r（マスター）.xlsx 一括高速読込エンジン（重複完全排除・決定版）
# ========================================================
def get_order_r_master_info():
    for pat in ["*order_r*マスター*.xls*", "*order_r*.xls*", "*マスター*.xls*"]:
        p, m = find_file_multidir(pat)
        if p:
            return p, m
    return None, 0

def _extract_jan_13digits(*texts):
    for val in texts:
        if pd.isna(val):
            continue
        s = str(val).strip()
        if not s or s.lower() == "nan":
            continue
        try:
            f_val = float(s)
            if f_val >= 1e11:
                i_str = str(int(round(f_val)))
                if len(i_str) in (12, 13):
                    return i_str
        except Exception:
            pass
        m = re.search(r'(?<!\d)(\d{13}|\d{12})(?!\d)', s)
        if m:
            return m.group(1)
    return ""

@st.cache_data(show_spinner=False)
def _load_order_r_all_sheets(master_path, master_mtime):
    if not master_path or not os.path.exists(master_path):
        return None
    try:
        xls = pd.ExcelFile(master_path)
        s_names = xls.sheet_names
        res = {
            "cost_dict": {}, "cust_sku_map": {}, "default_min_map": {},
            "sku_meta": {}, "grp_map": {}, "rank_df": pd.DataFrame(),
            "jan_dict": {}, "full_title_dict": {}, "raw_title_dict": {}, "std_price_dict": {},
            "customer_df": pd.DataFrame()
        }

        if "新原価" in s_names:
            cdf = pd.read_excel(xls, sheet_name="新原価")
            for _, r in cdf.iterrows():
                sk = clean_code_str(r.iloc[0]) if len(r) > 0 else ""
                if not sk: continue
                co = clean_num(r.iloc[1]) if len(r) > 1 else 0.0
                if co > 0: res["cost_dict"][sk] = co
                if len(r) > 3:
                    min_p = int(clean_num(r.iloc[3]))
                    if 50 <= min_p < 90000:
                        res["default_min_map"][sk] = min_p

        if "加盟店価格" in s_names:
            mdf = pd.read_excel(xls, sheet_name="加盟店価格")
            if len(mdf.columns) >= 14:
                for _, r in mdf.iterrows():
                    sku_mgmt = clean_code_str(r.iloc[3]) if len(r) > 3 else ""
                    sku_num = clean_code_str(r.iloc[4]) if len(r) > 4 else ""
                    skus = [s for s in [sku_mgmt, sku_num] if s]
                    if not skus: continue
                    j_price = int(clean_num(r.iloc[9])) if len(r) > 9 else 0
                    n_min_price = int(clean_num(r.iloc[13])) if len(r) > 13 else 0
                    m_comp = str(r.iloc[12]).strip() if len(r) > 12 and pd.notna(r.iloc[12]) else ""
                    for s in skus:
                        if 50 <= n_min_price < 90000:
                            res["default_min_map"][s] = n_min_price
                        if m_comp and m_comp not in ["nan", "None", "#N/A", ""]:
                            norm_comp = re.sub(r'[\s\u3000\xa0]+', '', m_comp)
                            if 50 <= j_price < 90000:
                                res["cust_sku_map"][(m_comp, s)] = j_price
                                res["cust_sku_map"][(norm_comp, s)] = j_price

        if "加盟店価格データ" in s_names:
            ldf = pd.read_excel(xls, sheet_name="加盟店価格データ")
            for _, r in ldf.drop_duplicates(subset=['商品番号']).iterrows():
                sk = clean_code_str(r.get('商品番号', ''))
                if not sk: continue
                res["sku_meta"][sk] = {
                    '商品規格ID': int(clean_num(r.get('商品規格ID', 0))) if pd.notna(r.get('商品規格ID')) else "",
                    '商品ID': int(clean_num(r.get('商品ID', 0))) if pd.notna(r.get('商品ID')) else "",
                    '商品管理番号': clean_code_str(r.get('商品管理番号', sk)) or sk,
                    '商品番号': sk
                }

        if "加盟店グループ管理" in s_names:
            gdf = pd.read_excel(xls, sheet_name="加盟店グループ管理")
            for _, r in gdf.iterrows():
                c_name = str(r.get('加盟店名', '')).strip()
                if not c_name or c_name == 'nan': continue
                g_no = clean_code_str(r.get('グループ番号', ''))
                m_cd = clean_code_str(r.get('加盟店コード', ''))
                res["grp_map"][c_name] = {'group_no': g_no, 'member_code': m_cd}
                norm_c = re.sub(r'[\s\u3000\xa0]+', '', c_name)
                if norm_c not in res["grp_map"]:
                    res["grp_map"][norm_c] = {'group_no': g_no, 'member_code': m_cd}

        if "ランク価格" in s_names:
            res["rank_df"] = pd.read_excel(xls, sheet_name="ランク価格")

        # item シート（B列[1]:商品番号 / F列[5]:】〜（の商品名 / O列[14]:カタログID[JAN] / Q列[16]:原価）
        if "item" in s_names:
            idf = pd.read_excel(xls, sheet_name="item")
            for _, r in idf.iterrows():
                sk = clean_code_str(r.iloc[1]) if len(r) > 1 else ""
                if not sk:
                    sk = clean_code_str(r.iloc[4]) if len(r) > 4 else ""
                if not sk:
                    continue
                raw_nm = str(r.iloc[5]).strip() if len(r) > 5 and pd.notna(r.iloc[5]) else ""
                cat_id = r.iloc[14] if len(r) > 14 else None

                jan_str = _extract_jan_13digits(cat_id, raw_nm)
                if jan_str:
                    res["jan_dict"][sk] = jan_str

                if raw_nm and raw_nm != "nan":
                    res["raw_title_dict"][sk] = raw_nm
                    nm_clean = raw_nm.split("】", 1)[-1] if "】" in raw_nm else raw_nm
                    nm_clean = re.sub(r'[\(（]\s*[\d０-９]{7,16}\s*[\)）]\s*$', '', nm_clean).strip()
                    res["full_title_dict"][sk] = nm_clean

                std_p = int(clean_num(r.iloc[9])) if len(r) > 9 else 0
                if std_p > 0: res["std_price_dict"][sk] = std_p
                sell_p = int(clean_num(r.iloc[8])) if len(r) > 8 else 0
                if sk not in res["default_min_map"] and 50 <= sell_p < 90000:
                    res["default_min_map"][sk] = sell_p
                # ★修正：itemシートのQ列（iloc[16]）の原価を最優先で登録！
                c_val = clean_num(r.iloc[16]) if len(r) > 16 else 0.0
                if c_val > 0:
                    res["cost_dict"][sk] = c_val

        # select シート（A列[0]:商品番号 / I列[8]:】以降の商品名 / Q列[16]:カタログID[JAN] / J列[9]:原価）
        if "select" in s_names:
            sdf = pd.read_excel(xls, sheet_name="select")
            for _, r in sdf.iterrows():
                sk = clean_code_str(r.iloc[0]) if len(r) > 0 else ""
                if not sk:
                    sk = clean_code_str(r.iloc[7]) if len(r) > 7 else ""
                if not sk:
                    continue
                raw_nm = str(r.iloc[8]).strip() if len(r) > 8 and pd.notna(r.iloc[8]) else ""
                cat_id = r.iloc[16] if len(r) > 16 else None
                opt_d = str(r.iloc[3]).strip() if len(r) > 3 and pd.notna(r.iloc[3]) else ""

                jan_str = _extract_jan_13digits(cat_id, opt_d, raw_nm)
                if jan_str:
                    res["jan_dict"][sk] = jan_str

                if raw_nm and raw_nm != "nan":
                    res["raw_title_dict"][sk] = raw_nm
                    nm_clean = raw_nm.split("】", 1)[-1].strip() if "】" in raw_nm else raw_nm.strip()
                    nm_clean = re.sub(r'[\(（]\s*[\d０-９]{7,16}\s*[\)）]\s*$', '', nm_clean).strip()
                    res["full_title_dict"][sk] = nm_clean

                std_p = int(clean_num(r.iloc[10])) if len(r) > 10 else 0
                if std_p > 0: res["std_price_dict"][sk] = std_p
                sell_p = int(clean_num(r.iloc[11])) if len(r) > 11 else 0
                if sk not in res["default_min_map"] and 50 <= sell_p < 90000:
                    res["default_min_map"][sk] = sell_p
                # ★修正：selectシートのJ列（iloc[9]）の原価を最優先で登録！
                c_val = clean_num(r.iloc[9]) if len(r) > 9 else 0.0
                if c_val > 0:
                    res["cost_dict"][sk] = c_val

        # select シート（A列[0]:商品番号 / I列[8]:】以降の商品名 / Q列[16]:カタログID[JAN]）
        if "select" in s_names:
            sdf = pd.read_excel(xls, sheet_name="select")
            for _, r in sdf.iterrows():
                sk = clean_code_str(r.iloc[0]) if len(r) > 0 else ""
                if not sk:
                    sk = clean_code_str(r.iloc[7]) if len(r) > 7 else ""
                if not sk:
                    continue
                raw_nm = str(r.iloc[8]).strip() if len(r) > 8 and pd.notna(r.iloc[8]) else ""
                cat_id = r.iloc[16] if len(r) > 16 else None
                opt_d = str(r.iloc[3]).strip() if len(r) > 3 and pd.notna(r.iloc[3]) else ""

                jan_str = _extract_jan_13digits(cat_id, opt_d, raw_nm)
                if jan_str:
                    res["jan_dict"][sk] = jan_str

                if raw_nm and raw_nm != "nan":
                    res["raw_title_dict"][sk] = raw_nm
                    nm_clean = raw_nm.split("】", 1)[-1].strip() if "】" in raw_nm else raw_nm.strip()
                    nm_clean = re.sub(r'[\(（]\s*[\d０-９]{7,16}\s*[\)）]\s*$', '', nm_clean).strip()
                    res["full_title_dict"][sk] = nm_clean

                std_p = int(clean_num(r.iloc[10])) if len(r) > 10 else 0
                if std_p > 0: res["std_price_dict"][sk] = std_p
                sell_p = int(clean_num(r.iloc[11])) if len(r) > 11 else 0
                if sk not in res["default_min_map"] and 50 <= sell_p < 90000:
                    res["default_min_map"][sk] = sell_p
                if sk not in res["cost_dict"] or res["cost_dict"][sk] == 0:
                    c_val = clean_num(r.iloc[9]) if len(r) > 9 else 0.0
                    if c_val > 0: res["cost_dict"][sk] = c_val

        if "注文者会社名" in s_names:
            res["customer_df"] = pd.read_excel(xls, sheet_name="注文者会社名")

        return res
    except Exception:
        return None

def get_cost_master_file_info():
    p, m = get_order_r_master_info()
    if p: return p, m
    return find_file_multidir("*原価*.xls*")

@st.cache_data(show_spinner=False)
def load_cost_master(file_path, file_mtime):
    m_p, m_t = get_order_r_master_info()
    if m_p:
        m_data = _load_order_r_all_sheets(m_p, m_t)
        if m_data and m_data["cost_dict"]:
            return m_data["cost_dict"]
    cost_dict = {}
    if file_path and os.path.exists(file_path):
        try:
            cdf = pd.read_excel(file_path)
            c_sku = [c for c in cdf.columns if any(k in str(c) for k in ['商品管理', 'SKU', 'コード', '品番', '商品番号'])]
            c_cost = [c for c in cdf.columns if '原価' in str(c)]
            sku_col = c_sku[0] if c_sku else cdf.columns[0]
            cost_col = c_cost[0] if c_cost else (cdf.columns[1] if len(cdf.columns) > 1 else None)
            if cost_col:
                for _, r in cdf.iterrows():
                    sk = clean_code_str(r[sku_col])
                    co = clean_num(r[cost_col])
                    if sk: cost_dict[sk] = co
        except Exception: pass
    return cost_dict

def get_member_price_file_info():
    p, m = get_order_r_master_info()
    if p: return p, m
    p, m = find_file_multidir("*加盟店価格*.xls*", exclude_substr="加盟店価格データ")
    if p: return p, m
    return find_file_multidir("*加盟店価格*.xls*")

@st.cache_data(show_spinner=False)
def load_member_price_master_v2(file_path, file_mtime):
    m_p, m_t = get_order_r_master_info()
    if m_p:
        m_data = _load_order_r_all_sheets(m_p, m_t)
        if m_data and (m_data["cust_sku_map"] or m_data["default_min_map"]):
            return m_data["cust_sku_map"], m_data["default_min_map"]
    cust_sku_map, default_min_map = {}, {}
    if file_path and os.path.exists(file_path):
        try:
            mdf = pd.read_excel(file_path)
            if len(mdf.columns) >= 14:
                for _, r in mdf.iterrows():
                    sku_mgmt = clean_code_str(r.iloc[3]) if len(r) > 3 else ""
                    sku_num = clean_code_str(r.iloc[4]) if len(r) > 4 else ""
                    skus = [s for s in [sku_mgmt, sku_num] if s]
                    if not skus: continue
                    j_price = int(clean_num(r.iloc[9])) if len(r) > 9 else 0
                    n_min_price = int(clean_num(r.iloc[13])) if len(r) > 13 else 0
                    m_comp = str(r.iloc[12]).strip() if len(r) > 12 and pd.notna(r.iloc[12]) else ""
                    for s in skus:
                        if 50 <= n_min_price < 90000:
                            default_min_map[s] = n_min_price
                        if m_comp and m_comp not in ["nan", "None", "#N/A", ""]:
                            norm_comp = re.sub(r'[\s\u3000\xa0]+', '', m_comp)
                            if 50 <= j_price < 90000:
                                cust_sku_map[(m_comp, s)] = j_price
                                cust_sku_map[(norm_comp, s)] = j_price
        except Exception: pass
    return cust_sku_map, default_min_map

def get_lotmember_template_file_info():
    p, m = get_order_r_master_info()
    if p: return p, m
    return find_file_multidir("*加盟店価格データ*.xls*")

@st.cache_data(show_spinner=False)
def load_lotmember_sku_master(file_path, file_mtime):
    m_p, m_t = get_order_r_master_info()
    if m_p:
        m_data = _load_order_r_all_sheets(m_p, m_t)
        if m_data and m_data["sku_meta"]:
            return m_data["sku_meta"]
    sku_meta = {}
    if file_path and os.path.exists(file_path):
        try:
            ldf = pd.read_excel(file_path)
            for _, r in ldf.drop_duplicates(subset=['商品番号']).iterrows():
                sk = clean_code_str(r.get('商品番号', ''))
                if not sk: continue
                sku_meta[sk] = {
                    '商品規格ID': int(clean_num(r.get('商品規格ID', 0))) if pd.notna(r.get('商品規格ID')) else "",
                    '商品ID': int(clean_num(r.get('商品ID', 0))) if pd.notna(r.get('商品ID')) else "",
                    '商品管理番号': clean_code_str(r.get('商品管理番号', sk)) or sk,
                    '商品番号': sk
                }
        except Exception: pass
    return sku_meta

def get_member_group_file_info():
    p, m = get_order_r_master_info()
    if p: return p, m
    return find_file_multidir("*加盟店グループ管理*.xls*")

@st.cache_data(show_spinner=False)
def load_member_group_master(file_path, file_mtime):
    m_p, m_t = get_order_r_master_info()
    if m_p:
        m_data = _load_order_r_all_sheets(m_p, m_t)
        if m_data and m_data["grp_map"]:
            return m_data["grp_map"]
    grp_map = {}
    if file_path and os.path.exists(file_path):
        try:
            gdf = pd.read_excel(file_path)
            for _, r in gdf.iterrows():
                c_name = str(r.get('加盟店名', '')).strip()
                if not c_name or c_name == 'nan': continue
                g_no = clean_code_str(r.get('グループ番号', ''))
                m_cd = clean_code_str(r.get('加盟店コード', ''))
                grp_map[c_name] = {'group_no': g_no, 'member_code': m_cd}
                norm_c = re.sub(r'[\s\u3000\xa0]+', '', c_name)
                if norm_c not in grp_map:
                    grp_map[norm_c] = {'group_no': g_no, 'member_code': m_cd}
        except Exception: pass
    return grp_map

def lookup_group_info_by_comp(comp_name, grp_master_dict):
    if not comp_name or not grp_master_dict:
        return {"group_no": "", "member_code": ""}
    c_trim = str(comp_name).strip()
    if c_trim in grp_master_dict:
        return grp_master_dict[c_trim]
    norm_c = re.sub(r'[\s\u3000\xa0]+', '', c_trim)
    if norm_c in grp_master_dict:
        return grp_master_dict[norm_c]
    for k, v in grp_master_dict.items():
        norm_k = re.sub(r'[\s\u3000\xa0]+', '', str(k))
        if norm_c and norm_k and (norm_c in norm_k or norm_k in norm_c):
            return v
    return {"group_no": "", "member_code": ""}

def create_lotmember_csv_bytes(items_list, group_no, member_code, sku_meta_dict, default_ctrl=None):
    out_rows = []
    g_str = clean_code_str(group_no)
    m_str = clean_code_str(member_code)
    for item in items_list:
        sk = str(item.get("商品コード", "")).strip()
        if not sk: continue
        if default_ctrl in ["n", "u"]:
            ctrl = default_ctrl
        else:
            reg_str = str(item.get("登録区分", "更新 (u)"))
            ctrl = "n" if ("n" in reg_str.lower() or "新規" in reg_str) else "u"
        raw_p = item.get("登録単価", item.get("見積単価", 0))
        u_price = int(clean_num(raw_p))
        meta = (sku_meta_dict or {}).get(sk, {'商品規格ID': '', '商品ID': '', '商品管理番号': sk, '商品番号': sk})
        out_rows.append({
            'コントロールカラム': ctrl,
            '商品規格ID': meta['商品規格ID'],
            '商品ID': meta['商品ID'],
            '商品管理番号': meta['商品管理番号'],
            '商品番号': meta['商品番号'],
            'グループ連番': g_str,
            '加盟店コード': "",
            '連番': 1,
            'ロット数': 1,
            'ロット単価': u_price
        })
        out_rows.append({
            'コントロールカラム': ctrl,
            '商品規格ID': meta['商品規格ID'],
            '商品ID': meta['商品ID'],
            '商品管理番号': meta['商品管理番号'],
            '商品番号': meta['商品番号'],
            'グループ連番': g_str,
            '加盟店コード': m_str,
            '連番': "",
            'ロット数': "",
            'ロット単価': ""
        })
    out_df = pd.DataFrame(out_rows, columns=LOTMEMBER_CSV_COLUMNS)
    return out_df.to_csv(index=False, encoding="cp932", lineterminator="\r\n", errors="replace").encode("cp932")

def get_rank_price_file_info():
    p, m = get_order_r_master_info()
    if p: return p, m
    for pat in ["*ランク価格*.xls*", "*ランク価格*.csv", "*ランク*.xls*", "*ランク*.csv"]:
        p, mt = find_file_multidir(pat)
        if p: return p, mt
    return None, 0

@st.cache_data(show_spinner=False)
def load_rank_price_master(file_path, file_mtime):
    m_p, m_t = get_order_r_master_info()
    if m_p:
        m_data = _load_order_r_all_sheets(m_p, m_t)
        if m_data and not m_data["rank_df"].empty:
            return m_data["rank_df"]
    if not file_path or not os.path.exists(file_path): return pd.DataFrame()
    try:
        if file_path.lower().endswith(".csv"):
            for enc in ["utf-8-sig", "cp932", "shift_jis", "utf-8"]:
                try: return pd.read_csv(file_path, encoding=enc, low_memory=False)
                except Exception: continue
        else: return pd.read_excel(file_path)
    except Exception: pass
    return pd.DataFrame()

def get_jan_file_info():
    p, m = get_order_r_master_info()
    if p: return p, m
    return find_file_multidir("*JAN*.xls*")

@st.cache_data(show_spinner=False)
def load_jan_master_v2(file_path, file_mtime):
    jan_dict, full_title_dict, std_price_dict = {}, {}, {}
    m_p, m_t = get_order_r_master_info()
    if m_p:
        m_data = _load_order_r_all_sheets(m_p, m_t)
        if m_data:
            jan_dict.update(m_data["jan_dict"])
            full_title_dict.update(m_data["full_title_dict"])
            std_price_dict.update(m_data["std_price_dict"])
    if not jan_dict:
        j_p, _ = find_file_multidir("*JAN*.xls*")
        if j_p and os.path.exists(j_p):
            try:
                jdf = pd.read_excel(j_p)
                c_sku = [c for c in jdf.columns if any(k in str(c) for k in ['商品番号', '商品コード', 'SKU', '品番'])]
                c_jan = [c for c in jdf.columns if any(k in str(c) for k in ['カタログID', 'JAN', 'JANコード'])]
                c_name = [c for c in jdf.columns if any(k in str(c) for k in ['商品名', '品名'])]
                c_std = [c for c in jdf.columns if any(k in str(c) for k in ['上代', '標準', '定価', '税抜上代'])]
                s_col = c_sku[0] if c_sku else jdf.columns[0]
                j_col = c_jan[0] if c_jan else (jdf.columns[1] if len(jdf.columns) > 1 else None)
                n_col = c_name[0] if c_name else None
                p_col = c_std[0] if c_std else None
                for _, r in jdf.iterrows():
                    sk = clean_code_str(r[s_col])
                    if not sk: continue
                    if j_col and pd.notna(r[j_col]):
                        jan_s = _extract_jan_13digits(r[j_col])
                        if jan_s: jan_dict[sk] = jan_s
                    if n_col and pd.notna(r[n_col]):
                        full_title_dict[sk] = clean_product_name(str(r[n_col]).strip())
                    if p_col and pd.notna(r[p_col]):
                        std_price_dict[sk] = int(clean_num(r[p_col]))
            except Exception: pass
    return jan_dict, full_title_dict, std_price_dict

def get_customer_file_info():
    p, m = get_order_r_master_info()
    if p: return p, m
    return find_file_multidir("*注文者会社名*.xls*")

@st.cache_data(show_spinner=False)
def load_customer_master(file_path, file_mtime):
    m_p, m_t = get_order_r_master_info()
    if m_p:
        m_data = _load_order_r_all_sheets(m_p, m_t)
        if m_data and not m_data["customer_df"].empty:
            return m_data["customer_df"]
    if file_path and os.path.exists(file_path):
        try: return pd.read_excel(file_path)
        except Exception: pass
    return pd.DataFrame()

def get_hanko_bytes():
    for d in [GDRIVE_SYNC_DIR, BASE_DIR, OTHER_DATA_DIR, BASE_SHARE]:
        if os.path.exists(d):
            for p in ["hanko.png", "*hanko*.png", "*hanko*.jpg", "*05171642*.png", "*印鑑*.png", "input_file_5.png"]:
                fl = glob.glob(os.path.join(d, p))
                if not fl:
                    fl = glob.glob(os.path.join(d, "**", p), recursive=True)
                if fl:
                    try:
                        with open(fl[0], "rb") as f: return f.read(), fl[0]
                    except Exception: pass
    return None, None

def load_saved_quotes():
    for f in [SAVED_QUOTES_FILE, LOCAL_QUOTES_FILE, os.path.join(GDRIVE_SYNC_DIR, "saved_estimates.json")]:
        if os.path.exists(f):
            try:
                with open(f, "r", encoding="utf-8") as fp: return json.load(fp)
            except Exception: pass
    return []

def save_quotes_list(q_list):
    target_f = SAVED_QUOTES_FILE if os.path.exists(os.path.dirname(SAVED_QUOTES_FILE)) else LOCAL_QUOTES_FILE
    try:
        with open(target_f, "w", encoding="utf-8") as f:
            json.dump(q_list, f, ensure_ascii=False, indent=2)
        upload_file_to_gdrive(target_f, "saved_estimates.json")
    except Exception:
        try:
            with open(LOCAL_QUOTES_FILE, "w", encoding="utf-8") as f:
                json.dump(q_list, f, ensure_ascii=False, indent=2)
            upload_file_to_gdrive(LOCAL_QUOTES_FILE, "saved_estimates.json")
        except Exception:
            pass

def load_saved_meeting_sheet_data():
    candidates = [
        SAVED_MEETING_FILE,
        LOCAL_MEETING_FILE,
        os.path.join(BASE_DIR, "meeting_sheet_saved.json"),
        os.path.join(GDRIVE_SYNC_DIR, "meeting_sheet_saved.json"),
        os.path.join(GDRIVE_SYNC_DIR, "saved_meeting_sheet.json"),
        os.path.join(BASE_DIR, "meeting_data_cache.json")
    ]
    for f in candidates:
        if os.path.exists(f):
            try:
                with open(f, "r", encoding="utf-8") as fp:
                    data = json.load(fp)
                    if isinstance(data, dict):
                        return data
            except Exception:
                pass
    return {}

def save_meeting_sheet_data(data_dict):
    target_f = SAVED_MEETING_FILE if os.path.exists(os.path.dirname(SAVED_MEETING_FILE)) else LOCAL_MEETING_FILE
    try:
        with open(target_f, "w", encoding="utf-8") as fp:
            json.dump(data_dict, fp, ensure_ascii=False, indent=2)
        upload_file_to_gdrive(target_f, "meeting_sheet_saved.json")
    except Exception:
        try:
            with open(LOCAL_MEETING_FILE, "w", encoding="utf-8") as fp:
                json.dump(data_dict, fp, ensure_ascii=False, indent=2)
            upload_file_to_gdrive(LOCAL_MEETING_FILE, "meeting_sheet_saved.json")
        except Exception:
            pass

def create_smaregi_order_csv_bytes(est_data, raw_title_map, std_price_map):
    rows = []
    pay_map = {"銀行振込": 3, "代金引換": 2, "現行通り": 3, "掛け払い（月末締め翌月末払い）": 3}
    pay_id = pay_map.get(est_data.get("payment", ""), 3)
    now_dt_str = datetime.now().strftime("%Y/%m/%d %H:%M")
    ship_dt_str = datetime.now().strftime("%Y/%m/%d")
    deliv_dt_str = (datetime.now() + timedelta(days=1)).strftime("%Y/%m/%d")

    for item in est_data.get("items", []):
        sku = str(item.get("商品コード", "")).strip()
        qty = int(clean_num(item.get("数量", 1)))
        unit_p = int(clean_num(item.get("見積単価", 0)))
        line_amt = unit_p * qty
        unit_tax = int(round(unit_p * 0.10))
        unit_tax_in = unit_p + unit_tax
        jan = str(item.get("JANコード", "")).strip()
        clean_name = str(item.get("商品名", "")).strip()
        full_prod_name = (raw_title_map or {}).get(sku, f"【{sku}】{clean_name}（{jan}）" if jan else f"【{sku}】{clean_name}")
        std_p = (std_price_map or {}).get(sku, "")

        r_dict = {col: "" for col in ORDER_CSV_COLUMNS}
        r_dict.update({
            '顧客ID': est_data.get("cust_id", -1), '要望等': est_data.get("memo", ""),
            '注文者お名前(姓)': est_data.get("cust_sei", "ご担当者"), '注文者お名前(名)': est_data.get("cust_mei", "様"),
            '注文者お名前(フリガナ・姓)': est_data.get("cust_kana_sei", ""), '注文者お名前(フリガナ名)': est_data.get("cust_kana_mei", ""),
            '注文者会社名': est_data.get("cust_comp", ""), '注文者メールアドレス': est_data.get("cust_email", ""),
            '注文者メールアドレスタイプ': 1, '注文者電話番号': est_data.get("cust_tel", ""), '注文者FAX': est_data.get("cust_fax", ""),
            '注文者郵便番号': est_data.get("cust_post", ""), '注文者都道府県': est_data.get("cust_pref", ""),
            '注文者住所1': est_data.get("cust_addr1", ""), '注文者住所2': est_data.get("cust_addr2", ""),
            '注文者性別ID': 1, '注文者性別': "男性", '注文者職種': est_data.get("cust_job_id", ""),
            '商品合計': line_amt, '調整額': 0, 'うち割引額１': 0, 'うち割引額２': 0, 'うち割引額３': 0, 'うちその他調整額': 0,
            '合計額': line_amt, '送料（税別）': est_data.get("shipping_fee", 0), '手数料（税別）': est_data.get("cod_fee", 0),
            'クーポン名': str(est_data.get("rank_name", "")).split(":")[0].strip(), 'クーポン利用': 0, 'ポイント利用': 0,
            '請求額': line_amt, '消費税額': unit_tax, '入金額': 0, 'EC種類': 16,
            'EC注文番号': est_data.get("q_no", "").replace("-", ""), 'EC店舗ID': 1000, 'EC店舗名': "旭電機化成",
            '販売ルートID': 1, '受注ステータス': 1, '支払方法ID': pay_id, '決済ステータス': 0, '配送方法ID': 1,
            '配送区分': 0, '申込番号種類': 0, '注文日時': est_data.get("created_at", now_dt_str),
            '出荷予定日': ship_dt_str, 'ひとことメモ': est_data.get("hitokoto", ""), '伝達事項': est_data.get("memo", ""),
            '再計算フラグ': 1, '送付先お名前(姓)': est_data.get("ship_sei", est_data.get("cust_sei", "ご担当者")),
            '送付先お名前(名)': est_data.get("ship_mei", est_data.get("cust_mei", "様")),
            '送付先会社名': est_data.get("ship_comp", est_data.get("cust_comp", "")),
            '送付先電話番号': est_data.get("ship_tel", est_data.get("cust_tel", "")),
            '送付先郵便番号': est_data.get("ship_post", est_data.get("cust_post", "")),
            '送付先都道府県': est_data.get("ship_pref", est_data.get("cust_pref", "")),
            '送付先住所1': est_data.get("ship_addr1", est_data.get("cust_addr1", "")),
            '送付先住所2': est_data.get("ship_addr2", est_data.get("cust_addr2", "")),
            'お届け日': deliv_dt_str, '出力フラグ(受注伝票)': 0, '出力フラグ(お手紙)': 0, '出力フラグ(納品書)': 1,
            '出力フラグ(領収書)': 0, '出力フラグ(請求書)': 0, '出力フラグ(EMS)': 0, '出力フラグ(送り状)': 0,
            'ご購入明細商品コード': sku, 'ご購入明細商品名': full_prod_name, 'ご購入明細消費税区分': "別",
            'ご購入明細消費税課税規則': 3, 'ご購入明細送料区分': "別", 'ご購入明細手数料区分': "別",
            'ご購入明細税抜単価': unit_p, 'ご購入明細税込単価': unit_tax_in, 'ご購入明細消費税': unit_tax,
            'ご購入明細消費税率': 10, 'ご購入明細数量': qty, 'ご購入明細金額': line_amt,
            'ご購入明細定期区分': "商品", 'ご購入明細標準販売価格': std_p, '消費税計算区分': "伝票単位(B)"
        })
        for flag_col in [
            '出力済(受注伝票)', '出力済(お手紙)', '出力済(納品書)', '出力済(領収書)', '出力済(請求書)',
            '出力済(EMS)', '出力済(送り状)', '出力済(後払い取引登録)', '出力済(後払い出荷登録)',
            '出力済(配送指示伝票)', '出力済(モール連携出荷)', '出力済(モール連携処理済)',
            'メール送信済(サンクス)', 'メール送信済(入金確認)', 'メール送信済(発送完了)',
            'メール送信済(フォロー)', 'メール送信済(その他１)', 'メール送信済(その他２)'
        ]:
            r_dict[flag_col] = 0
        rows.append(r_dict)

    out_df = pd.DataFrame(rows, columns=ORDER_CSV_COLUMNS)
    return out_df.to_csv(index=False, encoding="cp932", errors="replace").encode("cp932")

def _read_single_csv(f, explicit_name=None):
    fname = explicit_name if explicit_name else str(f)
    bname = os.path.basename(fname)
    path_date = None
    ym_match = re.search(r'(\d{4})年?[/\\_]?(\d{1,2})月?[/\\_]?yayoi_(\d{2})(\d{2})', fname, re.IGNORECASE)
    if ym_match:
        y, m, md_m, md_d = ym_match.groups()
        path_date = f"{y}{md_m}{md_d}"
    else:
        y_m = re.search(r'(\d{4})年', fname)
        d_m = re.search(r'yayoi_(\d{2})(\d{2})', fname, re.IGNORECASE)
        if y_m and d_m:
            path_date = f"{y_m.group(1)}{d_m.group(1)}{d_m.group(2)}"
        elif d_m:
            mm_str, dd_str = d_m.group(1), d_m.group(2)
            mm_int = int(mm_str) if mm_str.isdigit() else 11
            inferred_y = 2026 if mm_int >= 7 else 2027
            path_date = f"{inferred_y}{mm_str}{dd_str}"

    for enc in ["cp932", "shift_jis", "utf-8", "utf-8-sig"]:
        try:
            if hasattr(f, "seek"): f.seek(0)
            d = pd.read_csv(f, encoding=enc, encoding_errors="replace", low_memory=False)
            if d.empty:
                continue
            cols = [str(c).strip() for c in d.columns]
            has_header = any(any(k in c for k in ["商品", "得意先", "数量", "単価", "金額", "日付", "売上", "伝票"]) for c in cols)

            if not has_header and len(cols) >= 20:
                if hasattr(f, "seek"): f.seek(0)
                d = pd.read_csv(f, encoding=enc, encoding_errors="replace", header=None, low_memory=False)
                col_map = {3: "売上日", 5: "得意先名", 11: "担当者コード", 14: "商品コード", 15: "商品名", 19: "数量", 21: "単価", 22: "金額"}
                d = d.rename(columns={idx: name for idx, name in col_map.items() if idx < len(d.columns)})
                cols = [str(c).strip() for c in d.columns]

            qty_cols = [c for c in cols if "数量" in c]
            d["_auto_qty"] = d[qty_cols[0]].apply(clean_num) if qty_cols else (d.iloc[:, 19].apply(clean_num) if len(cols) >= 20 else 0.0)
            price_cols = [c for c in cols if "単価" in c]
            d["_auto_price"] = d[price_cols[0]].apply(clean_num) if price_cols else (d.iloc[:, 21].apply(clean_num) if len(cols) >= 22 else 0.0)
            amt_cols = [c for c in cols if c in ["金額", "売上金額", "販売金額"]]
            if amt_cols:
                d["_auto_amt"] = d[amt_cols[0]].apply(clean_num)
                d["_auto_amt"] = d.apply(lambda r: (r["_auto_qty"] * r["_auto_price"]) if r["_auto_amt"] == 0 and r["_auto_price"] > 0 else r["_auto_amt"], axis=1)
            else:
                d["_auto_amt"] = d["_auto_qty"] * d["_auto_price"]
            d["_fallback_date"] = path_date
            d["_source_file"] = bname
            return d
        except Exception:
            continue
    return None

def _parse_ym_from_filename_or_override(fname, override_ym=None):
    if override_ym:
        m_ov = re.search(r'(\d{4})\D+(\d{1,2})', str(override_ym))
        if m_ov:
            return int(m_ov.group(1)), int(m_ov.group(2))
    norm_f = str(fname).replace(os.sep, '/')
    m_full = re.search(r'(20\d{2})[年/\-_ ]*(\d{1,2})月?', norm_f)
    if m_full:
        y_i, m_i = int(m_full.group(1)), int(m_full.group(2))
        if 1 <= m_i <= 12: return y_i, m_i
    m_month_only = re.search(r'(?<!\d)(1[0-2]|0?[1-9])月', norm_f)
    if m_month_only:
        m_i = int(m_month_only.group(1))
        return (2026 if m_i >= 7 else 2027), m_i
    now_dt = datetime.now()
    return now_dt.year, now_dt.month

def _read_amazon_seller_csv(f, explicit_name=None, override_ym=None):
    fname = explicit_name if explicit_name else str(f)
    y_i, m_i = _parse_ym_from_filename_or_override(fname, override_ym)
    target_ym = f"{y_i}年{m_i:02d}月"
    try: dt_last = pd.Timestamp(year=y_i, month=m_i, day=1) + pd.offsets.MonthEnd(1)
    except Exception: dt_last = pd.Timestamp(year=y_i, month=m_i, day=28)

    for enc in ["utf-8-sig", "utf-8", "cp932", "shift_jis"]:
        try:
            if hasattr(f, "seek"): f.seek(0)
            d = pd.read_csv(f, encoding=enc, encoding_errors="replace", low_memory=False)
            if d.empty: continue
            sku_series = next((d[c].astype(str).str.strip() for c in d.columns if str(c).strip().upper() == "SKU"), None)
            if sku_series is None:
                sku_series = d.iloc[:, 3].astype(str).str.strip() if len(d.columns) >= 4 else d.iloc[:, 0].astype(str).str.strip()
            qty_series = next((d[c].apply(clean_num).astype(int) for c in d.columns if any(k in str(c) for k in ['点数', '数量', '個数'])), None)
            if qty_series is None:
                qty_series = d.iloc[:, 14].apply(clean_num).astype(int) if len(d.columns) >= 15 else (d.iloc[:, 2].apply(clean_num).astype(int) if len(d.columns) >= 3 else pd.Series([1] * len(d)))
            amt_series = None
            for c in d.columns:
                if any(k in str(c) for k in ['売上額', '売上金額', '売上', '金額']):
                    amt_series = (d[c].apply(clean_num) / 1.10).round().astype(int)
                    break
            if amt_series is None:
                if len(d.columns) >= 19: amt_series = (d.iloc[:, 18].apply(clean_num) / 1.10).round().astype(int)
                elif len(d.columns) >= 2: amt_series = (d.iloc[:, 1].apply(clean_num) / 1.10).round().astype(int)
                else: continue
            res = pd.DataFrame({
                "商品コード": sku_series, "元商品名": sku_series, "商品名": sku_series,
                "得意先名": "アマゾンセラー", "担当者": "水原", "数量": qty_series, "金額": amt_series,
                "日付_dt": dt_last, "日付": dt_last.strftime('%Y/%m/%d'), "年月": target_ym,
                "_source_file": os.path.basename(fname)
            })
            res = res[res["商品コード"].astype(str).str.strip().ne("") & ~res["商品コード"].astype(str).isin(["nan", "None"])].copy()
            res["単価"] = (res["金額"] / res["数量"].replace(0, 1)).round().astype(int)
            return res
        except Exception: continue
    return None

def _parse_raw_df_to_standard(df_all):
    if df_all.empty: return pd.DataFrame()
    def find_col(candidates, cols):
        for cand in candidates:
            for c in cols:
                if cand == c or cand in c: return c
        return None
    all_cols = [str(c).strip() for c in df_all.columns]
    c_date = find_col(["納品日", "伝票日付", "売上日", "日付"], all_cols)
    c_pcode = find_col(["商品コード", "商品CD", "品番"], all_cols)
    c_pname = find_col(["商品名/入金内容", "商品名", "商品名_摘要", "品名"], all_cols)
    c_cname = find_col(["得意先名称", "得意先名", "納品先", "得意先"], all_cols)
    c_staff = find_col(["担当者コード", "担当者CD", "担当者名", "担当者", "担当", "営業担当"], all_cols)

    res = pd.DataFrame()
    if c_pcode: res["商品コード"] = df_all[c_pcode].fillna("").astype(str).str.strip()
    else: res["商品コード"] = ""
    if c_pname:
        res["元商品名"] = df_all[c_pname].fillna("").astype(str).str.strip()
        res["商品名"] = df_all[c_pname].fillna("").astype(str).map(clean_product_name)
    else:
        res["元商品名"], res["商品名"] = "", ""
    res["得意先名"] = df_all[c_cname].fillna("").astype(str).str.strip().map(clean_cust_name) if c_cname else ""
    res["担当者"] = df_all[c_staff].map(convert_staff_name).fillna("").astype(str).replace("nan", "") if c_staff else ""
    res["数量"] = df_all["_auto_qty"].fillna(0).astype(int)
    res["単価"] = df_all["_auto_price"].fillna(0).astype(int)
    res["金額"] = df_all["_auto_amt"].fillna(res["数量"] * res["単価"]).astype(int)
    if "_source_file" in df_all.columns: res["_source_file"] = df_all["_source_file"].fillna("").astype(str)

    raw_s = df_all[c_date].fillna("").astype(str).str.strip() if c_date else pd.Series([""] * len(df_all))
    clean_date = raw_s.str.replace(r'\.0$', '', regex=True).str.replace(r'[-/.\s年月日]', '', regex=True)
    dt = pd.to_datetime(clean_date, format="%Y%m%d", errors="coerce").fillna(pd.to_datetime(raw_s, errors="coerce"))
    if "_fallback_date" in df_all.columns:
        fb_dt = pd.to_datetime(df_all["_fallback_date"].fillna("").astype(str), format="%Y%m%d", errors="coerce")
        dt = dt.fillna(fb_dt)

    res["日付_dt"] = dt
    res["日付"] = dt.dt.strftime('%Y/%m/%d').fillna("")
    res["年月"] = dt.dt.strftime('%Y年%m月').fillna("")
    res = res[(res["商品コード"] != "") & ~res["商品コード"].isin(["nan", "None"])].copy()
    if "商品名" in res.columns:
        res = res[~res["商品名"].str.contains("送料", na=False)].copy()
    return res

def finalize_sales_df(df, cost_master, code_name_dict=None):
    if df.empty: return df
    df["日付_dt"] = pd.to_datetime(df["日付_dt"], errors="coerce")
    for s_col in ["商品コード", "商品名", "元商品名", "得意先名", "担当者", "日付", "年月", "_source_file"]:
        if s_col in df.columns:
            df[s_col] = df[s_col].fillna("").astype(str).replace({"nan": "", "None": ""})
        else:
            df[s_col] = ""

    if code_name_dict:
        mask_same = (df["商品名"] == df["商品コード"]) | (df["商品名"] == "")
        df.loc[mask_same, "商品名"] = df.loc[mask_same, "商品コード"].map(lambda c: code_name_dict.get(c, c))

    if "担当者" in df.columns and "得意先名" in df.columns:
        valid_staff_df = df[df["担当者"].str.strip().ne("")].sort_values("日付_dt")
        if not valid_staff_df.empty:
            cust_current_staff_map = valid_staff_df.groupby("得意先名")["担当者"].last().to_dict()
            df["担当者"] = df["得意先名"].map(lambda c: cust_current_staff_map.get(c, ""))

    for n_col in ["数量", "単価", "金額"]:
        if n_col in df.columns:
            df[n_col] = df[n_col].apply(clean_num).round().astype(int)

    if "商品コード" in df.columns:
        df["単価原価"] = df["商品コード"].map(lambda c: cost_master.get(str(c).strip(), 0.0))
        df["原価合計"] = (df["数量"] * df["単価原価"]).round().astype(int)
        df["粗利額"] = df["金額"] - df["原価合計"]
    return df

def _get_data_files_signature():
    sig = []
    for d in [GDRIVE_SYNC_DIR, BASE_DIR]:
        if os.path.exists(d):
            for fn in sorted(os.listdir(d)):
                if fn.lower().endswith((".csv", ".parquet", ".xlsx")):
                    full_p = os.path.join(d, fn)
                    try:
                        sig.append(f"{fn}_{int(os.path.getmtime(full_p))}_{os.path.getsize(full_p)}")
                    except Exception:
                        pass
    return "|".join(sig)

@st.cache_data(show_spinner=False)
def load_and_prep_data(_file_sig=""):
    cost_path, cost_mtime = get_cost_master_file_info()
    cost_master = load_cost_master(cost_path, cost_mtime)
    jan_path, jan_mtime = get_jan_file_info()
    _, jan_title_dict_local, _ = load_jan_master_v2(jan_path, jan_mtime)

    cached_df = pd.DataFrame()
    existing_sources = set()
    existing_date_sources = set()
    gdrive_pq = os.path.join(GDRIVE_SYNC_DIR, "sales_cache.parquet")
    cache_candidates = [gdrive_pq, CACHE_FILE]

    loaded_pqs = []
    for cf in cache_candidates:
        if os.path.exists(cf):
            try:
                tmp_df = pd.read_parquet(cf)
                if not tmp_df.empty:
                    loaded_pqs.append(tmp_df)
            except Exception:
                pass

    if len(loaded_pqs) == 1:
        cached_df = loaded_pqs[0]
    elif len(loaded_pqs) >= 2:
        loaded_pqs.sort(key=lambda x: len(x), reverse=True)
        base_pq = loaded_pqs[0]
        other_pq = loaded_pqs[1]
        if "_source_file" in base_pq.columns and "_source_file" in other_pq.columns and "日付" in base_pq.columns and "日付" in other_pq.columns:
            base_keys = set(zip(base_pq["日付"].astype(str), base_pq["_source_file"].astype(str)))
            other_keys = list(zip(other_pq["日付"].astype(str), other_pq["_source_file"].astype(str)))
            mask_missing = [k not in base_keys for k in other_keys]
            missing_part = other_pq[mask_missing]
            cached_df = pd.concat([base_pq, missing_part], ignore_index=True) if not missing_part.empty else base_pq
        else:
            cached_df = base_pq

    if not cached_df.empty and "_source_file" in cached_df.columns:
        existing_sources = set(cached_df["_source_file"].dropna().astype(str).unique())
        if "日付" in cached_df.columns:
            # ★重要：「2025/11/11 の yayoi_1111.csv」と「2026/11/11 の yayoi_1111.csv」を明確に区別する！
            existing_date_sources = set(zip(cached_df["日付"].fillna("").astype(str), cached_df["_source_file"].fillna("").astype(str)))

    raw_candidates = []
    for search_root in [GDRIVE_SYNC_DIR, BASE_SHARE, BASE_DIR]:
        if os.path.exists(search_root):
            raw_candidates.extend(glob.glob(os.path.join(search_root, "*.csv")))

    unique_by_name = {}
    for f in raw_candidates:
        bname = os.path.basename(f)
        if bname not in unique_by_name or GDRIVE_SYNC_DIR in os.path.abspath(f):
            unique_by_name[bname] = f

    csv_candidates = []
    for bname in sorted(unique_by_name.keys()):
        f = unique_by_name[bname]
        if bname.startswith("order_") or bname.startswith("item_class_lotmember") or "ランク" in bname or "ロジザード" in bname or "在庫" in bname:
            continue
        if AMZ_DAILY_STORE_DIR in os.path.abspath(f) or re.match(r'^\d{1,2}[-_]\d{1,2}\.csv$', bname):
            continue

        # yayoi_MMDD.csv の場合、今年の推定日付（例: 2026/11/11）がすでにキャッシュにあるか確認する
        m_yayoi = re.search(r'yayoi_(\d{2})(\d{2})', bname, re.IGNORECASE)
        if m_yayoi:
            mm_s, dd_s = m_yayoi.group(1), m_yayoi.group(2)
            inf_y = 2026 if int(mm_s) >= 7 else 2027
            expected_dt_str = f"{inf_y}/{mm_s}/{dd_s}"
            # 昨年の yayoi_1111.csv が入っていても、2026/11/11 の yayoi_1111.csv が無ければ必ず取り込む！
            if (expected_dt_str, bname) not in existing_date_sources:
                csv_candidates.append(f)
        else:
            if bname not in existing_sources:
                csv_candidates.append(f)

    new_yayoi_df = pd.DataFrame()
    new_amz_df = pd.DataFrame()
    if csv_candidates:
        yayoi_files = [f for f in csv_candidates if "yayoi" in os.path.basename(f).lower()]
        amz_files = [f for f in csv_candidates if f not in yayoi_files and any(k in f.lower() for k in ["アマゾン", "amazon", "businessreport", "月"])]
        if not yayoi_files and not amz_files:
            yayoi_files = csv_candidates

        if yayoi_files:
            with ThreadPoolExecutor(max_workers=8) as executor:
                res_list = list(executor.map(_read_single_csv, yayoi_files))
            valid_res = [d for d in res_list if d is not None and not d.empty]
            if valid_res:
                new_yayoi_df = _parse_raw_df_to_standard(pd.concat(valid_res, ignore_index=True))

        if amz_files:
            amz_res = [_read_amazon_seller_csv(af) for af in amz_files]
            valid_amz = [ad for ad in amz_res if ad is not None and not ad.empty]
            if valid_amz:
                new_amz_df = pd.concat(valid_amz, ignore_index=True)

    # もし同じ（日付, _source_file）のデータが既にあれば重複しないように入れ替える（昨年の同名ファイルは絶対に消さない！）
    if not cached_df.empty and not new_yayoi_df.empty:
        new_keys = set(zip(new_yayoi_df["日付"].astype(str), new_yayoi_df["_source_file"].astype(str)))
        old_keys = list(zip(cached_df["日付"].astype(str), cached_df["_source_file"].astype(str)))
        cached_df = cached_df[[k not in new_keys for k in old_keys]].copy()

    if not cached_df.empty and not new_amz_df.empty:
        amz_yms = set(new_amz_df["年月"].dropna().unique())
        cached_df = cached_df[~((cached_df["得意先名"] == "アマゾンセラー") & (cached_df["年月"].isin(amz_yms)))].copy()

    all_parts = [d for d in [cached_df, new_yayoi_df, new_amz_df] if d is not None and not d.empty]
    df = pd.concat(all_parts, ignore_index=True) if all_parts else pd.DataFrame()

    if not df.empty:
        name_lookup = dict(jan_title_dict_local or {})
        valid_names = df[(df["商品名"].notna()) & (df["商品名"].astype(str) != df["商品コード"].astype(str)) & (df["商品名"].astype(str).str.strip() != "")]
        if not valid_names.empty:
            name_lookup.update(valid_names.drop_duplicates("商品コード").set_index("商品コード")["商品名"].to_dict())
        df = finalize_sales_df(df, cost_master, name_lookup)
        if not new_yayoi_df.empty or not new_amz_df.empty:
            try:
                save_parquet_and_sync(df, CACHE_FILE, "sales_cache.parquet")
            except Exception:
                pass
    return df, len(csv_candidates), cost_master

def render_universal_table(df_data, max_h=600):
    if df_data is None or df_data.empty:
        st.info("該当するデータがありません。")
        return
    center_cols = {"番号", "順位", "日付", "販売期間", "最終売上日", "初回売上日", "未販売期間", "目標達成率", "通期進捗率", "前年比", "粗利率", "構成比", "倉庫", "センター", "判定", "受注日", "納入期日"}
    left_cols = {"得意先名", "商品コード", "商品名", "商品番号", "ランク名", "担当者", "客先名", "販売価格帯", "備考", "決定事項", "入荷予定日"}
    narrow_cols = {"番号", "順位"}

    header_cells = []
    for c in df_data.columns:
        align = 'center' if c in center_cols else ('left' if c in left_cols else 'right')
        w_style = "width: 48px; min-width: 48px; max-width: 52px;" if c in narrow_cols else ""
        header_cells.append(f"<th style='border:1px solid #cbd5e1; padding:8px 10px; background:#f1f5f9; text-align:{align}; {w_style} white-space:nowrap;'>{c}</th>")
    header_html = "".join(header_cells)

    body_rows = []
    for r_i, row in df_data.iterrows():
        bg = "#ffffff" if r_i % 2 == 0 else "#f8fafc"
        cells = []
        for c in df_data.columns:
            val = row[c]
            val_str = "" if pd.isna(val) else str(val)
            align = "center" if c in center_cols else ("left" if c in left_cols else "right")
            w_style = "width: 48px; min-width: 48px; max-width: 52px;" if c in narrow_cols else ""
            if isinstance(val, (int, float)) and c not in center_cols:
                val_str = f"{int(val):+,}" if c in ["増減", "増減額", "目標差額", "粗利増減", "発注検討差", "確保ズレ"] else f"{int(val):,}"
            is_neg = (isinstance(val, (int, float)) and val < 0) or (isinstance(val_str, str) and (val_str.startswith("-") or val_str.startswith("▲-") or "🚨" in val_str))
            style_extra = "color: #dc2626; font-weight: bold;" if is_neg else ("color: #16a34a; font-weight: bold;" if "✅" in val_str else "color: #0f172a;")
            cells.append(f"<td style='border:1px solid #e2e8f0; padding:6px 10px; text-align:{align}; {w_style} font-variant-numeric:tabular-nums; {style_extra} white-space:nowrap;'>{val_str}</td>")
        body_rows.append(f"<tr style='background:{bg};'>{''.join(cells)}</tr>")

    table_html = f"""
    <div style="overflow-x: auto; max-height: {max_h}px; border: 1px solid #cbd5e1; border-radius: 6px; background-color: #ffffff;">
    <table style="width: 100%; border-collapse: collapse; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; font-size: 13px;">
    <thead><tr style="position: sticky; top: 0; z-index: 1;">{header_html}</tr></thead>
    <tbody>{''.join(body_rows)}</tbody>
    </table>
    </div>
    """
    components.html(table_html, height=min(max_h + 20, 80 + len(df_data) * 36), scrolling=True)

_CLICKABLE_TABLE_DIR = os.path.join(BASE_DIR, "_cust_table_comp_v3")
os.makedirs(_CLICKABLE_TABLE_DIR, exist_ok=True)
_CLICKABLE_INDEX_HTML = os.path.join(_CLICKABLE_TABLE_DIR, "index.html")
_COMP_HTML_CONTENT = """<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
  body { margin: 0; padding: 0; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; font-size: 13px; background: transparent; }
  .tbl-wrap { overflow-x: auto; max-height: 580px; border: 1px solid #cbd5e1; border-radius: 6px; background-color: #ffffff; display: inline-block; max-width: 100%; }
  table { border-collapse: collapse; table-layout: fixed; width: auto; }
  th { border: 1px solid #cbd5e1; padding: 8px 10px; background: #f1f5f9; white-space: nowrap; position: sticky; top: 0; z-index: 1; color: #0f172a; }
  td { border: 1px solid #e2e8f0; padding: 6px 10px; font-variant-numeric: tabular-nums; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; color: #0f172a; }
  tr:nth-child(even) { background: #f8fafc; }
  tr:nth-child(odd) { background: #ffffff; }
  tr:hover { background: #eff6ff !important; }
  tr.selected-row { background: #dbeafe !important; }
  .cust-link { color: #1d4ed8; font-weight: bold; text-decoration: underline; cursor: pointer; }
  .cust-link:hover { color: #1e40af; }
</style>
</head>
<body>
<div id="root"></div>
<script>
function sendMessageToStreamlitClient(type, data) {
  var outData = Object.assign({ isStreamlitMessage: true, type: type }, data);
  window.parent.postMessage(outData, "*");
}
function init() { sendMessageToStreamlitClient("streamlit:componentReady", {apiVersion: 1}); }
function setFrameHeight(height) { sendMessageToStreamlitClient("streamlit:setFrameHeight", {height: height}); }
function sendValue(value) { sendMessageToStreamlitClient("streamlit:setComponentValue", {value: value}); }

window.addEventListener("message", function(event) {
  if (event.data.type !== "streamlit:render") return;
  var args = event.data.args;
  var root = document.getElementById("root");
  root.innerHTML = args.table_html || "";
  var links = root.querySelectorAll(".cust-link");
  links.forEach(function(el) {
    el.addEventListener("click", function() {
      var cname = this.getAttribute("data-cust");
      var screenW = window.parent ? window.parent.innerWidth : window.innerWidth;
      sendValue({cust: cname, ts: Date.now(), is_mobile: (screenW <= 768)});
    });
  });
  setFrameHeight(args.height || 595);
});
init();
</script>
</body>
</html>"""
try:
    with open(_CLICKABLE_INDEX_HTML, "w", encoding="utf-8") as _f:
        _f.write(_COMP_HTML_CONTENT)
except Exception:
    pass
_clickable_cust_table_func = components.declare_component("clickable_cust_table_v3", path=_CLICKABLE_TABLE_DIR)

def render_clickable_customer_table(df_data, selected_cust=""):
    if df_data is None or df_data.empty:
        st.info("該当するデータがありません。")
        return None
    center_cols = {"番号", "順位", "日付", "販売期間", "最終売上日", "初回売上日", "未販売期間", "目標達成率", "通期進捗率", "前年比", "粗利率", "構成比"}
    left_cols = {"得意先名", "商品コード", "商品名", "商品番号", "ランク名", "担当者", "客先名", "販売価格帯"}

    col_width_map = {
        "順位": "44px", "得意先名": "220px", "担当者": "64px", "販売金額": "105px",
        "粗利額": "100px", "粗利率": "68px", "原価合計": "100px", "数量": "74px", "取扱品目数": "74px", "構成比": "68px"
    }

    header_cells = []
    for c in df_data.columns:
        align = 'center' if c in center_cols else ('left' if c in left_cols else 'right')
        w_px = col_width_map.get(c, "90px")
        header_cells.append(f"<th style='text-align:{align}; width:{w_px}; min-width:{w_px}; max-width:{w_px};'>{c}</th>")

    body_rows = []
    for _, row in df_data.iterrows():
        cells = []
        row_cust = str(row.get("得意先名", ""))
        row_cls = " class='selected-row'" if row_cust and row_cust == selected_cust else ""
        for c in df_data.columns:
            val = row[c]
            val_str = "" if pd.isna(val) else str(val)
            align = "center" if c in center_cols else ("left" if c in left_cols else "right")
            w_px = col_width_map.get(c, "90px")
            if isinstance(val, (int, float)) and c not in center_cols:
                val_str = f"{int(val):+,}" if c in ["増減", "増減額", "目標差額", "粗利増減"] else f"{int(val):,}"
            is_neg = (isinstance(val, (int, float)) and val < 0) or (isinstance(val_str, str) and (val_str.startswith("-") or val_str.startswith("▲-")))
            style_extra = "color: #dc2626; font-weight: bold;" if is_neg else ""

            if c == "得意先名":
                safe_c = val_str.replace('"', '&quot;')
                cells.append(f"<td style='text-align:left; width:{w_px}; min-width:{w_px}; max-width:{w_px};' title=\"{safe_c}\"><span class='cust-link' data-cust=\"{safe_c}\">🔍 {val_str}</span></td>")
            else:
                cells.append(f"<td style='text-align:{align}; width:{w_px}; min-width:{w_px}; max-width:{w_px}; {style_extra}'>{val_str}</td>")
        body_rows.append(f"<tr{row_cls}>{''.join(cells)}</tr>")

    full_html = f"<div class='tbl-wrap'><table><thead><tr>{''.join(header_cells)}</tr></thead><tbody>{''.join(body_rows)}</tbody></table></div>"
    calc_h = min(600, 75 + len(df_data) * 35)
    return _clickable_cust_table_func(table_html=full_html, height=calc_h, key="cust_rank_clickable_tbl_v3")

def build_customer_all_products_df(sub_df):
    if sub_df is None or sub_df.empty:
        return pd.DataFrame()
    p_agg = sub_df.groupby(["商品コード", "商品名"]).agg(
        数量=("数量", "sum"), 販売金額=("金額", "sum"), 粗利額=("粗利額", "sum"),
        単価_min=("単価", "min"), 単価_max=("単価", "max")
    ).reset_index().sort_values(by="販売金額", ascending=False)
    p_agg["販売価格帯"] = p_agg.apply(lambda r: f"¥{int(r['単価_min']):,}" if r['単価_min'] == r['単価_max'] else f"¥{int(r['単価_min']):,}〜¥{int(r['単価_max']):,}", axis=1)
    p_agg["粗利率"] = p_agg.apply(lambda r: f"{(r['粗利額'] / r['販売金額'] * 100):.1f}%" if r['販売金額'] > 0 else "-", axis=1)
    p_agg.insert(0, "順位", range(1, len(p_agg) + 1))
    return p_agg[["順位", "商品コード", "商品名", "数量", "販売価格帯", "販売金額", "粗利額", "粗利率"]]

@st.dialog("🏢 得意先別 購入商品一覧（全件）", width="large")
def show_customer_top10_dialog(cust_name, sub_df, period_str):
    st.markdown(f"#### 🎯 対象得意先: **{cust_name}**")
    st.caption(f"指定期間: {period_str}")
    show_p_df = build_customer_all_products_df(sub_df)
    if show_p_df.empty:
        st.info("該当期間の売上データはありません。")
        return
    tot_s = int(show_p_df["販売金額"].sum())
    tot_q = int(show_p_df["数量"].sum())
    tot_cnt = len(show_p_df)
    c_d1, c_d2 = st.columns([3.2, 1.5])
    with c_d1:
        st.markdown(f"<div class='total-card' style='margin-bottom:0;'><b>該当品目数:</b> {tot_cnt:,} 品目 &nbsp;|&nbsp; <b>数量合計:</b> {tot_q:,} 個 &nbsp;|&nbsp; <b>売上合計:</b> <span style='color:#d93025; font-weight:bold;'>¥{tot_s:,}</span></div>", unsafe_allow_html=True)
    with c_d2:
        st.download_button(
            label="📥 この得意先の内訳をExcel保存",
            data=generate_excel_bytes(show_p_df, f"{cust_name}_商品内訳"),
            file_name=f"得意先別商品内訳_{cust_name}_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key=f"dl_modal_cust_{cust_name}"
        )
    st.write("")
    render_universal_table(show_p_df, max_h=500)

def _extract_top5_and_other(sub_df):
    if sub_df is None or sub_df.empty:
        return [{"sku": "", "price": 0, "qty": 0, "amt": 0} for _ in range(5)], 0, 0
    agg = sub_df.groupby("商品コード").agg(数量=("数量", "sum"), 金額=("金額", "sum")).reset_index().sort_values(by="金額", ascending=False)
    total_amt = int(agg["金額"].sum())
    top5_df = agg.head(5)
    top5_sum = int(top5_df["金額"].sum())
    other_amt = max(0, total_amt - top5_sum)
    rows = []
    for _, r in top5_df.iterrows():
        q, a = int(r["数量"]), int(r["金額"])
        rows.append({"sku": str(r["商品コード"]), "price": int(round(a / q)) if q > 0 else 0, "qty": q, "amt": a})
    while len(rows) < 5:
        rows.append({"sku": "", "price": 0, "qty": 0, "amt": 0})
    return rows, other_amt, total_amt

def build_meeting_comparison_card_data(cust, cid, cur_m, ly_m, next_m, prev_fy, tgt_y, tgt_m, next_y, next_m_val, f_year, saved_dict=None):
    c_cur = cur_m[cur_m["得意先名"] == cust] if not cur_m.empty else pd.DataFrame()
    c_ly = ly_m[ly_m["得意先名"] == cust] if not ly_m.empty else pd.DataFrame()
    c_next = next_m[next_m["得意先名"] == cust] if not next_m.empty else pd.DataFrame()
    c_prev_fy = prev_fy[prev_fy["得意先名"] == cust] if not prev_fy.empty else pd.DataFrame()

    ly_rows, ly_other, ly_total = _extract_top5_and_other(c_ly)
    cur_rows, cur_other, cur_total = _extract_top5_and_other(c_cur)

    saved_cust_data = (saved_dict or {}).get(cust, {})
    if "custom_next_rows" in saved_cust_data:
        raw_saved = saved_cust_data["custom_next_rows"] or []
        next_rows = []
        for r in raw_saved[:5]:
            sk = str(r.get("sku", "") or "").strip()
            if sk in ["None", "nan"]: sk = ""
            p = int(clean_num(r.get("price", 0)))
            q = int(clean_num(r.get("qty", 0)))
            a = int(clean_num(r.get("amt", 0)))
            next_rows.append({"sku": sk, "price": p, "qty": q, "amt": a})
        while len(next_rows) < 5:
            next_rows.append({"sku": "", "price": 0, "qty": 0, "amt": 0})
        next_other = int(clean_num(saved_cust_data.get("custom_next_other", 0)))
        next_row_sum = sum(int(r["amt"]) for r in next_rows if r.get("sku") != "" or int(r.get("amt", 0)) > 0)
        next_total = next_row_sum + next_other
    else:
        next_rows, next_other, next_total = _extract_top5_and_other(c_next)

    prev_fy_total = int(c_prev_fy["金額"].sum()) if not c_prev_fy.empty else 0
    yoy_rate = f"{int(round(cur_total / ly_total * 100))}%" if ly_total > 0 else "-"
    next_yoy_rate = f"{int(round(next_total / prev_fy_total * 100))}%" if prev_fy_total > 0 else "0%"

    if not c_next.empty and "日付_dt" in c_next.columns and c_next["日付_dt"].notna().any():
        as_of_str = f"{next_y}年{next_m_val}月{c_next['日付_dt'].max().day}日現在"
    else:
        as_of_str = f"{next_y}年{next_m_val}月現在"

    return {
        "cust_name": cust, "cust_id": cid,
        "ly_label": f"{tgt_y - 1}年{tgt_m}月実績", "cur_label": f"{tgt_y}年{tgt_m}月実績",
        "next_label": f"{next_y}年{next_m_val}月売上進捗", "prev_fy_label": f"{next_y - 1}年{next_m_val}月度実績",
        "as_of_label": as_of_str, "ly_rows": ly_rows, "ly_other": ly_other, "ly_total": ly_total,
        "cur_rows": cur_rows, "cur_other": cur_other, "cur_total": cur_total, "yoy_rate": yoy_rate,
        "next_rows": next_rows, "next_other": next_other, "next_total": next_total,
        "prev_fy_total": prev_fy_total, "next_yoy_rate": next_yoy_rate,
        "activity_text": saved_cust_data.get("activity", ""), "plan_text": saved_cust_data.get("plan", "")
    }

def render_meeting_card_html(card):
    def fmt_row_v(r_obj, key):
        sk = str(r_obj.get("sku", "")).strip()
        v = r_obj.get(key)
        if key == "sku": return sk if sk else "&nbsp;"
        if not sk and (v is None or v == 0 or v == ""): return "&nbsp;"
        if v is None or v == "" or (isinstance(v, float) and pd.isna(v)): return "&nbsp;"
        return f"{int(v):,}" if isinstance(v, (int, float)) else str(v)

    top_rows_list = []
    for i in range(5):
        lr, cr = card["ly_rows"][i], card["cur_rows"][i]
        top_rows_list.append(
            f"<tr><td style='border:1px solid #333; padding:2px 4px; text-align:center; font-size:11px;'>{fmt_row_v(lr, 'sku')}</td>"
            f"<td style='border:1px solid #333; padding:2px 4px; text-align:right; font-size:11px;'>{fmt_row_v(lr, 'price')}</td>"
            f"<td style='border:1px solid #333; padding:2px 4px; text-align:right; font-size:11px;'>{fmt_row_v(lr, 'qty')}</td>"
            f"<td style='border:1px solid #333; padding:2px 4px; text-align:right; font-size:11px;'>{fmt_row_v(lr, 'amt')}</td>"
            f"<td style='border:1px solid #333; border-left:2px solid #111; padding:2px 4px; text-align:center; font-size:11px;'>{fmt_row_v(cr, 'sku')}</td>"
            f"<td style='border:1px solid #333; padding:2px 4px; text-align:right; font-size:11px;'>{fmt_row_v(cr, 'price')}</td>"
            f"<td style='border:1px solid #333; padding:2px 4px; text-align:right; font-size:11px;'>{fmt_row_v(cr, 'qty')}</td>"
            f"<td style='border:1px solid #333; padding:2px 4px; text-align:right; font-size:11px;'>{fmt_row_v(cr, 'amt')}</td></tr>"
        )
    act_html = str(card.get("activity_text", "")).replace("\n", "<br>") if card.get("activity_text") else "&nbsp;"
    plan_html = str(card.get("plan_text", "")).replace("\n", "<br>") if card.get("plan_text") else "&nbsp;"

    next_rows_list = []
    for i in range(5):
        nr = card["next_rows"][i]
        extra_td = f"<td colspan='4' rowspan='7' style='border:1px solid #333; border-left:2px solid #111; padding:4px 6px; vertical-align:top; font-size:11px; text-align:left;'>{act_html}</td>" if i == 0 else ""
        next_rows_list.append(
            f"<tr><td style='border:1px solid #333; padding:2px 4px; text-align:center; font-size:11px;'>{fmt_row_v(nr, 'sku')}</td>"
            f"<td style='border:1px solid #333; padding:2px 4px; text-align:right; font-size:11px;'>{fmt_row_v(nr, 'price')}</td>"
            f"<td style='border:1px solid #333; padding:2px 4px; text-align:right; font-size:11px;'>{fmt_row_v(nr, 'qty')}</td>"
            f"<td style='border:1px solid #333; padding:2px 4px; text-align:right; font-size:11px;'>{fmt_row_v(nr, 'amt')}</td>{extra_td}</tr>"
        )

    raw_html = f"""
<div style="border:1px solid #94a3b8; padding:10px; background:#ffffff; margin-bottom:14px; font-family:'Yu Gothic','Meiryo',sans-serif; color:#111;">
<div style="display:flex; justify-content:space-between; align-items:baseline; margin-bottom:6px;">
<span style="font-size:15px; font-weight:bold; text-decoration:underline;">{card['cust_name']}</span>
<span style="font-size:12px; color:#333; margin-right:35%;">{card['cust_id']}</span>
</div>
<table style="width:100%; border-collapse:collapse; table-layout:fixed; margin-bottom:2px;">
<thead>
<tr style="background:#f8fafc;">
<th colspan="4" style="border:1px solid #333; padding:2px 5px; text-align:left; font-size:11px;">{card['ly_label']}</th>
<th colspan="4" style="border:1px solid #333; border-left:2px solid #111; padding:2px 5px; text-align:left; font-size:11px;">{card['cur_label']}</th>
</tr>
<tr style="background:#f1f5f9;">
<th style="border:1px solid #333; padding:2px; font-size:11px; width:14%;">品番</th><th style="border:1px solid #333; padding:2px; font-size:11px; width:10%;">単価</th><th style="border:1px solid #333; padding:2px; font-size:11px; width:10%;">数量</th><th style="border:1px solid #333; padding:2px; font-size:11px; width:16%;">金額</th>
<th style="border:1px solid #333; border-left:2px solid #111; padding:2px; font-size:11px; width:14%;">品番</th><th style="border:1px solid #333; padding:2px; font-size:11px; width:10%;">単価</th><th style="border:1px solid #333; padding:2px; font-size:11px; width:10%;">数量</th><th style="border:1px solid #333; padding:2px; font-size:11px; width:16%;">金額</th>
</tr>
</thead>
<tbody>
{''.join(top_rows_list)}
<tr>
<td colspan="3" style="border:1px solid #333; padding:2px 5px; text-align:left; font-size:11px;">その他</td><td style="border:1px solid #333; padding:2px 4px; text-align:right; font-size:11px;">{card['ly_other']:,}</td>
<td colspan="3" style="border:1px solid #333; border-left:2px solid #111; padding:2px 5px; text-align:left; font-size:11px;">その他</td><td style="border:1px solid #333; padding:2px 4px; text-align:right; font-size:11px;">{card['cur_other']:,}</td>
</tr>
<tr style="font-weight:bold; background:#f8fafc;">
<td colspan="3" style="border:1px solid #333; padding:2px 5px; text-align:left; font-size:11px;">合計</td><td style="border:1px solid #333; padding:2px 4px; text-align:right; font-size:11px;">{card['ly_total']:,}</td>
<td colspan="3" style="border:1px solid #333; border-left:2px solid #111; padding:2px 5px; text-align:left; font-size:11px;">合計</td><td style="border:1px solid #333; padding:2px 4px; text-align:right; font-size:11px;">{card['cur_total']:,}</td>
</tr>
</tbody>
</table>
<div style="text-align:right; font-size:11px; font-weight:bold; margin-bottom:8px; padding-right:4px;">昨対比 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; {card['yoy_rate']}</div>
<table style="width:100%; border-collapse:collapse; table-layout:fixed;">
<thead>
<tr style="background:#f8fafc;">
<th colspan="4" style="border:1px solid #333; padding:2px 5px; text-align:left; font-size:11px;">{card['next_label']}</th>
<th colspan="4" style="border:1px solid #333; border-left:2px solid #111; padding:2px 5px; text-align:left; font-size:11px;">この1ヶ月の主な営業活動</th>
</tr>
<tr style="background:#f1f5f9;">
<th style="border:1px solid #333; padding:2px; font-size:11px; width:14%;">品番</th><th style="border:1px solid #333; padding:2px; font-size:11px; width:10%;">単価</th><th style="border:1px solid #333; padding:2px; font-size:11px; width:10%;">数量</th><th style="border:1px solid #333; padding:2px; font-size:11px; width:16%;">金額</th>
<th colspan="4" style="border:1px solid #333; border-left:2px solid #111; padding:0; height:2px;"></th>
</tr>
</thead>
<tbody>
{''.join(next_rows_list)}
<tr><td colspan="3" style="border:1px solid #333; padding:2px 5px; text-align:left; font-size:11px;">その他</td><td style="border:1px solid #333; padding:2px 4px; text-align:right; font-size:11px;">{card['next_other']:,}</td></tr>
<tr style="font-weight:bold; background:#f8fafc;"><td colspan="3" style="border:1px solid #333; padding:2px 5px; text-align:left; font-size:11px;">合計</td><td style="border:1px solid #333; padding:2px 4px; text-align:right; font-size:11px;">{card['next_total']:,}</td></tr>
<tr>
<td colspan="4" style="border:1px solid #333; padding:4px 6px; vertical-align:top; height:68px; font-size:11px;"><div style="color:#64748b; font-size:10px;">【今後の予定】</div><div>{plan_html}</div></td>
<td colspan="4" style="border:1px solid #333; border-left:2px solid #111; padding:4px 6px; vertical-align:top; font-size:11px;">
<div style="display:flex; justify-content:space-between;"><span>{card['prev_fy_label']}</span><span>{card['prev_fy_total']:,}</span></div>
<div style="display:flex; justify-content:space-between;"><span>{card['as_of_label']}</span><span>{card['next_total']:,}</span></div>
<div style="display:flex; justify-content:space-between; font-weight:bold;"><span>昨対進捗</span><span>{card['next_yoy_rate']}</span></div>
</td>
</tr>
</tbody>
</table>
</div>
"""
    return "\n".join(line.strip() for line in raw_html.splitlines() if line.strip())

def generate_meeting_comparison_excel(cards_data, tgt_ym):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "上位社売上比較シート"
    font_b = Font(name="Yu Gothic", size=10, bold=True)
    font_m = Font(name="Yu Gothic", size=9)
    thin = Side(style='thin', color='333333')
    bd = Border(left=thin, right=thin, top=thin, bottom=thin)
    fill_h = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")

    for col_idx, width_val in {1: 15, 2: 9, 3: 9, 4: 14, 5: 15, 6: 9, 7: 9, 8: 14, 9: 4, 10: 15, 11: 9, 12: 9, 13: 14, 14: 15, 15: 9, 16: 9, 17: 14}.items():
        ws.column_dimensions[get_column_letter(col_idx)].width = width_val

    for idx, c in enumerate(cards_data):
        grid_r, grid_c = idx // 2, idx % 2
        base_r, base_c = 2 + grid_r * 26, 1 + grid_c * 9
        ws.cell(row=base_r, column=base_c, value=c["cust_name"]).font = Font(name="Yu Gothic", size=11, bold=True, underline="single")
        ws.cell(row=base_r, column=base_c + 2, value=c["cust_id"]).font = font_m

        r_top = base_r + 1
        ws.cell(row=r_top, column=base_c, value=c["ly_label"]).font = font_b
        ws.cell(row=r_top, column=base_c + 4, value=c["cur_label"]).font = font_b
        for h_i, h_t in enumerate(["品番", "単価", "数量", "金額", "品番", "単価", "数量", "金額"]):
            cell = ws.cell(row=r_top + 1, column=base_c + h_i, value=h_t)
            cell.font = font_b; cell.border = bd; cell.fill = fill_h; cell.alignment = Alignment(horizontal="center")

        for i in range(5):
            curr_row = r_top + 2 + i
            lr, cr = c["ly_rows"][i], c["cur_rows"][i]
            l_empty = not str(lr.get("sku", "")).strip() and int(lr.get("amt") or 0) == 0
            c_empty = not str(cr.get("sku", "")).strip() and int(cr.get("amt") or 0) == 0
            vals = [
                None if l_empty else lr["sku"], None if l_empty else lr["price"], None if l_empty else lr["qty"], None if l_empty else lr["amt"],
                None if c_empty else cr["sku"], None if c_empty else cr["price"], None if c_empty else cr["qty"], None if c_empty else cr["amt"]
            ]
            for v_i, v_val in enumerate(vals):
                cell = ws.cell(row=curr_row, column=base_c + v_i, value=v_val)
                cell.font = font_m; cell.border = bd
                if v_i in [1, 2, 3, 5, 6, 7] and isinstance(v_val, (int, float)): cell.number_format = '#,##0'

        r_oth, r_tot = r_top + 7, r_top + 8
        ws.cell(row=r_oth, column=base_c, value="その他").font = font_m
        ws.cell(row=r_oth, column=base_c + 3, value=c["ly_other"]).number_format = '#,##0'
        ws.cell(row=r_oth, column=base_c + 4, value="その他").font = font_m
        ws.cell(row=r_oth, column=base_c + 7, value=c["cur_other"]).number_format = '#,##0'
        ws.cell(row=r_tot, column=base_c, value="合計").font = font_b
        ws.cell(row=r_tot, column=base_c + 3, value=c["ly_total"]).number_format = '#,##0'
        ws.cell(row=r_tot, column=base_c + 4, value="合計").font = font_b
        ws.cell(row=r_tot, column=base_c + 7, value=c["cur_total"]).number_format = '#,##0'
        for h_i in range(8):
            ws.cell(row=r_oth, column=base_c + h_i).border = bd
            ws.cell(row=r_tot, column=base_c + h_i).border = bd

        ws.cell(row=r_tot + 1, column=base_c + 5, value="昨対比").font = font_b
        c_yoy = ws.cell(row=r_tot + 1, column=base_c + 7, value=c["yoy_rate"])
        c_yoy.font = font_b; c_yoy.alignment = Alignment(horizontal="right")

        r_nx = r_tot + 2
        ws.cell(row=r_nx, column=base_c, value=c["next_label"]).font = font_b
        ws.cell(row=r_nx, column=base_c + 4, value="この1ヶ月の主な営業活動").font = font_b
        for h_i, h_t in enumerate(["品番", "単価", "数量", "金額"]):
            cell = ws.cell(row=r_nx + 1, column=base_c + h_i, value=h_t)
            cell.font = font_b; cell.border = bd; cell.fill = fill_h; cell.alignment = Alignment(horizontal="center")

        for i in range(5):
            curr_row = r_nx + 2 + i
            nr = c["next_rows"][i]
            n_empty = not str(nr.get("sku", "")).strip() and int(nr.get("amt") or 0) == 0
            n_vals = [None, None, None, None] if n_empty else [nr["sku"], nr["price"], nr["qty"], nr["amt"]]
            for v_i, v_val in enumerate(n_vals):
                cell = ws.cell(row=curr_row, column=base_c + v_i, value=v_val)
                cell.font = font_m; cell.border = bd
                if v_i in [1, 2, 3] and isinstance(v_val, (int, float)): cell.number_format = '#,##0'

        ws.merge_cells(start_row=r_nx + 1, start_column=base_c + 4, end_row=r_nx + 8, end_column=base_c + 7)
        act_c = ws.cell(row=r_nx + 1, column=base_c + 4, value=c.get("activity_text", ""))
        act_c.font = font_m; act_c.alignment = Alignment(vertical="top", wrap_text=True)
        for rr in range(r_nx + 1, r_nx + 9):
            for cc in range(base_c + 4, base_c + 8): ws.cell(row=rr, column=cc).border = bd

        r_nx_oth, r_nx_tot, r_bot = r_nx + 7, r_nx + 8, r_nx + 9
        ws.cell(row=r_nx_oth, column=base_c, value="その他").font = font_m
        ws.cell(row=r_nx_oth, column=base_c + 3, value=c["next_other"]).number_format = '#,##0'
        ws.cell(row=r_nx_tot, column=base_c, value="合計").font = font_b
        ws.cell(row=r_nx_tot, column=base_c + 3, value=c["next_total"]).number_format = '#,##0'
        for h_i in range(4):
            ws.cell(row=r_nx_oth, column=base_c + h_i).border = bd
            ws.cell(row=r_nx_tot, column=base_c + h_i).border = bd

        ws.merge_cells(start_row=r_bot, start_column=base_c, end_row=r_bot + 2, end_column=base_c + 3)
        pln_c = ws.cell(row=r_bot, column=base_c, value=f"【今後の予定】\n{c.get('plan_text', '')}")
        pln_c.font = font_m; pln_c.alignment = Alignment(vertical="top", wrap_text=True)
        for rr in range(r_bot, r_bot + 3):
            for cc in range(base_c, base_c + 4): ws.cell(row=rr, column=cc).border = bd

        ws.cell(row=r_bot, column=base_c + 4, value=c["prev_fy_label"]).font = font_m
        ws.cell(row=r_bot, column=base_c + 7, value=c["prev_fy_total"]).number_format = '#,##0'
        ws.cell(row=r_bot + 1, column=base_c + 4, value=c["as_of_label"]).font = font_m
        ws.cell(row=r_bot + 1, column=base_c + 7, value=c["next_total"]).number_format = '#,##0'
        ws.cell(row=r_bot + 2, column=base_c + 4, value="昨対進捗").font = font_b
        c_nyoy = ws.cell(row=r_bot + 2, column=base_c + 7, value=c["next_yoy_rate"])
        c_nyoy.font = font_b; c_nyoy.alignment = Alignment(horizontal="right")
        for rr in range(r_bot, r_bot + 3):
            for cc in range(base_c + 4, base_c + 8): ws.cell(row=rr, column=cc).border = bd

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()

def generate_meeting_top_cust_excel(res_df, title_str):
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine='openpyxl') as writer:
        res_df.to_excel(writer, index=False, sheet_name="上位客先実績")
    return out.getvalue()

# ========================================================
# ★ データ読込 & マスタ初期化（ファイル変化シグネチャを渡して自動更新）
# ========================================================
df, file_count, cost_master_dict = load_and_prep_data(_get_data_files_signature())
member_p_path, member_p_mtime = get_member_price_file_info()
member_cust_sku_map, member_default_min_map = load_member_price_master_v2(member_p_path, member_p_mtime)
lot_tpl_path, lot_tpl_mtime = get_lotmember_template_file_info()
lotmember_sku_meta_dict = load_lotmember_sku_master(lot_tpl_path, lot_tpl_mtime)
grp_path, grp_mtime = get_member_group_file_info()
member_group_master_dict = load_member_group_master(grp_path, grp_mtime)
jan_path, jan_mtime = get_jan_file_info()
jan_master_dict, jan_title_dict, jan_std_price_dict = load_jan_master_v2(jan_path, jan_mtime)
cust_path, cust_mtime = get_customer_file_info()
customer_master_df = load_customer_master(cust_path, cust_mtime)
target_path, target_mtime = get_target_file_info()
staff_target_master = load_staff_target_master(target_path, target_mtime)
hanko_bytes, hanko_file_path = get_hanko_bytes()

ao_info, ng_info, kk_info, hc_info = get_inventory_files_info()
logi_inv_map, yayoi_kakuho_map, yayoi_kakuho_df, hacchu_master_map = load_inventory_and_order_data(ao_info, ng_info, kk_info, hc_info)

raw_prod_title_map = {}
if "元商品名" in df.columns and "商品コード" in df.columns:
    raw_prod_title_map = df.dropna(subset=["商品コード", "元商品名"]).drop_duplicates(subset=["商品コード"]).set_index("商品コード")["元商品名"].to_dict()
_m_p_raw, _m_t_raw = get_order_r_master_info()
if _m_p_raw:
    _m_data_raw = _load_order_r_all_sheets(_m_p_raw, _m_t_raw)
    if _m_data_raw and _m_data_raw.get("raw_title_dict"):
        raw_prod_title_map.update(_m_data_raw["raw_title_dict"])
for k_t, v_t in (jan_title_dict or {}).items():
    sk_trim = str(k_t).strip()
    if sk_trim not in raw_prod_title_map:
        raw_prod_title_map[sk_trim] = str(v_t).strip()

col_status, col_mob_toggle, col_upload = st.columns([2.4, 1.0, 1.8])
with col_status:
    _yayoi_only = df[df["得意先名"] != "アマゾンセラー"] if not df.empty and "得意先名" in df.columns else df
    _max_dt_str = _yayoi_only["日付_dt"].max().strftime("%Y/%m/%d") if (not _yayoi_only.empty and "日付_dt" in _yayoi_only.columns and _yayoi_only["日付_dt"].notna().any()) else "未取得"
    _diag_msg = st.session_state.get("_gdrive_diag", "同期ステータス未実行")
    st.caption(f"📊 **反映済データ最新日: {_max_dt_str}** （全 {len(df):,} 行）\n\n🔍 {_diag_msg}")
with col_mob_toggle:
    is_mobile_view = st.toggle("📱 スマホ最適化表示", value=False, key="mobile_opt_toggle")
with col_upload:
    with st.expander("📥 データ追加・在庫更新・バックアップ", expanded=False):
        if st.button("☁️ Googleドライブの最新ファイルを今すぐ同期", use_container_width=True):
            sync_all_from_gdrive(force=True)
            st.cache_data.clear()
            st.rerun()

        up_tab1, up_tab2, up_tab3 = st.tabs([
            "📄 売上CSV追加",
            "📦 在庫・確保更新",
            "💾 バックアップ"
        ])

        with up_tab1:
            up_csv_type = st.radio("種類", ["弥生売上CSV", "アマゾンセラーCSV"], horizontal=True, key="up_csv_type_radio")
            amz_ym_input = ""
            if up_csv_type == "アマゾンセラーCSV":
                amz_ym_input = st.text_input("対象年月 (ファイル名に年月が無い場合に適用)", value=datetime.now().strftime("%Y年%m月"), key="amz_ym_input_box")

            uploaded_files = st.file_uploader("CSVファイルを選択 (複数可)", type=["csv"], accept_multiple_files=True, key="csv_uploader_widget")
            if uploaded_files and st.button("＋ 売上データベースに追加反映", type="primary", use_container_width=True):
                added_yayoi, added_amz = [], []
                for uf in uploaded_files:
                    if up_csv_type == "弥生売上CSV":
                        raw_d = _read_single_csv(uf, explicit_name=uf.name)
                        if raw_d is not None and not raw_d.empty:
                            added_yayoi.append(_parse_raw_df_to_standard(raw_d))
                    else:
                        has_m_in_name = bool(re.search(r'(20\d{2})[年/\-_]*(\d{1,2})月?', uf.name))
                        amz_d = _read_amazon_seller_csv(uf, explicit_name=uf.name, override_ym=(None if has_m_in_name else amz_ym_input))
                        if amz_d is not None and not amz_d.empty:
                            added_amz.append(amz_d)

                base_df = df.copy()
                if added_amz:
                    amz_concat = pd.concat(added_amz, ignore_index=True)
                    replace_yms = set(amz_concat["年月"].dropna().unique())
                    base_df = base_df[~((base_df["得意先名"] == "アマゾンセラー") & (base_df["年月"].isin(replace_yms)))].copy()
                    added_yayoi.append(amz_concat)

                if added_yayoi:
                    new_merged = pd.concat([base_df] + added_yayoi, ignore_index=True)
                    name_lookup = dict(jan_title_dict or {})
                    valid_names = new_merged[(new_merged["商品名"].notna()) & (new_merged["商品名"].astype(str) != new_merged["商品コード"].astype(str)) & (new_merged["商品名"].astype(str).str.strip() != "")]
                    if not valid_names.empty:
                        name_lookup.update(valid_names.drop_duplicates(subset=["商品コード"]).set_index("商品コード")["商品名"].to_dict())
                    new_merged = finalize_sales_df(new_merged, cost_master_dict, name_lookup)

                    try:
                        save_parquet_and_sync(new_merged, CACHE_FILE, "sales_cache.parquet")
                        st.cache_data.clear()
                        st.success("データを追加・保存しました！")
                        st.rerun()
                    except Exception as e:
                        st.error(f"保存エラー: {e}")

        with up_tab2:
            st.caption("ロジザード在庫csv（青山・名古屋）、在庫確保.xlsx、発注検討表.xlsx、order_r（マスター）.xlsxを最新に上書きします。")
            inv_up_files = st.file_uploader("在庫・確保・マスターファイルを選択", type=["csv", "xlsx", "xls"], accept_multiple_files=True, key="inv_uploader_widget")
            if inv_up_files and st.button("📁 データを最新に更新", type="primary", use_container_width=True):
                for iuf in inv_up_files:
                    fn = iuf.name
                    save_name = fn
                    if "青山" in fn and fn.endswith(".csv"): save_name = "ロジザード在庫 (青山).csv"
                    elif "名古屋" in fn and fn.endswith(".csv"): save_name = "ロジザード在庫 (名古屋).csv"
                    elif "在庫確保" in fn: save_name = "在庫確保.xlsx"
                    elif "発注検討" in fn: save_name = "発注検討表.xlsx"
                    elif "order_r" in fn.lower() or "マスター" in fn: save_name = "order_r（マスター）.xlsx"
                    for target_d in [GDRIVE_SYNC_DIR, BASE_DIR]:
                        with open(os.path.join(target_d, save_name), "wb") as f_out:
                            f_out.write(iuf.getbuffer())
                    upload_file_to_gdrive(os.path.join(BASE_DIR, save_name), save_name)
                st.cache_data.clear()
                st.success("✅ データを更新しました！")
                time.sleep(0.4)
                st.rerun()

        with up_tab3:
            if os.path.exists(CACHE_FILE):
                with open(CACHE_FILE, "rb") as f_pq:
                    st.download_button("💾 ① 売上DB (sales_cache.parquet) DL", data=f_pq.read(), file_name="sales_cache.parquet", mime="application/octet-stream", use_container_width=True)
            quotes_now = load_saved_quotes()
            st.download_button("💾 ② 保存済み見積 (saved_estimates.json) DL", data=json.dumps(quotes_now, ensure_ascii=False, indent=2).encode("utf-8"), file_name="saved_estimates.json", mime="application/json", use_container_width=True)
            meeting_now = load_saved_meeting_sheet_data()
            st.download_button("💾 ③ 会議シート (saved_meeting_sheet.json) DL", data=json.dumps(meeting_now, ensure_ascii=False, indent=2).encode("utf-8"), file_name="saved_meeting_sheet.json", mime="application/json", use_container_width=True)
            restore_json = st.file_uploader("📂 JSONバックアップを復元", type=["json"], key="restore_json_up")
            if restore_json and st.button("♻️ JSONを復元反映", use_container_width=True):
                try:
                    loaded_j = json.loads(restore_json.getvalue().decode("utf-8"))
                    if isinstance(loaded_j, list):
                        save_quotes_list(loaded_j)
                        st.session_state.saved_estimates = loaded_j
                        st.success("✅ 保存済み見積を復元しました！")
                    elif isinstance(loaded_j, dict):
                        save_meeting_sheet_data(loaded_j)
                        st.success("✅ 営業会議シートを復元しました！")
                    time.sleep(0.4)
                    st.rerun()
                except Exception as e:
                    st.error(f"復元エラー: {e}")

if df.empty:
    st.error("売上データが見つかりませんでした。右上の「📥 データ追加・在庫更新・バックアップ」から「☁️ Googleドライブの最新ファイルを今すぐ同期」を押すか、GitHubの `sales_cache.parquet` を確認してください。")
    st.stop()

valid_ym = sorted([str(x).strip() for x in df["年月"].dropna().unique() if str(x).strip() not in ["", "nan", "None"]])
latest_ym = valid_ym[-1] if valid_ym else "2026年01月"
desc_valid_ym = sorted(valid_ym, reverse=True)

valid_dts = df["日付_dt"].dropna()
min_date = valid_dts.min().date() if not valid_dts.empty else datetime(2022, 1, 1).date()
max_date = valid_dts.max().date() if not valid_dts.empty else datetime.now().date()

curr_y = int(latest_ym[:4]) if len(latest_ym) >= 4 else 2026
curr_m = int(latest_ym[5:7]) if len(latest_ym) >= 7 else 1
fiscal_year = curr_y if curr_m >= 7 else curr_y - 1

VIEW_MODES = [
    '🏢 得意先別ランキング',
    '📦 商品別売上ランキング',
    '🧮 簡易見積＆利益シミュレーター',
    '📊 営業会議用：上位社 売上比較シート',
    '📑 営業会議用：上位客先別売上実績',
    '🛒 アマゾン日別注文数＆利益計算出力',
    '🏷️ ランク価格マスタ一覧',
    '👤 担当者別 売上＆目標達成分析',
    '🏭 在庫・確保数＆発注検討アラート',
    '✨ 新商品売上分析（直近1年）',
    '📈 商品別 前年実績比較',
    '📊 得意先別 前年実績比較'
]
col_top1, col_top2 = st.columns([1.6, 2])
with col_top1:
    view_type = st.selectbox("表示モード", VIEW_MODES)
with col_top2:
    if view_type == '🧮 簡易見積＆利益シミュレーター':
        st.write("")
        jan_status = f"JAN: **{len(jan_master_dict):,} 件**" if jan_master_dict else "⚠️ JAN未検出"
        hanko_status = "角印: **読込完了**" if hanko_bytes else "⚠️ 印鑑未検出"
        grp_status = f"グループ管理: **{len(member_group_master_dict):,} 件**" if member_group_master_dict else "⚠️ グループ未検出"
        st.info(f"📁 原価: **{len(cost_master_dict):,} 件** | {grp_status} | {jan_status} | {hanko_status}")
    elif view_type == '🏭 在庫・確保数＆発注検討アラート':
        st.write("")
        st.info(f"📦 ロジザード登録: **{len(logi_inv_map):,} 品番** | 🔒 弥生確保: **{len(yayoi_kakuho_map):,} 品番** | 📋 発注検討表: **{len(hacchu_master_map):,} 品番**")
    else:
        quick_choice = st.radio("期間ショートカット", ['当月', '前月', '当期累計', '前期累計', '全データ累計', '手動指定'], horizontal=True, key="period_shortcut_radio_top")

# ========================================================
# ★ 🛒 アマゾン日別注文数＆利益計算出力（セラー日次CSV ＆ 弥生ベンダー自動連動）
# ========================================================
@st.cache_data(show_spinner=False)
def load_amazon_sku_and_profit_masters():
    asin_to_sku = {}
    sku_to_clean_name = {}
    sku_f, _ = find_file_multidir("*SKU*.xls*")
    if sku_f and os.path.exists(sku_f):
        try:
            sdf = pd.read_excel(sku_f)
            for _, r in sdf.iterrows():
                c_asin = str(r.get("(子) ASIN", r.iloc[1] if len(r) > 1 else "")).strip()
                p_asin = str(r.iloc[0]).strip() if len(r) > 0 else ""
                sk = str(r.get("SKU", r.iloc[2] if len(r) > 2 else "")).strip()
                nm = str(r.get("商品名", r.iloc[3] if len(r) > 3 else "")).strip()
                if sk and sk not in ["nan", "None"]:
                    if c_asin and c_asin not in ["nan", "None"]: asin_to_sku[c_asin] = (sk, nm)
                    if p_asin and p_asin not in ["nan", "None"] and p_asin not in asin_to_sku: asin_to_sku[p_asin] = (sk, nm)
                    if sk not in sku_to_clean_name and nm not in ["nan", "None", ""]: sku_to_clean_name[sk] = nm
        except Exception: pass

    seller_prof_map = {}
    vendor_prof_map = {}
    prof_f, _ = find_file_multidir("*アマゾン利益計算*.xls*")
    if prof_f and os.path.exists(prof_f):
        try:
            xls_p = pd.ExcelFile(prof_f)
            if "SKU" in xls_p.sheet_names and not asin_to_sku:
                sdf2 = pd.read_excel(prof_f, sheet_name="SKU")
                for _, r in sdf2.iterrows():
                    c_asin = str(r.get("(子) ASIN", "")).strip()
                    sk = str(r.get("SKU", "")).strip()
                    nm = str(r.get("商品名", "")).strip()
                    if sk and sk != "nan":
                        if c_asin and c_asin != "nan": asin_to_sku[c_asin] = (sk, nm)
                        if sk not in sku_to_clean_name: sku_to_clean_name[sk] = nm
            if "アマゾン出荷商品" in xls_p.sheet_names:
                adf = pd.read_excel(prof_f, sheet_name="アマゾン出荷商品")
                for _, r in adf.iterrows():
                    sk = str(r.get("商品", "")).strip()
                    if not sk or sk == "nan": continue
                    seller_prof_map[sk] = {
                        "unit_profit": clean_num(r.get("粗利", 0)),
                        "margin_rate": clean_num(r.get("粗利率", 0)),
                        "asahi_price": clean_num(r.get("旭電機の売値", 0)),
                        "name": str(r.get("商品名", "")).strip()
                    }
            if "ベンダー" in xls_p.sheet_names:
                vdf = pd.read_excel(prof_f, sheet_name="ベンダー")
                for _, r in vdf.iterrows():
                    sk = str(r.get("SKU", "")).strip()
                    if not sk or sk == "nan": continue
                    vendor_prof_map[sk] = {
                        "sell_price": clean_num(r.get("売値", 0)),
                        "cost": clean_num(r.get("原価", 0)),
                        "margin_rate": clean_num(r.get("利益率", 0)),
                        "unit_profit": clean_num(r.get("粗利", 0)),
                        "name": str(r.get("商品名", "")).strip()
                    }
        except Exception: pass

    return asin_to_sku, sku_to_clean_name, seller_prof_map, vendor_prof_map

def parse_month_day_from_csv_name(fname, default_month=None):
    bname = os.path.basename(fname)
    m_ymd = re.search(r'20\d{2}[年/\-_](\d{1,2})[月/\-_](\d{1,2})', bname)
    if m_ymd: return int(m_ymd.group(1)), int(m_ymd.group(2))
    m_md = re.search(r'(?<!\d)(\d{1,2})[月/\-_](\d{1,2})(?!\d)', bname)
    if m_md: return int(m_md.group(1)), int(m_md.group(2))
    m_d = re.search(r'^(\d{1,2})\.csv$', bname, re.IGNORECASE)
    if m_d and default_month: return int(default_month), int(m_d.group(1))
    return None, None

def read_seller_daily_business_report(f_obj):
    for enc in ["utf-8-sig", "utf-8", "cp932", "shift_jis"]:
        try:
            if hasattr(f_obj, "seek"): f_obj.seek(0)
            d = pd.read_csv(f_obj, encoding=enc, low_memory=False)
            if d.empty: continue
            c_child = next((c for c in d.columns if "子" in str(c) and "ASIN" in str(c)), None)
            if not c_child: c_child = next((c for c in d.columns if "ASIN" in str(c)), d.columns[1] if len(d.columns) > 1 else d.columns[0])
            c_qty = next((c for c in d.columns if "注文された商品点数" in str(c) or "注文点数" in str(c)), None)
            c_sales = next((c for c in d.columns if "注文商品の売上額" in str(c) and "B2B" not in str(c)), None)
            c_sess = next((c for c in d.columns if "セッション数" in str(c) and "B2B" not in str(c)), None)
            c_pv = next((c for c in d.columns if "ページビュー" in str(c) and "B2B" not in str(c)), None)
            c_buybox = next((c for c in d.columns if "おすすめ出品" in str(c) and "B2B" not in str(c)), None)
            c_title = next((c for c in d.columns if "タイトル" in str(c) or "商品名" in str(c)), None)

            rows = []
            for _, r in d.iterrows():
                asin = str(r.get(c_child, "")).strip()
                if not asin or asin in ["nan", "None"]: continue
                qty = int(clean_num(r.get(c_qty, 0))) if c_qty else 0
                sales = int(clean_num(r.get(c_sales, 0))) if c_sales else 0
                sess = int(clean_num(r.get(c_sess, 0))) if c_sess else 0
                pv = int(clean_num(r.get(c_pv, 0))) if c_pv else 0
                bb_raw = str(r.get(c_buybox, "0")).strip() if c_buybox else "0"
                bb_val = clean_num(bb_raw) / 100.0 if "%" in bb_raw or clean_num(bb_raw) > 1.0 else clean_num(bb_raw)
                rows.append({
                    "asin": asin, "raw_title": str(r.get(c_title, "")).strip() if c_title else "",
                    "qty": qty, "sales": sales, "sess": sess, "pv": pv, "buybox": bb_val, "raw_row": r.to_dict()
                })
            return pd.DataFrame(rows), d
        except Exception: continue
    return pd.DataFrame(), pd.DataFrame()

if view_type == '🛒 アマゾン日別注文数＆利益計算出力':
    asin_map, sku_name_map, seller_prof_map, vendor_prof_map = load_amazon_sku_and_profit_masters()
    code_to_name_local = df.dropna(subset=["商品コード", "商品名"]).drop_duplicates(subset=["商品コード"]).set_index("商品コード")["商品名"].to_dict() if not df.empty else {}
    for k_t, v_t in (jan_title_dict or {}).items():
        if str(k_t).strip() not in code_to_name_local: code_to_name_local[str(k_t).strip()] = str(v_t).strip()

    st.markdown("#### 🛒 アマゾン日別注文数 ＆『アマゾン利益計算』シート自動作成")
    st.caption("※セラーの日次CSV（例：`11-2.csv`）と弥生の売上データから、`注文数`（セラー）・`ベンダー注文数`（ベンダー）および日別31シート＋合計のExcelをワンクリックで作成します。")

    col_ym_sel, col_up_daily = st.columns([1.2, 2.8])
    with col_ym_sel:
        amz_target_ym = st.selectbox("📅 集計対象の年月を選択", desc_valid_ym, index=0, key="amz_daily_ym_sel")
        tgt_y_int, tgt_m_int = int(amz_target_ym[:4]), int(amz_target_ym[5:7])
        st.info(f"📁 `SKU.xlsx` 登録: **{len(sku_name_map):,} 品番**\n\n📊 利益計算マスタ: **読込済**")

    with col_up_daily:
        up_daily_csvs = st.file_uploader(f"📥 アマゾンセラー日次CSV（例: `{tgt_m_int}-1.csv`, `{tgt_m_int}-2.csv` ...）をドラッグ＆ドロップ（複数日まとめ入れOK）", type=["csv"], accept_multiple_files=True, key="amz_daily_csv_uploader")
        if up_daily_csvs:
            for u_f in up_daily_csvs:
                save_d_p = os.path.join(AMZ_DAILY_STORE_DIR, u_f.name)
                with open(save_d_p, "wb") as f_out: f_out.write(u_f.getbuffer())
            st.success(f"✅ {len(up_daily_csvs)} 件の日次CSVを読み込みました！")

    daily_csv_map, raw_daily_df_map, search_csv_paths = {}, {}, []
    for s_dir in [AMZ_DAILY_STORE_DIR, GDRIVE_SYNC_DIR, BASE_DIR, BASE_SHARE]:
        if os.path.exists(s_dir): search_csv_paths.extend(glob.glob(os.path.join(s_dir, "*.csv")))
    for c_path in sorted(list(set(search_csv_paths))):
        m_num, d_num = parse_month_day_from_csv_name(c_path, default_month=None)
        if m_num == tgt_m_int and d_num and 1 <= d_num <= 31:
            parsed_d, raw_d = read_seller_daily_business_report(c_path)
            if not parsed_d.empty: daily_csv_map[d_num], raw_daily_df_map[d_num] = parsed_d, raw_d

    seller_sku_days, seller_metrics = {}, {}
    for d_num, p_df in daily_csv_map.items():
        for _, r in p_df.iterrows():
            asin = r["asin"]
            sk, clean_nm = asin_map.get(asin, ("", ""))
            if not sk:
                m_sk = re.search(r'([A-Z]{2,4}-\d{2,5}[A-Z0-9]*)', r["raw_title"])
                sk = m_sk.group(1) if m_sk else asin
            if not clean_nm: clean_nm = sku_name_map.get(sk, code_to_name_local.get(sk, r["raw_title"][:25]))
            if sk not in seller_sku_days:
                seller_sku_days[sk] = {"asin": asin, "name": clean_nm, "days": {d: 0 for d in range(1, 32)}}
                seller_metrics[sk] = {"sales": 0, "sess": 0, "pv": 0, "bb_sum": 0.0, "bb_cnt": 0}
            seller_sku_days[sk]["days"][d_num] += int(r["qty"])
            seller_metrics[sk]["sales"] += int(r["sales"])
            seller_metrics[sk]["sess"] += int(r["sess"])
            seller_metrics[sk]["pv"] += int(r["pv"])
            if r["buybox"] > 0:
                seller_metrics[sk]["bb_sum"] += float(r["buybox"])
                seller_metrics[sk]["bb_cnt"] += 1

    seller_rows = []
    for sk, info in seller_sku_days.items():
        tot_q = sum(info["days"].values())
        m_info, p_info = seller_metrics[sk], seller_prof_map.get(sk, {})
        u_prof, m_rate, asahi_p = p_info.get("unit_profit", 0.0), p_info.get("margin_rate", 0.0), p_info.get("asahi_price", 0.0)
        avg_bb = round(m_info["bb_sum"] / m_info["bb_cnt"], 4) if m_info["bb_cnt"] > 0 else 0.0
        row_dict = {"（子）ASIN": info["asin"], "タイトル": sk, "商品名": info["name"]}
        for d in range(1, 32): row_dict[f"{d}日"] = info["days"][d] if d in daily_csv_map else None
        row_dict.update({"合計": tot_q, "売上額": m_info["sales"], "粗利": u_prof, "粗利率": f"{m_rate*100:.1f}%" if m_rate else "-", "粗利合計": round(u_prof * tot_q, 1), "セッション": m_info["sess"], "ページビュー": m_info["pv"], "おすすめ率": f"{avg_bb*100:.1f}%" if avg_bb else "-", "旭電機売上計": round(asahi_p * tot_q, 1)})
        seller_rows.append(row_dict)

    seller_summary_df = pd.DataFrame(seller_rows)
    if not seller_summary_df.empty: seller_summary_df = seller_summary_df.sort_values(by=["合計", "売上額"], ascending=[False, False]).reset_index(drop=True)

    v_df_month = df[(df["年月"] == amz_target_ym) & (df["得意先名"].str.contains("アマゾンベンダー", na=False))].copy()
    v_active_days, v_sku_days = set(), {sk: {d: 0 for d in range(1, 32)} for sk in (list(vendor_prof_map.keys()) if vendor_prof_map else list(sku_name_map.keys()))}
    if not v_df_month.empty:
        for _, r in v_df_month.iterrows():
            sk, dt_val = str(r["商品コード"]).strip(), r["日付_dt"]
            if pd.isna(dt_val) or not sk: continue
            d_num = int(dt_val.day)
            v_active_days.add(d_num)
            if sk not in v_sku_days: v_sku_days[sk] = {d: 0 for d in range(1, 32)}
            v_sku_days[sk][d_num] += int(r["数量"])

    max_v_day = max(v_active_days) if v_active_days else 0
    vendor_rows = []
    for sk, d_map in v_sku_days.items():
        tot_q = sum(d_map.values())
        if tot_q == 0 and sk not in vendor_prof_map and sk not in sku_name_map: continue
        vp = vendor_prof_map.get(sk, {})
        clean_nm = sku_name_map.get(sk, vp.get("name", code_to_name_local.get(sk, sk)))
        sell_p = int(vp.get("sell_price", 0))
        if sell_p == 0 and not v_df_month[v_df_month["商品コード"] == sk].empty: sell_p = int(v_df_month[v_df_month["商品コード"] == sk]["単価"].iloc[-1])
        cost_p = int(vp.get("cost", cost_master_dict.get(sk, 0)))
        u_prof = sell_p - cost_p if sell_p > 0 else 0
        r_dict = {"SKU": sk, "商品名": clean_nm}
        for d in range(1, 32): r_dict[f"{d}日"] = d_map[d] if (d <= max_v_day or d in v_active_days) else None
        r_dict.update({"合計": tot_q, "売値": sell_p, "原価": cost_p, "利益率": f"{(u_prof/sell_p)*100:.1f}%" if sell_p > 0 else "-", "粗利": u_prof, "売上合計": tot_q * sell_p, "粗利合計": tot_q * u_prof})
        vendor_rows.append(r_dict)

    vendor_summary_df = pd.DataFrame(vendor_rows)
    if not vendor_summary_df.empty: vendor_summary_df = vendor_summary_df.sort_values(by=["合計", "売上合計"], ascending=[False, False]).reset_index(drop=True)

    def generate_amazon_profit_sheets_excel(s_df, v_df):
        out = io.BytesIO()
        with pd.ExcelWriter(out, engine="openpyxl") as writer:
            if not s_df.empty: s_df.drop(columns=["（子）ASIN"], errors="ignore").rename(columns={"商品名": ""}).to_excel(writer, index=False, sheet_name="注文数")
            else: pd.DataFrame(columns=["タイトル", ""] + [f"{d}日" for d in range(1, 32)] + ["合計", "売上額", "粗利", "粗利率", "粗利合計", "セッション", "ページビュー", "おすすめ率", "旭電機売上計"]).to_excel(writer, index=False, sheet_name="注文数")
            if not v_df.empty: v_df.to_excel(writer, index=False, sheet_name="ベンダー注文数")
            else: pd.DataFrame(columns=["SKU", "商品名"] + [f"{d}日" for d in range(1, 32)] + ["合計", "売値", "原価", "利益率", "粗利", "売上合計", "粗利合計"]).to_excel(writer, index=False, sheet_name="ベンダー注文数")
        return out.getvalue()

    def generate_monthly_31_sheets_excel(mode_type, s_df, v_df, raw_s_map, v_month_df):
        out = io.BytesIO()
        with pd.ExcelWriter(out, engine="openpyxl") as writer:
            for d in range(1, 32):
                if mode_type == "seller":
                    (raw_s_map[d] if d in raw_s_map else pd.DataFrame()).to_excel(writer, index=False, sheet_name=str(d))
                else:
                    sub_v = v_month_df[v_month_df["日付_dt"].dt.day == d] if not v_month_df.empty else pd.DataFrame()
                    piv_v = sub_v.groupby("商品コード", as_index=False)["数量"].sum().rename(columns={"商品コード": "行ラベル", "数量": "合計 / ご購入明細数量"}) if not sub_v.empty else pd.DataFrame()
                    piv_v.to_excel(writer, index=False, sheet_name=str(d))
            if mode_type == "seller":
                cols_g = ["（子）ASIN", "タイトル", "商品名"] + [f"{d}日" for d in range(1, 32)] + ["合計"]
                (s_df[[c for c in cols_g if c in s_df.columns]].rename(columns={"商品名": ""}) if not s_df.empty else pd.DataFrame(columns=cols_g)).to_excel(writer, index=False, sheet_name="合計")
            else:
                cols_vg = ["SKU", "商品名"] + [f"{d}日" for d in range(1, 32)] + ["合計"]
                (v_df[[c for c in cols_vg if c in v_df.columns]] if not v_df.empty else pd.DataFrame(columns=cols_vg)).to_excel(writer, index=False, sheet_name="合計")
        return out.getvalue()

    dl_c1, dl_c2, dl_c3 = st.columns(3)
    with dl_c1:
        st.download_button("📥 ① アマゾン利益計算用（注文数＋ベンダー注文数）Excel", data=generate_amazon_profit_sheets_excel(seller_summary_df, vendor_summary_df), file_name=f"アマゾン利益計算_注文数シート_{amz_target_ym}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary", use_container_width=True)
    with dl_c2:
        st.download_button(f"📥 ② セラー日別（{tgt_m_int}月.xlsx 形式：1〜31＋合計）", data=generate_monthly_31_sheets_excel("seller", seller_summary_df, vendor_summary_df, raw_daily_df_map, v_df_month), file_name=f"{tgt_m_int}月.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
    with dl_c3:
        st.download_button(f"📥 ③ ベンダー日別（{tgt_m_int}月ベンダー.xlsx 形式）", data=generate_monthly_31_sheets_excel("vendor", seller_summary_df, vendor_summary_df, raw_daily_df_map, v_df_month), file_name=f"{tgt_m_int}月ベンダー.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

    tab_s_view, tab_v_view = st.tabs([
        f"🛒 セラー『注文数』シート（取込済日数: {len(daily_csv_map)}日分 / 合計: {int(seller_summary_df['合計'].sum()) if not seller_summary_df.empty else 0:,}個）",
        f"🏢 ベンダー『ベンダー注文数』シート（弥生連動 / 合計: {int(vendor_summary_df['合計'].sum()) if not vendor_summary_df.empty else 0:,}個）"
    ])
    with tab_s_view:
        if not seller_summary_df.empty: render_universal_table(seller_summary_df, max_h=560)
        else: st.info(f"まだ {amz_target_ym} のセラー日次CSV（例: `{tgt_m_int}-2.csv`）がアップロードされていません。上の枠からCSVを入れてください。")
    with tab_v_view:
        if not vendor_summary_df.empty: render_universal_table(vendor_summary_df[vendor_summary_df["合計"] > 0] if st.checkbox("注文数1個以上の商品のみ画面表示（Excelには全件出力されます）", value=True) else vendor_summary_df, max_h=560)
        else: st.info(f"{amz_target_ym} のアマゾンベンダー売上データがまだありません。")
    st.stop()

# ========================================================
# ★ 🏭 在庫・確保数＆発注検討アラート
# ========================================================
if view_type == '🏭 在庫・確保数＆発注検討アラート':
    code_to_name_map = df.dropna(subset=["商品コード", "商品名"]).drop_duplicates(subset=["商品コード"]).set_index("商品コード")["商品名"].to_dict()
    for k_t, v_t in jan_title_dict.items():
        if v_t and str(v_t).strip() not in ["", "nan"]:
            code_to_name_map[str(k_t).strip()] = str(v_t).strip()

    dt_30_ago = pd.Timestamp(max_date) - pd.Timedelta(days=30)
    dt_90_ago = pd.Timestamp(max_date) - pd.Timedelta(days=90)
    sales_30_map = df[df["日付_dt"] >= dt_30_ago].groupby("商品コード")["数量"].sum().to_dict() if not df.empty else {}
    sales_90_map = df[df["日付_dt"] >= dt_90_ago].groupby("商品コード")["数量"].sum().to_dict() if not df.empty else {}

    all_inv_skus = sorted(list(set(list(logi_inv_map.keys()) + list(yayoi_kakuho_map.keys()) + list(hacchu_master_map.keys()))))
    inv_rows = []
    for sku in all_inv_skus:
        if not sku or sku in ["nan", "None"]: continue
        l_info = logi_inv_map.get(sku, {})
        h_info = hacchu_master_map.get(sku, {})
        y_res = yayoi_kakuho_map.get(sku, h_info.get("sheet_yayoi_res", 0))

        if sku in logi_inv_map:
            good_q = l_info.get("good_qty", 0)
            uninsp_q = l_info.get("uninspected_qty", 0)
            total_logi_q = good_q + uninsp_q
            logi_res_q = l_info.get("reserved_logi", 0)
            inc_q = max(l_info.get("incoming_qty", 0), h_info.get("sheet_incoming", 0))
            wh_str = "/".join(sorted(list(l_info.get("wh_set", set())))) or h_info.get("center", "-")
        else:
            good_q = h_info.get("sheet_logi_stock", 0)
            uninsp_q = 0
            total_logi_q = good_q
            logi_res_q = h_info.get("sheet_logi_res", 0)
            inc_q = h_info.get("sheet_incoming", 0)
            wh_str = h_info.get("center", "-")

        p_name = code_to_name_map.get(sku, h_info.get("category", clean_product_name(l_info.get("name", sku))))
        ord_pt = int(round(h_info.get("order_point", 0)))
        free_after_inc = (total_logi_q + inc_q) - y_res
        immediate_free = good_q - y_res
        order_diff = free_after_inc - ord_pt
        res_gap = logi_res_q - y_res

        q_30 = int(sales_30_map.get(sku, 0))
        q_90 = int(sales_90_map.get(sku, 0))
        daily_avg = q_90 / 90.0 if q_90 > 0 else (h_info.get("term77_qty", 0) / 365.0)
        days_left_str = f"{int(round(max(0, free_after_inc) / daily_avg))}日分" if daily_avg > 0 else "-"

        if h_info.get("discontinued") == "処分品":
            status_lbl = "⏹️ 処分品"
        elif free_after_inc < 0:
            status_lbl = "🚨 確保割れ(不足)"
        elif ord_pt > 0 and order_diff < 0:
            status_lbl = "🚨 要発注検討"
        elif immediate_free < 0 and inc_q > 0:
            status_lbl = "⏳ 入荷・検品待ち"
        else:
            status_lbl = "✅ 適正"

        inv_rows.append({
            "判定": status_lbl,
            "倉庫": wh_str,
            "商品コード": sku,
            "商品名": p_name,
            "良品在庫": good_q,
            "未検品": uninsp_q,
            "入荷予定": inc_q,
            "弥生確保数": y_res,
            "ロジ確保分": logi_res_q,
            "確保ズレ": res_gap,
            "有効フリー在庫": free_after_inc,
            "発注点(5ヶ月)": ord_pt,
            "発注検討差": order_diff,
            "直近30日出荷": q_30,
            "在庫日数目安": days_left_str,
            "入荷予定日": h_info.get("incoming_date", ""),
            "決定事項": h_info.get("decision") or h_info.get("prev_note", "")
        })

    inv_df = pd.DataFrame(inv_rows)

    f_col1, f_col2, f_col3, f_col4 = st.columns([1.6, 1.2, 1.8, 1.0])
    with f_col1:
        inv_filter_mode = st.selectbox("🎯 ステータス絞り込み", ["（すべて表示）", "🚨 要発注・確保割れのみ", "🔒 弥生確保ありのみ", "⚖️ 確保数にズレがある商品", "✅ 適正のみ"])
    with f_col2:
        wh_opts = ["（すべて）"] + sorted([x for x in inv_df["倉庫"].unique() if x and x != "-"])
        sel_wh = st.selectbox("🏭 倉庫・センター", wh_opts)
    with f_col3:
        kw_inv = st.text_input("🔍 品番・商品名で検索", placeholder="例: ABO-08PK, AHL-1400, ABA-105...")
    with f_col4:
        inv_limit = st.selectbox("表示件数", [50, 100, 200, 500, "全件"], index=1, key="inv_limit_sel")

    work_inv = inv_df.copy()
    if inv_filter_mode == "🚨 要発注・確保割れのみ":
        work_inv = work_inv[work_inv["判定"].str.contains("🚨")]
    elif inv_filter_mode == "🔒 弥生確保ありのみ":
        work_inv = work_inv[work_inv["弥生確保数"] > 0]
    elif inv_filter_mode == "⚖️ 確保数にズレがある商品":
        work_inv = work_inv[work_inv["確保ズレ"] != 0]
    elif inv_filter_mode == "✅ 適正のみ":
        work_inv = work_inv[work_inv["判定"] == "✅ 適正"]

    if sel_wh != "（すべて）":
        work_inv = work_inv[work_inv["倉庫"].str.contains(sel_wh, na=False)]
    if kw_inv.strip():
        kw_i = kw_inv.strip()
        work_inv = work_inv[work_inv["商品コード"].str.contains(kw_i, case=False, na=False) | work_inv["商品名"].str.contains(kw_i, case=False, na=False)]

    work_inv = work_inv.sort_values(by=["発注検討差", "有効フリー在庫"], ascending=[True, True]).reset_index(drop=True)
    work_inv.insert(0, "番号", range(1, len(work_inv) + 1))

    alert_cnt = len(inv_df[inv_df["判定"].str.contains("🚨")])
    short_cnt = len(inv_df[inv_df["有効フリー在庫"] < 0])
    sum_c1, sum_c2 = st.columns([4.0, 1.2])
    with sum_c1:
        st.markdown(f"""
        <div class="total-card" style="margin-bottom:0;">
            <b>表示対象:</b> <span style="color:#1a73e8; font-weight:bold;">{len(work_inv):,} 品番</span> &nbsp;|&nbsp;
            <b>🚨 要発注検討(発注点割れ):</b> <span style="color:#d93025; font-size:16px; font-weight:bold;">{alert_cnt:,} 品番</span> &nbsp;|&nbsp;
            <b>⚠️ 確保割れ(フリー在庫マイナス):</b> <span style="color:#d93025; font-weight:bold;">{short_cnt:,} 品番</span>
        </div>
        """, unsafe_allow_html=True)
    with sum_c2:
        st.download_button("📥 在庫・発注検討表をExcel保存", data=generate_excel_bytes(work_inv, "在庫_発注検討一覧"), file_name=f"在庫_発注検討一覧_{datetime.now().strftime('%Y%m%d')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

    st.write("")
    insp_col1, insp_col2 = st.columns([1.5, 2.5])
    sku_choices = work_inv["商品コード"].tolist() if not work_inv.empty else all_inv_skus
    with insp_col1:
        sel_insp_sku = st.selectbox("🔍 商品を選択して「得意先別の確保内訳・倉庫別在庫」を詳しく見る", sku_choices if sku_choices else ["（なし）"])
    if sel_insp_sku and sel_insp_sku != "（なし）":
        row_s = inv_df[inv_df["商品コード"] == sel_insp_sku].iloc[0] if not inv_df[inv_df["商品コード"] == sel_insp_sku].empty else {}
        l_det = logi_inv_map.get(sel_insp_sku, {})
        with insp_col2:
            st.markdown(f"""
            <div class="total-card" style="margin-top:2px; padding:8px 14px; background:#eff6ff; border-left:5px solid #1d4ed8;">
                <b>【{sel_insp_sku}】 {row_s.get('商品名', '')}</b> &nbsp;➔&nbsp;
                良品: <b>{int(row_s.get('良品在庫', 0)):,}</b> (青山:{l_det.get('aoyama_good', 0):,} / 名古屋:{l_det.get('nagoya_good', 0):,}) ｜
                未検品: <b>{int(row_s.get('未検品', 0)):,}</b> ｜
                弥生確保: <b style="color:#d93025;">{int(row_s.get('弥生確保数', 0)):,}</b> ｜
                有効フリー在庫: <b style="color:#16a34a; font-size:15px;">{int(row_s.get('有効フリー在庫', 0)):,} 個</b>
            </div>
            """, unsafe_allow_html=True)

        if not yayoi_kakuho_df.empty:
            sub_k_df = yayoi_kakuho_df[yayoi_kakuho_df["商品コード"] == sel_insp_sku].sort_values(by=["受注日", "受注番号"], ascending=[False, False]).reset_index(drop=True)
            if not sub_k_df.empty:
                with st.expander(f"🔒 【{sel_insp_sku}】弥生・在庫確保の得意先別明細（{len(sub_k_df)}件 / 合計 {int(sub_k_df['確保残数'].sum()):,} 個）", expanded=True):
                    render_universal_table(sub_k_df[["受注日", "受注番号", "得意先名", "担当者", "倉庫", "確保残数", "納入期日", "備考"]], max_h=240)

    disp_inv_df = work_inv.head(int(inv_limit)) if inv_limit != "全件" else work_inv
    if is_mobile_view:
        mob_inv_cols = ["番号", "判定", "商品コード", "商品名", "良品在庫", "弥生確保数", "有効フリー在庫", "発注検討差"]
        disp_inv_df = disp_inv_df[[c for c in mob_inv_cols if c in disp_inv_df.columns]]
    render_universal_table(disp_inv_df, max_h=560)
    st.stop()

# ========================================================
# ★ 📦 商品コードクリック対応テーブル ＆ 在庫・確保詳細ポップアップ（POP）
# ========================================================
_CLICKABLE_PROD_DIR = os.path.join(BASE_DIR, "_prod_table_comp_v1")
os.makedirs(_CLICKABLE_PROD_DIR, exist_ok=True)
_CLICKABLE_PROD_HTML = os.path.join(_CLICKABLE_PROD_DIR, "index.html")
_PROD_COMP_CONTENT = """<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
  body { margin: 0; padding: 0; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; font-size: 13px; background: transparent; }
  .tbl-wrap { overflow-x: auto; max-height: 600px; border: 1px solid #cbd5e1; border-radius: 6px; background-color: #ffffff; }
  table { width: 100%; border-collapse: collapse; }
  th { border: 1px solid #cbd5e1; padding: 8px 10px; background: #f1f5f9; white-space: nowrap; position: sticky; top: 0; z-index: 1; color: #0f172a; }
  td { border: 1px solid #e2e8f0; padding: 6px 10px; font-variant-numeric: tabular-nums; white-space: nowrap; color: #0f172a; }
  tr:nth-child(even) { background: #f8fafc; }
  tr:nth-child(odd) { background: #ffffff; }
  tr:hover { background: #eff6ff !important; }
  .sku-link { color: #1d4ed8; font-weight: bold; text-decoration: underline; cursor: pointer; }
  .sku-link:hover { color: #1e40af; }
</style>
</head>
<body>
<div id="root"></div>
<script>
function sendMessageToStreamlitClient(type, data) {
  var outData = Object.assign({ isStreamlitMessage: true, type: type }, data);
  window.parent.postMessage(outData, "*");
}
function init() { sendMessageToStreamlitClient("streamlit:componentReady", {apiVersion: 1}); }
function setFrameHeight(height) { sendMessageToStreamlitClient("streamlit:setFrameHeight", {height: height}); }
function sendValue(value) { sendMessageToStreamlitClient("streamlit:setComponentValue", {value: value}); }

window.addEventListener("message", function(event) {
  if (event.data.type !== "streamlit:render") return;
  var args = event.data.args;
  var root = document.getElementById("root");
  root.innerHTML = args.table_html || "";
  var links = root.querySelectorAll(".sku-link");
  links.forEach(function(el) {
    el.addEventListener("click", function() {
      var s_code = this.getAttribute("data-sku");
      sendValue({sku: s_code, ts: Date.now()});
    });
  });
  setFrameHeight(args.height || 610);
});
init();
</script>
</body>
</html>"""
try:
    with open(_CLICKABLE_PROD_HTML, "w", encoding="utf-8") as _f:
        _f.write(_PROD_COMP_CONTENT)
except Exception:
    pass
_clickable_prod_table_func = components.declare_component("clickable_prod_table_v1", path=_CLICKABLE_PROD_DIR)

def render_clickable_product_table(df_data, comp_key="prod_rank_clickable_tbl_v1", max_h=620):
    if df_data is None or df_data.empty:
        st.info("該当するデータがありません。")
        return None
    center_cols = {"番号", "順位", "日付", "販売期間", "最終売上日", "初回売上日", "未販売期間", "目標達成率", "通期進捗率", "前年比", "粗利率", "構成比"}
    left_cols = {"得意先名", "商品コード", "商品名", "商品番号", "ランク名", "担当者", "客先名", "販売価格帯"}
    narrow_cols = {"番号", "順位"}

    header_cells = []
    for c in df_data.columns:
        align = 'center' if c in center_cols else ('left' if c in left_cols else 'right')
        w_style = "width: 48px; min-width: 48px; max-width: 52px;" if c in narrow_cols else ""
        header_cells.append(f"<th style='text-align:{align}; {w_style}'>{c}</th>")

    body_rows = []
    for _, row in df_data.iterrows():
        cells = []
        for c in df_data.columns:
            val = row[c]
            val_str = "" if pd.isna(val) else str(val)
            align = "center" if c in center_cols else ("left" if c in left_cols else "right")
            w_style = "width: 48px; min-width: 48px; max-width: 52px;" if c in narrow_cols else ""
            if isinstance(val, (int, float)) and c not in center_cols:
                val_str = f"{int(val):+,}" if c in ["増減", "増減額", "目標差額", "粗利増減"] else f"{int(val):,}"
            is_neg = (isinstance(val, (int, float)) and val < 0) or (isinstance(val_str, str) and (val_str.startswith("-") or val_str.startswith("▲-")))
            style_extra = "color: #dc2626; font-weight: bold;" if is_neg else ""

            if c == "商品コード":
                safe_s = val_str.replace('"', '&quot;')
                cells.append(f"<td style='text-align:left;' title='クリックして在庫・確保情報を表示'><span class='sku-link' data-sku=\"{safe_s}\">🔍 {val_str}</span></td>")
            else:
                cells.append(f"<td style='text-align:{align}; {w_style} {style_extra}'>{val_str}</td>")
        body_rows.append(f"<tr>{''.join(cells)}</tr>")

    full_html = f"<div class='tbl-wrap' style='max-height:{max_h}px;'><table><thead><tr>{''.join(header_cells)}</tr></thead><tbody>{''.join(body_rows)}</tbody></table></div>"
    calc_h = min(max_h + 20, 80 + len(df_data) * 36)
    return _clickable_prod_table_func(table_html=full_html, height=calc_h, key=comp_key)

@st.dialog("📦 商品別 在庫・確保数・発注検討ステータス詳細", width="large")
def show_product_inventory_dialog(sku_code, prod_name):
    l_info = logi_inv_map.get(sku_code, {})
    h_info = hacchu_master_map.get(sku_code, {})
    y_res = yayoi_kakuho_map.get(sku_code, h_info.get("sheet_yayoi_res", 0))

    if sku_code in logi_inv_map:
        good_q = l_info.get("good_qty", 0)
        ao_good = l_info.get("aoyama_good", 0)
        ng_good = l_info.get("nagoya_good", 0)
        uninsp_q = l_info.get("uninspected_qty", 0)
        total_logi_q = good_q + uninsp_q
        logi_res_q = l_info.get("reserved_logi", 0)
        inc_q = max(l_info.get("incoming_qty", 0), h_info.get("sheet_incoming", 0))
        wh_str = "/".join(sorted(list(l_info.get("wh_set", set())))) or h_info.get("center", "-")
    else:
        good_q = h_info.get("sheet_logi_stock", 0)
        ao_good, ng_good, uninsp_q = 0, 0, 0
        total_logi_q = good_q
        logi_res_q = h_info.get("sheet_logi_res", 0)
        inc_q = h_info.get("sheet_incoming", 0)
        wh_str = h_info.get("center", "-")

    ord_pt = int(round(h_info.get("order_point", 0)))
    free_after_inc = (total_logi_q + inc_q) - y_res
    order_diff = free_after_inc - ord_pt
    res_gap = logi_res_q - y_res
    inc_date_str = h_info.get("incoming_date", "-") or "-"
    note_str = h_info.get("decision") or h_info.get("prev_note") or "-"

    dt_30_ago = pd.Timestamp(max_date) - pd.Timedelta(days=30)
    dt_90_ago = pd.Timestamp(max_date) - pd.Timedelta(days=90)
    sub_s = df[df["商品コード"] == sku_code] if not df.empty else pd.DataFrame()
    q_30 = int(sub_s[sub_s["日付_dt"] >= dt_30_ago]["数量"].sum()) if not sub_s.empty else 0
    q_90 = int(sub_s[sub_s["日付_dt"] >= dt_90_ago]["数量"].sum()) if not sub_s.empty else 0
    daily_avg = q_90 / 90.0 if q_90 > 0 else (h_info.get("term77_qty", 0) / 365.0)
    days_left_str = f"約 {int(round(max(0, free_after_inc) / daily_avg))} 日分" if daily_avg > 0 else "-"

    if h_info.get("discontinued") == "処分品":
        badge_html = "<span style='background:#64748b; color:white; padding:4px 10px; border-radius:4px; font-weight:bold;'>⏹️ 処分品</span>"
    elif free_after_inc < 0:
        badge_html = "<span style='background:#dc2626; color:white; padding:4px 10px; border-radius:4px; font-weight:bold;'>🚨 確保割れ（在庫不足）</span>"
    elif ord_pt > 0 and order_diff < 0:
        badge_html = "<span style='background:#ea580c; color:white; padding:4px 10px; border-radius:4px; font-weight:bold;'>🚨 要発注検討（発注点割れ）</span>"
    else:
        badge_html = "<span style='background:#16a34a; color:white; padding:4px 10px; border-radius:4px; font-weight:bold;'>✅ 在庫適正</span>"

    st.markdown(f"#### 🎯 【{sku_code}】 {prod_name} &nbsp; {badge_html}", unsafe_allow_html=True)

    st.markdown(f"""
    <div style="display:grid; grid-template-columns: repeat(4, 1fr); gap:10px; margin-top:10px; margin-bottom:12px;">
        <div style="border:1px solid #cbd5e1; border-top:4px solid #2563eb; border-radius:6px; padding:10px; background:#f8fafc;">
            <div style="font-size:12px; color:#475569;">📦 ロジザード良品在庫</div>
            <div style="font-size:20px; font-weight:bold; color:#0f172a; margin-top:2px;">{good_q:,} 個</div>
            <div style="font-size:11px; color:#64748b; margin-top:2px;">青山: <b>{ao_good:,}</b> / 名古屋: <b>{ng_good:,}</b> (倉庫:{wh_str})</div>
        </div>
        <div style="border:1px solid #cbd5e1; border-top:4px solid #0284c7; border-radius:6px; padding:10px; background:#f8fafc;">
            <div style="font-size:12px; color:#475569;">🚢 未検品 ・ 入荷予定</div>
            <div style="font-size:20px; font-weight:bold; color:#0f172a; margin-top:2px;">未検: {uninsp_q:,} / 入荷: {inc_q:,}</div>
            <div style="font-size:11px; color:#64748b; margin-top:2px;">入荷予定日: <b>{inc_date_str}</b></div>
        </div>
        <div style="border:1px solid #cbd5e1; border-top:4px solid #dc2626; border-radius:6px; padding:10px; background:#fef2f2;">
            <div style="font-size:12px; color:#991b1b;">🔒 弥生 確保数（引当残）</div>
            <div style="font-size:20px; font-weight:bold; color:#dc2626; margin-top:2px;">{y_res:,} 個</div>
            <div style="font-size:11px; color:#7f1d1d; margin-top:2px;">ロジ確保分: <b>{logi_res_q:,}</b> (ズレ: {res_gap:+,})</div>
        </div>
        <div style="border:1px solid #cbd5e1; border-top:4px solid #16a34a; border-radius:6px; padding:10px; background:#f0fdf4;">
            <div style="font-size:12px; color:#166534;">✅ 有効フリー在庫（販売可能）</div>
            <div style="font-size:20px; font-weight:bold; color:{'#dc2626' if free_after_inc < 0 else '#15803d'}; margin-top:2px;">{free_after_inc:,} 個</div>
            <div style="font-size:11px; color:#14532d; margin-top:2px;">直近30日出荷: <b>{q_30:,}個</b> (在庫目安: {days_left_str})</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown(f"""
    <div class="total-card" style="margin-bottom:12px; font-size:13px;">
        <b>📊 発注検討表データ</b> &nbsp;|&nbsp;
        <b>発注点(5ヶ月):</b> {ord_pt:,} 個 &nbsp;|&nbsp;
        <b>発注検討差(フリー在庫 − 発注点):</b> <span style="color:{'#dc2626' if order_diff < 0 else '#15803d'}; font-weight:bold; font-size:15px;">{order_diff:+,} 個</span> &nbsp;|&nbsp;
        <b>丸藤確保:</b> {h_info.get('marufuji_res', 0):,} 個 &nbsp;|&nbsp;
        <b>リンベル確保:</b> {h_info.get('ringbell_res', 0):,} 個 &nbsp;|&nbsp;
        <b>決定事項・備考:</b> <span style="color:#1e40af; font-weight:bold;">{note_str}</span>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("##### 🔒 弥生・在庫確保の得意先別内訳（受注日が新しい順）")
    if not yayoi_kakuho_df.empty:
        sub_k = yayoi_kakuho_df[yayoi_kakuho_df["商品コード"] == sku_code].sort_values(by=["受注日", "受注番号"], ascending=[False, False]).reset_index(drop=True)
        if not sub_k.empty:
            render_universal_table(sub_k[["受注日", "受注番号", "得意先名", "担当者", "倉庫", "確保残数", "納入期日", "備考"]], max_h=260)
        else:
            st.info("現在、弥生で確保（受注残）されている明細はありません。")
    else:
        st.info("現在、弥生で確保（受注残）されている明細はありません。")

# ========================================================
# ★ 🧮 簡易見積＆利益シミュレーター（＋ 🏷️ 加盟店価格登録CSV作成）
# ========================================================
if view_type == '🧮 簡易見積＆利益シミュレーター':
    code_to_name_map = df.dropna(subset=["商品コード", "商品名"]).drop_duplicates(subset=["商品コード"]).set_index("商品コード")["商品名"].to_dict()
    for k_t, v_t in jan_title_dict.items():
        sk_trim = str(k_t).strip()
        if v_t and str(v_t).strip() not in ["", "nan"]:
            code_to_name_map[sk_trim] = str(v_t).strip()

    raw_calc_codes = sorted(list(set(
        [str(c).strip() for c in code_to_name_map.keys()] +
        [str(c).strip() for c in cost_master_dict.keys()] +
        [str(c).strip() for c in jan_master_dict.keys()] +
        [str(c).strip() for c in jan_title_dict.keys()] +
        [str(c).strip() for c in member_default_min_map.keys()] +
        [str(c).strip() for c in lotmember_sku_meta_dict.keys()]
    )))
    if "ABA-40" in raw_calc_codes: raw_calc_codes.remove("ABA-40")
    if "" in raw_calc_codes: raw_calc_codes.remove("")

    parent_sku_candidates = set()
    for c in raw_calc_codes:
        for other in raw_calc_codes:
            if other != c and other.startswith(c) and len(other) > len(c):
                parent_sku_candidates.add(c)
                break
    all_calc_codes = [c for c in raw_calc_codes if c not in parent_sku_candidates]

    STANDARD_RANKS = ["ロット単価", "1.05倍", "1.07倍", "1.1倍", "1.15倍", "1.2倍", "1.25倍", "1.3倍", "1.35倍", "1.4倍", "1.45倍", "1.5倍", "0.9倍"]
    RANK_MULTIPLIERS = {"1.05倍": 1.05, "1.07倍": 1.07, "1.1倍": 1.10, "1.15倍": 1.15, "1.2倍": 1.20, "1.25倍": 1.25, "1.3倍": 1.30, "1.35倍": 1.35, "1.4倍": 1.40, "1.45倍": 1.45, "1.5倍": 1.50, "0.9倍": 0.90}

    def get_sku_details(sku_code, rank_selection="ロット単価", cust_comp_name=""):
        s_trim = str(sku_code).strip()
        matched_key = s_trim
        if s_trim not in jan_title_dict and s_trim not in cost_master_dict and s_trim not in member_default_min_map:
            candidates = [k for k in jan_title_dict.keys() if str(k).strip().startswith(s_trim) or s_trim in str(k)]
            if candidates: matched_key = str(candidates[0]).strip()
        p_name = jan_title_dict.get(matched_key, "")
        if not p_name or p_name in ["nan", matched_key]:
            p_name = code_to_name_map.get(matched_key, s_trim)
        p_jan = jan_master_dict.get(matched_key, "")
        p_cost = int(cost_master_dict.get(matched_key, 0.0))
        norm_cust = re.sub(r'[\s\u3000\xa0]+', '', str(cust_comp_name)) if cust_comp_name else ""
        min_p = member_default_min_map.get(matched_key, 0)
        comp_lot_p = 0
        if cust_comp_name and (cust_comp_name, matched_key) in member_cust_sku_map:
            comp_lot_p = member_cust_sku_map[(cust_comp_name, matched_key)]
        elif norm_cust and (norm_cust, matched_key) in member_cust_sku_map:
            comp_lot_p = member_cust_sku_map[(norm_cust, matched_key)]
        clean_r = str(rank_selection).strip() if rank_selection else "ロット単価"
        if clean_r in RANK_MULTIPLIERS:
            p_price = int(round(min_p * RANK_MULTIPLIERS[clean_r])) if min_p > 0 else 99999
        else:
            p_price = comp_lot_p if comp_lot_p > 0 else (min_p if min_p > 0 else 99999)
        return p_name, p_jan, (p_price if p_price > 0 else 99999), p_cost, min_p, clean_r

    if "saved_estimates" not in st.session_state:
        st.session_state.saved_estimates = load_saved_quotes()

    est_staff_list = list(STAFF_MAP.values()) + ["（すべて表示）"]
    col_menu_left, col_menu_right = st.columns([2.3, 1.3])
    with col_menu_right:
        est_staff = st.selectbox("👤 担当者（バッティング防止）", est_staff_list, index=7 if "水原" in est_staff_list else 0, key="est_staff_selector_box_v3")

    s_key = est_staff
    k_items, k_ver = f"quotation_items_{s_key}", f"editor_version_{s_key}"
    k_comp, k_person, k_post, k_addr = f"in_cust_comp_{s_key}", f"in_cust_person_{s_key}", f"in_cust_post_{s_key}", f"in_cust_addr_{s_key}"
    k_s_comp, k_s_person, k_s_post, k_s_addr, k_s_tel = f"in_ship_comp_{s_key}", f"in_ship_person_{s_key}", f"in_ship_post_{s_key}", f"in_ship_addr_{s_key}", f"in_ship_tel_{s_key}"
    k_sel_c = f"sel_cust_choice_{s_key}"

    if k_items not in st.session_state: st.session_state[k_items] = []
    if k_ver not in st.session_state: st.session_state[k_ver] = 0
    for k_f in [k_comp, k_person, k_post, k_addr, k_s_comp, k_s_person, k_s_post, k_s_addr, k_s_tel]:
        if k_f not in st.session_state: st.session_state[k_f] = ""

    if "load_target_estimate" in st.session_state and st.session_state["load_target_estimate"]:
        loaded_est = st.session_state.pop("load_target_estimate")
        st.session_state[k_comp] = loaded_est.get("cust_comp", "")
        st.session_state[k_person] = loaded_est.get("cust_person", "")
        st.session_state[k_post] = loaded_est.get("cust_post", "")
        st.session_state[k_addr] = loaded_est.get("cust_addr", "")
        st.session_state[k_s_comp] = loaded_est.get("ship_comp", "")
        st.session_state[k_s_person] = loaded_est.get("ship_person", "")
        st.session_state[k_s_post] = loaded_est.get("ship_post", "")
        st.session_state[k_s_addr] = loaded_est.get("ship_addr", "")
        st.session_state[k_s_tel] = loaded_est.get("ship_tel", "")
        st.session_state[k_sel_c] = loaded_est.get("cust_comp", "（手動入力）") if loaded_est.get("cust_comp") else "（手動入力）"
        st.session_state[f"q_no_{s_key}"] = loaded_est.get("q_no", f"{datetime.now().strftime('%Y%m%d')}-01")
        st.session_state[f"q_memo_{s_key}"] = loaded_est.get("memo", "")
        st.session_state[f"q_smemo_{s_key}"] = loaded_est.get("hitokoto", f"営業担当：{est_staff}")
        st.session_state[f"q_ship_{s_key}"] = int(clean_num(loaded_est.get("shipping_fee", 0)))
        st.session_state[k_items] = loaded_est.get("items", [])
        st.session_state[k_ver] += 1
        st.session_state["est_active_tab_v3"] = "📝 新規見積作成・シミュレーター"

    staff_Filter_list = [e for e in st.session_state.saved_estimates if est_staff == "（すべて表示）" or e.get("staff", "") == est_staff or (est_staff in str(e.get("hitokoto", "")))]
    tab_opt_new = "📝 新規見積作成・シミュレーター"
    tab_opt_lot = "🏷️ 加盟店価格登録CSV作成（item_class_lotmember）"
    tab_opt_list = f"📋 保存済み見積一覧・スマレジ注文CSV出力（{len(staff_Filter_list)}件）"
    if "est_active_tab_v3" in st.session_state and st.session_state["est_active_tab_v3"] not in [tab_opt_new, tab_opt_lot, tab_opt_list]:
        if str(st.session_state["est_active_tab_v3"]).startswith("📋"):
            st.session_state["est_active_tab_v3"] = tab_opt_list
        elif str(st.session_state["est_active_tab_v3"]).startswith("🏷️"):
            st.session_state["est_active_tab_v3"] = tab_opt_lot
        else:
            st.session_state["est_active_tab_v3"] = tab_opt_new

    with col_menu_left:
        tab_choice = st.radio("機能メニュー", [tab_opt_new, tab_opt_lot, tab_opt_list], horizontal=True, key="est_active_tab_v3")

    if tab_choice == tab_opt_lot:
        st.markdown(f"##### 🏷️ 加盟店価格登録用CSV作成（`item_class_lotmember.csv`） &nbsp;<span style='font-size:13px; color:#2563eb;'>［作業担当: {est_staff}］</span>", unsafe_allow_html=True)
        lot_items_key = f"lotmember_items_{s_key}"
        lot_ver_key = f"lotmember_ver_{s_key}"
        if lot_items_key not in st.session_state: st.session_state[lot_items_key] = []
        if lot_ver_key not in st.session_state: st.session_state[lot_ver_key] = 0

        grp_comp_list = [k for k in member_group_master_dict.keys() if not re.search(r'^\s*$', str(k))]
        default_nascom_idx = grp_comp_list.index("株式会社ナスコム") if "株式会社ナスコム" in grp_comp_list else 0
        c_lot1, c_lot2, c_lot3 = st.columns([2.2, 1.0, 1.0])
        with c_lot1:
            sel_lot_comp = st.selectbox("🏢 登録対象の会社名（加盟店グループ管理マスタ）", grp_comp_list if grp_comp_list else ["株式会社ナスコム"], index=default_nascom_idx, key=f"lot_comp_sel_{s_key}")
        grp_info = lookup_group_info_by_comp(sel_lot_comp, member_group_master_dict)
        with c_lot2:
            in_grp_no = st.text_input("F列：グループ連番（自動反映）", value=grp_info.get("group_no", ""), key=f"lot_grp_no_{sel_lot_comp}_{s_key}")
        with c_lot3:
            in_mem_cd = st.text_input("G列：加盟店コード（自動反映）", value=grp_info.get("member_code", ""), key=f"lot_mem_cd_{sel_lot_comp}_{s_key}")

        REG_TYPE_OPTIONS = ["新規 (n)", "更新 (u)"]

        def apply_rank_to_lot_items(target_rank, current_cust=""):
            if st.session_state[lot_items_key]:
                for item in st.session_state[lot_items_key]:
                    _, _, new_p, _, _, applied_lbl = get_sku_details(item["商品コード"], target_rank, current_cust)
                    item["適用ランク"] = applied_lbl
                    item["登録単価"] = int(new_p)
                    item["粗利額"] = int(new_p) - int(item["単価原価"])
                    item["粗利率"] = f"{(item['粗利額'] / int(new_p) * 100):.1f}%" if int(new_p) > 0 else "-"
                st.session_state[lot_ver_key] += 1

        def apply_regtype_to_lot_items(target_reg):
            if st.session_state[lot_items_key]:
                for item in st.session_state[lot_items_key]:
                    item["登録区分"] = target_reg
                st.session_state[lot_ver_key] += 1

        lc_del, lc_clr, lc_reg_all, lc_rk_all = st.columns([1.4, 1.0, 1.6, 1.8])
        trigger_lot_del = False
        with lc_del:
            if st.button("🗑️ 選択した行を削除", key=f"lot_btn_del_{s_key}"): trigger_lot_del = True
        with lc_clr:
            if st.button("全クリア", key=f"lot_btn_clr_{s_key}"):
                st.session_state[lot_items_key] = []
                st.session_state[lot_ver_key] += 1
                st.rerun()
        with lc_reg_all:
            reg_all_k = f"lot_reg_all_{s_key}"
            st.selectbox("🔄 登録区分一括変更", REG_TYPE_OPTIONS, index=1, key=reg_all_k, on_change=lambda: apply_regtype_to_lot_items(st.session_state[reg_all_k]))
        with lc_rk_all:
            rk_all_k = f"lot_rk_all_{s_key}"
            st.selectbox("🏷️ 価格ランク一括指定", STANDARD_RANKS, key=rk_all_k, on_change=lambda: apply_rank_to_lot_items(st.session_state[rk_all_k], sel_lot_comp))

        lot_cols_order = ["選択", "登録区分", "商品コード", "商品名", "最低下代", "適用ランク", "登録単価", "単価原価", "粗利額", "粗利率"]
        if st.session_state[lot_items_key]:
            raw_lot_df = pd.DataFrame(st.session_state[lot_items_key])
            for c in lot_cols_order:
                if c not in raw_lot_df.columns:
                    raw_lot_df[c] = False if c == "選択" else ("更新 (u)" if c == "登録区分" else ("ロット単価" if c == "適用ランク" else ("-" if c == "粗利率" else 0)))
            raw_lot_df["最低下代"] = raw_lot_df["最低下代"].apply(lambda x: int(clean_num(x)))
            raw_lot_df["登録単価"] = raw_lot_df["登録単価"].apply(lambda x: int(clean_num(x)))
            raw_lot_df["単価原価"] = raw_lot_df["単価原価"].apply(lambda x: int(clean_num(x)))
            raw_lot_df["粗利額"] = raw_lot_df["登録単価"] - raw_lot_df["単価原価"]
            raw_lot_df["粗利率"] = raw_lot_df.apply(lambda r: f"{(r['粗利額'] / r['登録単価'] * 100):.1f}%" if r['登録単価'] > 0 else "-", axis=1)
            raw_lot_df = raw_lot_df[lot_cols_order]
        else:
            raw_lot_df = pd.DataFrame([{"選択": False, "登録区分": "更新 (u)", "商品コード": "", "商品名": "", "最低下代": 0, "適用ランク": "ロット単価", "登録単価": 0, "単価原価": 0, "粗利額": 0, "粗利率": "-"}], columns=lot_cols_order)

        edited_lot_df = st.data_editor(
            raw_lot_df,
            column_config={
                "選択": st.column_config.CheckboxColumn("選択", default=False, width="small"),
                "登録区分": st.column_config.SelectboxColumn("A列：登録区分 (▼)", options=REG_TYPE_OPTIONS, required=True, width="small"),
                "商品コード": st.column_config.SelectboxColumn("商品コード（検索・選択）", options=all_calc_codes, required=True, width="medium"),
                "商品名": st.column_config.TextColumn("商品名", width="large"),
                "最低下代": st.column_config.NumberColumn("最低下代", disabled=True, format="¥%d"),
                "適用ランク": st.column_config.SelectboxColumn("適用ランク (▼)", options=STANDARD_RANKS, width="medium"),
                "登録単価": st.column_config.NumberColumn("J列：登録単価（円）", min_value=0, step=1, format="¥%d"),
                "単価原価": st.column_config.NumberColumn("単価原価", format="¥%d"),
                "粗利額": st.column_config.NumberColumn("1個あたり粗利額", disabled=True, format="¥%d"),
                "粗利率": st.column_config.TextColumn("粗利率", disabled=True)
            },
            num_rows="dynamic", use_container_width=True, hide_index=True, key=f"lot_editor_{s_key}_{st.session_state[lot_ver_key]}"
        )

        updated_lot_records, lot_needs_rerun = [], False
        for idx_l, row_l in edited_lot_df.iterrows():
            is_sel = bool(row_l.get("選択", False))
            if trigger_lot_del and is_sel:
                lot_needs_rerun = True
                continue
            c_code = str(row_l.get("商品コード", "")).strip() if pd.notna(row_l.get("商品コード")) else ""
            if not c_code or c_code in ["None", "ABA-40"]: continue

            prev_rec = st.session_state[lot_items_key][idx_l] if idx_l < len(st.session_state[lot_items_key]) else {}
            prev_code = str(prev_rec.get("商品コード", "")).strip()
            prev_rank = str(prev_rec.get("適用ランク", "ロット単価")).strip()
            prev_price = int(clean_num(prev_rec.get("登録単価", 0)))
            prev_reg = str(prev_rec.get("登録区分", "更新 (u)")).strip()

            cur_reg = str(row_l.get("登録区分", "更新 (u)")).strip() if pd.notna(row_l.get("登録区分")) else "更新 (u)"
            if cur_reg not in REG_TYPE_OPTIONS: cur_reg = "更新 (u)"
            cur_rank = str(row_l.get("適用ランク", "ロット単価")).strip() if pd.notna(row_l.get("適用ランク")) else "ロット単価"
            cur_price = int(clean_num(row_l.get("登録単価", 0)))
            cur_cost = int(clean_num(row_l.get("単価原価", 0)))
            cur_name = str(row_l.get("商品名", "")).strip() if pd.notna(row_l.get("商品名")) else ""
            cur_min = int(clean_num(row_l.get("最低下代", 0)))

            auto_name_chk, _, auto_price_chk, auto_cost_chk, auto_min_chk, applied_lbl_chk = get_sku_details(c_code, cur_rank if cur_rank != "None" else "ロット単価", sel_lot_comp)

            if c_code != prev_code or not cur_name or cur_name in ["None", c_code] or cur_price == 0:
                if cur_name in ["", "None", c_code] and auto_name_chk and auto_name_chk != cur_name:
                    cur_name = auto_name_chk
                    lot_needs_rerun = True
                if c_code != prev_code or cur_price == 0:
                    cur_name, cur_price, cur_cost, cur_min, cur_rank = auto_name_chk, int(auto_price_chk), int(auto_cost_chk), int(auto_min_chk), applied_lbl_chk
                    lot_needs_rerun = True
            elif cur_rank != prev_rank:
                cur_price, cur_rank = int(auto_price_chk), applied_lbl_chk
                lot_needs_rerun = True
            elif cur_price != prev_price or cur_reg != prev_reg:
                lot_needs_rerun = True

            unit_profit = cur_price - cur_cost
            updated_lot_records.append({
                "選択": False if trigger_lot_del else is_sel,
                "登録区分": cur_reg,
                "商品コード": c_code,
                "商品名": cur_name,
                "最低下代": cur_min,
                "適用ランク": cur_rank,
                "登録単価": cur_price,
                "単価原価": cur_cost,
                "粗利額": unit_profit,
                "粗利率": f"{(unit_profit / cur_price * 100):.1f}%" if cur_price > 0 else "-"
            })

        if trigger_lot_del or lot_needs_rerun or len(updated_lot_records) != len(st.session_state[lot_items_key]):
            st.session_state[lot_items_key] = updated_lot_records
            st.session_state[lot_ver_key] += 1
            st.rerun()
        else:
            st.session_state[lot_items_key] = updated_lot_records

        if updated_lot_records:
            l_df = pd.DataFrame(updated_lot_records)
            avg_margin = (l_df["粗利額"].sum() / l_df["登録単価"].sum() * 100) if l_df["登録単価"].sum() > 0 else 0.0
            st.markdown(f"""
            <div class="total-card" style="margin-top:10px;">
                <b>対象会社:</b> <span style="color:#1a73e8; font-weight:bold;">{sel_lot_comp}</span> (F列: <b>{in_grp_no}</b> / G列: <b>{in_mem_cd}</b>) &nbsp;|&nbsp;
                <b>登録商品数:</b> <span style="font-size:16px; font-weight:bold; color:#d93025;">{len(l_df):,} 品目</span> (CSV出力: {len(l_df)*2:,} 行) &nbsp;|&nbsp;
                <b>平均粗利率:</b> <span style="color:#188038; font-size:16px; font-weight:bold;">{avg_margin:.1f}%</span>
            </div>
            """, unsafe_allow_html=True)

            csv_bytes_lot = create_lotmember_csv_bytes(updated_lot_records, in_grp_no, in_mem_cd, lotmember_sku_meta_dict)
            st.download_button(
                label="📥 item_class_lotmember.csv をダウンロード",
                data=csv_bytes_lot,
                file_name="item_class_lotmember.csv",
                mime="text/csv",
                type="primary",
                use_container_width=True,
                key=f"dl_lotmember_csv_{s_key}"
            )

    elif tab_choice.startswith("📝"):
        st.markdown(f"##### 🏢 提出先（得意先）の選択 &nbsp;<span style='font-size:13px; color:#2563eb;'>［作業担当: {est_staff}］</span>", unsafe_allow_html=True)
        cust_options, cust_lookup = ["（手動入力）"], {}
        if not customer_master_df.empty:
            c_name_col = "加盟店名" if "加盟店名" in customer_master_df.columns else customer_master_df.columns[0]
            for _, c_row in customer_master_df.iterrows():
                c_name = str(c_row[c_name_col]).strip()
                if c_name and c_name != "nan":
                    cust_options.append(c_name)
                    cust_lookup[c_name] = c_row

        def on_customer_change():
            selected = st.session_state.get(k_sel_c, "（手動入力）")
            if selected == "（手動入力）" or selected not in cust_lookup:
                for k_f in [k_comp, k_person, k_post, k_addr, k_s_comp, k_s_person, k_s_post, k_s_addr, k_s_tel]: st.session_state[k_f] = ""
            else:
                row_c = cust_lookup[selected]
                sei = str(row_c.get("責任者姓", "")).strip() if pd.notna(row_c.get("責任者姓", "")) else ""
                mei = str(row_c.get("責任者名", "")).strip() if pd.notna(row_c.get("責任者名", "")) else ""
                p_name = f"{sei} {mei}".strip()
                if p_name and "様" not in p_name: p_name += " 様"
                p_code = str(row_c.get("郵便番号", "")).replace(".0", "").strip() if pd.notna(row_c.get("郵便番号", "")) else ""
                pref = str(row_c.get("都道府県", "")).strip() if pd.notna(row_c.get("都道府県", "")) else ""
                a1 = str(row_c.get("住所１", "")).strip() if pd.notna(row_c.get("住所１", "")) else ""
                a2 = str(row_c.get("住所２", "")).strip() if pd.notna(row_c.get("住所２", "")) else ""
                full_addr = f"{pref}{a1} {a2}".strip()
                tel = str(row_c.get("電話番号", "")).strip() if pd.notna(row_c.get("電話番号", "")) else ""
                st.session_state[k_comp] = st.session_state[k_s_comp] = selected
                st.session_state[k_person] = st.session_state[k_s_person] = p_name
                st.session_state[k_post] = st.session_state[k_s_post] = p_code
                st.session_state[k_addr] = st.session_state[k_s_addr] = full_addr
                st.session_state[k_s_tel] = tel
                if st.session_state[k_items]:
                    for itm in st.session_state[k_items]:
                        if itm.get("適用ランク") == "ロット単価":
                            _, _, new_p, _, new_min, _ = get_sku_details(itm["商品コード"], "ロット単価", selected)
                            itm["見積単価"], itm["最低下代"] = new_p, new_min
                    st.session_state[k_ver] += 1

        if st.session_state.get(k_sel_c) not in cust_options:
            st.session_state[k_sel_c] = "（手動入力）"
        st.selectbox("🔍 加盟店・取引先を選択（検索可能）", cust_options, key=k_sel_c, on_change=on_customer_change)

        col_addr1, col_addr2 = st.columns(2)
        with col_addr1:
            st.markdown("###### 📄 宛先情報（見積書 表面）")
            q_cust_comp = st.text_input("宛先 会社名", key=k_comp)
            q_cust_person = st.text_input("宛先 担当者名", key=k_person)
            q_cust_post = st.text_input("宛先 郵便番号", key=k_post)
            q_cust_addr = st.text_input("宛先 住所", key=k_addr)
        with col_addr2:
            st.markdown("###### 🚚 お届け先情報（送り先）")
            ship_cust_comp = st.text_input("お届け先 会社名", key=k_s_comp)
            ship_cust_person = st.text_input("お届け先 担当者名", key=k_s_person)
            ship_cust_post = st.text_input("お届け先 郵便番号", key=k_s_post)
            ship_cust_addr = st.text_input("お届け先 住所", key=k_s_addr)
            ship_cust_tel = st.text_input("お届け先 TEL", key=k_s_tel)

        hq1, hq2, hq3, hq4 = st.columns(4)
        with hq1: q_no = st.text_input("見積番号", value=f"{datetime.now().strftime('%Y%m%d')}-01", key=f"q_no_{s_key}")
        with hq2: q_date = st.date_input("お見積日", value=datetime.now().date(), key=f"q_date_{s_key}")
        with hq3: q_exp_date = st.date_input("見積有効期限", value=(datetime.now() + timedelta(days=30)).date(), key=f"q_exp_{s_key}")
        with hq4: q_payment = st.selectbox("お支払い方法", ["銀行振込", "代金引換", "現行通り", "掛け払い（月末締め翌月末払い）"], index=0, key=f"q_pay_{s_key}")

        hq5, hq6, hq7 = st.columns([2, 2, 1])
        with hq5: q_memo = st.text_input("伝達事項 / 発注NO等（CSV要望欄に連動）", placeholder="例: 発注NO.2352640", key=f"q_memo_{s_key}")
        with hq6: q_staff_memo = st.text_input("営業担当・社内備考", value=f"営業担当：{est_staff}", key=f"q_smemo_{s_key}")
        with hq7: q_shipping_fee = st.number_input("送料(税抜)", min_value=0, value=0, step=100, key=f"q_ship_{s_key}")

        st.markdown("---")
        st.markdown("##### 📑 複数品目のまとめ見積リスト（見積書作成）")

        def apply_rank_to_all_items(target_rank, current_cust=""):
            if st.session_state[k_items]:
                for item in st.session_state[k_items]:
                    _, _, new_p, _, _, applied_lbl = get_sku_details(item["商品コード"], target_rank, current_cust)
                    item["適用ランク"], item["見積単価"] = applied_lbl, int(new_p)
                    item["合計金額"] = int(item["数量"]) * int(new_p)
                    item["粗利額"] = item["合計金額"] - (int(item["数量"]) * int(item["単価原価"]))
                    item["粗利率"] = f"{(item['粗利額'] / item['合計金額'] * 100):.1f}%" if item["合計金額"] > 0 else "-"
                st.session_state[k_ver] += 1

        c_btn_del, c_btn_clear, c_rank_box, _ = st.columns([1.5, 1.2, 2.0, 1.3])
        trigger_delete_selected = False
        with c_btn_del:
            if st.button("🗑️ 選択した行を削除", key=f"btn_del_{s_key}"): trigger_delete_selected = True
        with c_btn_clear:
            if st.button("全クリア", key=f"btn_clr_{s_key}"):
                st.session_state[k_items] = []
                st.session_state[k_ver] += 1
                st.rerun()
        with c_rank_box:
            rk_key = f"in_cust_rank_{s_key}"
            st.selectbox("🏷️ 価格ランク一括指定", STANDARD_RANKS, key=rk_key, on_change=lambda: apply_rank_to_all_items(st.session_state[rk_key], q_cust_comp), label_visibility="collapsed")

        cols_order = ["選択", "商品コード", "商品名", "JANコード", "最低下代", "適用ランク", "数量", "見積単価", "合計金額", "単価原価", "粗利額", "粗利率"]
        if st.session_state[k_items]:
            raw_items_df = pd.DataFrame(st.session_state[k_items])
            for c in cols_order:
                if c not in raw_items_df.columns:
                    raw_items_df[c] = False if c == "選択" else ("ロット単価" if c == "適用ランク" else ("-" if c == "粗利率" else 0))
            raw_items_df["数量"] = raw_items_df["数量"].apply(lambda x: max(1, int(clean_num(x))))
            raw_items_df["最低下代"] = raw_items_df["最低下代"].apply(lambda x: int(clean_num(x)))
            raw_items_df["見積単価"] = raw_items_df["見積単価"].apply(lambda x: int(clean_num(x)))
            raw_items_df["単価原価"] = raw_items_df["単価原価"].apply(lambda x: int(clean_num(x)))
            raw_items_df["合計金額"] = raw_items_df["数量"] * raw_items_df["見積単価"]
            raw_items_df["粗利額"] = raw_items_df["合計金額"] - (raw_items_df["数量"] * raw_items_df["単価原価"])
            raw_items_df["粗利率"] = raw_items_df.apply(lambda r: f"{(r['粗利額'] / r['合計金額'] * 100):.1f}%" if r['合計金額'] > 0 else "-", axis=1)
            raw_items_df = raw_items_df[cols_order]
        else:
            raw_items_df = pd.DataFrame([{"選択": False, "商品コード": "", "商品名": "", "JANコード": "", "最低下代": 0, "適用ランク": "ロット単価", "数量": 1, "見積単価": 0, "合計金額": 0, "単価原価": 0, "粗利額": 0, "粗利率": "-"}], columns=cols_order)

        edited_df = st.data_editor(
            raw_items_df,
            column_config={
                "選択": st.column_config.CheckboxColumn("選択", default=False, width="small"),
                "商品コード": st.column_config.SelectboxColumn("商品コード（検索・選択）", options=all_calc_codes, required=True, width="medium"),
                "商品名": st.column_config.TextColumn("商品名", width="large"),
                "JANコード": st.column_config.TextColumn("JANコード"),
                "最低下代": st.column_config.NumberColumn("最低下代", disabled=True, format="¥%d"),
                "適用ランク": st.column_config.SelectboxColumn("適用ランク (▼)", options=STANDARD_RANKS, width="medium"),
                "数量": st.column_config.NumberColumn("数量", min_value=1, step=1),
                "見積単価": st.column_config.NumberColumn("見積単価（円）", min_value=0, step=10, format="¥%d"),
                "合計金額": st.column_config.NumberColumn("合計金額", disabled=True, format="¥%d"),
                "単価原価": st.column_config.NumberColumn("単価原価", format="¥%d"),
                "粗利額": st.column_config.NumberColumn("粗利額", disabled=True, format="¥%d"),
                "粗利率": st.column_config.TextColumn("粗利率", disabled=True)
            },
            num_rows="dynamic", use_container_width=True, hide_index=True, key=f"quote_editor_{s_key}_{st.session_state[k_ver]}"
        )

        updated_records, needs_rerun = [], False
        for idx_e, row_e in edited_df.iterrows():
            is_selected = bool(row_e.get("選択", False))
            if trigger_delete_selected and is_selected:
                needs_rerun = True
                continue
            c_code = str(row_e["商品コード"]).strip() if pd.notna(row_e["商品コード"]) else ""
            if not c_code or c_code in ["None", "ABA-40"]: continue

            prev_rec = st.session_state[k_items][idx_e] if idx_e < len(st.session_state[k_items]) else {}
            prev_code = str(prev_rec.get("商品コード", "")).strip()
            prev_rank = str(prev_rec.get("適用ランク", "ロット単価")).strip()
            prev_qty = int(clean_num(prev_rec.get("数量", 1)))
            prev_price = int(clean_num(prev_rec.get("見積単価", 0)))

            cur_qty = max(1, int(clean_num(row_e.get("数量", 1))))
            cur_rank = str(row_e.get("適用ランク", "ロット単価")).strip() if pd.notna(row_e.get("適用ランク")) else "ロット単価"
            cur_price = int(clean_num(row_e.get("見積単価", 0)))
            cur_cost = int(clean_num(row_e.get("単価原価", 0)))
            cur_name = str(row_e.get("商品名", "")).strip() if pd.notna(row_e.get("商品名")) else ""
            cur_jan = str(row_e.get("JANコード", "")).strip() if pd.notna(row_e.get("JANコード")) else ""
            cur_min = int(clean_num(row_e.get("最低下代", 0)))

            auto_name, auto_jan, auto_price, auto_cost, auto_min, applied_lbl = get_sku_details(c_code, cur_rank if cur_rank != "None" else "ロット単価", q_cust_comp)

            if c_code != prev_code or cur_price == 0:
                cur_name, cur_jan, cur_price, cur_cost, cur_min, cur_rank = auto_name, str(auto_jan), int(auto_price), int(auto_cost), int(auto_min), applied_lbl
                needs_rerun = True
            else:
                if (not cur_name or cur_name in ["None", "nan", c_code]) and auto_name and auto_name != cur_name:
                    cur_name = auto_name
                    needs_rerun = True
                if (not cur_jan or cur_jan in ["None", "nan"]) and auto_jan and str(auto_jan) != cur_jan:
                    cur_jan = str(auto_jan)
                    needs_rerun = True
                if int(auto_cost) > 0 and cur_cost != int(auto_cost):
                 cur_cost = int(auto_cost)
                 needs_rerun = True
                if cur_min == 0 and int(auto_min) > 0:
                    cur_min = int(auto_min)
                    needs_rerun = True
                if cur_rank != prev_rank:
                    cur_price, cur_rank = int(auto_price), applied_lbl
                    needs_rerun = True
                elif cur_qty != prev_qty or cur_price != prev_price:
                    needs_rerun = True

            new_total = cur_qty * cur_price
            new_profit = new_total - (cur_qty * cur_cost)
            updated_records.append({
                "選択": False if trigger_delete_selected else is_selected, "商品コード": c_code, "商品名": cur_name,
                "JANコード": cur_jan, "最低下代": cur_min, "適用ランク": cur_rank, "数量": cur_qty,
                "見積単価": cur_price, "合計金額": new_total, "単価原価": cur_cost, "粗利額": new_profit,
                "粗利率": f"{(new_profit / new_total * 100):.1f}%" if new_total > 0 else "-"
            })

        if trigger_delete_selected or needs_rerun or len(updated_records) != len(st.session_state[k_items]):
            st.session_state[k_items] = updated_records
            st.session_state[k_ver] += 1
            st.rerun()
        else:
            st.session_state[k_items] = updated_records

        if updated_records:
            calc_df = pd.DataFrame(updated_records)
            q_tot_qty, q_tot_unit_price, q_tot_sales, q_tot_profit = int(calc_df["数量"].sum()), int(calc_df["見積単価"].sum()), int(calc_df["合計金額"].sum()), int(calc_df["粗利額"].sum())
            q_tot_margin = (q_tot_profit / q_tot_sales * 100) if q_tot_sales > 0 else 0.0
            q_tax = int(round((q_tot_sales + q_shipping_fee) * 0.10))
            q_grand_total = q_tot_sales + q_shipping_fee + q_tax

            st.markdown(f"""
            <div class="total-card" style="margin-top:10px;">
                <b>数量合計:</b> <span style="font-size:16px; font-weight:bold; color:#1a73e8;">{q_tot_qty:,} 個</span> &nbsp;|&nbsp;
                <b>税抜商品計:</b> <span style="font-size:16px; font-weight:bold;">¥{q_tot_sales:,}</span> &nbsp;|&nbsp;
                <b>送料(税抜):</b> ¥{q_shipping_fee:,} &nbsp;|&nbsp;
                <b>消費税:</b> ¥{q_tax:,} &nbsp;|&nbsp;
                <b>税込合計額:</b> <span style="color:#d93025; font-size:18px; font-weight:bold;">¥{q_grand_total:,}</span> &nbsp;|&nbsp;
                <b>粗利額:</b> <span style="color:#188038; font-size:16px; font-weight:bold;">¥{q_tot_profit:,}</span> (粗利率: <b>{q_tot_margin:.1f}%</b>)
            </div>
            """, unsafe_allow_html=True)

            row_cust_master = cust_lookup.get(q_cust_comp, {})
            cust_id_val = str(row_cust_master.get("加盟店ID", row_cust_master.get("顧客ID", -1))).replace(".0", "") if isinstance(row_cust_master, pd.Series) else -1
            cust_sei_val = str(row_cust_master.get("責任者姓", "ご担当者")).strip() if isinstance(row_cust_master, pd.Series) and pd.notna(row_cust_master.get("責任者姓")) else "ご担当者"
            cust_mei_val = str(row_cust_master.get("責任者名", "様")).strip() if isinstance(row_cust_master, pd.Series) and pd.notna(row_cust_master.get("責任者名")) else "様"
            cust_k_sei = str(row_cust_master.get("責任者姓カナ", "")).strip() if isinstance(row_cust_master, pd.Series) and pd.notna(row_cust_master.get("責任者姓カナ")) else ""
            cust_k_mei = str(row_cust_master.get("責任者名カナ", "")).strip() if isinstance(row_cust_master, pd.Series) and pd.notna(row_cust_master.get("責任者名カナ")) else ""
            cust_email_val = str(row_cust_master.get("メールアドレス", "")).strip() if isinstance(row_cust_master, pd.Series) and pd.notna(row_cust_master.get("メールアドレス")) else ""
            cust_fax_val = str(row_cust_master.get("FAX番号", "")).strip() if isinstance(row_cust_master, pd.Series) and pd.notna(row_cust_master.get("FAX番号")) else ""
            cust_pref_val = str(row_cust_master.get("都道府県", "")).strip() if isinstance(row_cust_master, pd.Series) and pd.notna(row_cust_master.get("都道府県")) else ""
            cust_addr1_val = str(row_cust_master.get("住所１", q_cust_addr)).strip() if isinstance(row_cust_master, pd.Series) and pd.notna(row_cust_master.get("住所１")) else q_cust_addr
            cust_addr2_val = str(row_cust_master.get("住所２", "")).strip() if isinstance(row_cust_master, pd.Series) and pd.notna(row_cust_master.get("住所２")) else ""
            cust_rank_label = str(row_cust_master.get("加盟店ランク", "ロット単価")).strip() if isinstance(row_cust_master, pd.Series) and pd.notna(row_cust_master.get("加盟店ランク")) else "ロット単価"

            current_estimate_obj = {
                "id": len(st.session_state.saved_estimates) + 840, "q_no": q_no, "staff": est_staff,
                "created_at": datetime.now().strftime("%Y/%m/%d %H:%M"), "q_date": q_date.strftime("%Y/%m/%d"),
                "q_exp_date": q_exp_date.strftime("%Y/%m/%d"), "payment": q_payment, "cust_id": cust_id_val,
                "cust_comp": q_cust_comp, "cust_person": q_cust_person, "cust_sei": cust_sei_val, "cust_mei": cust_mei_val,
                "cust_kana_sei": cust_k_sei, "cust_kana_mei": cust_k_mei, "cust_post": q_cust_post, "cust_pref": cust_pref_val,
                "cust_addr": q_cust_addr, "cust_addr1": cust_addr1_val, "cust_addr2": cust_addr2_val, "cust_tel": ship_cust_tel,
                "cust_fax": cust_fax_val, "cust_email": cust_email_val, "rank_name": cust_rank_label, "ship_comp": ship_cust_comp,
                "ship_person": ship_cust_person, "ship_sei": cust_sei_val, "ship_mei": cust_mei_val, "ship_post": ship_cust_post,
                "ship_pref": cust_pref_val, "ship_addr": ship_cust_addr, "ship_addr1": ship_cust_addr, "ship_addr2": "",
                "ship_tel": ship_cust_tel, "shipping_fee": q_shipping_fee, "cod_fee": 0, "sub_total": q_tot_sales,
                "tax": q_tax, "grand_total": q_grand_total, "memo": q_memo, "hitokoto": q_staff_memo, "items": updated_records
            }

            def generate_quote_excel_exact_style(df_calc):
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = "見積書"
                font_title, font_section, font_main, font_bold = Font(name="Yu Gothic", size=18, bold=True), Font(name="Yu Gothic", size=11, bold=True), Font(name="Yu Gothic", size=9), Font(name="Yu Gothic", size=9, bold=True)
                thin, thick = Side(style='thin', color='A0A0A0'), Side(style='medium', color='000000')
                border_cell, border_header = Border(left=thin, right=thin, top=thin, bottom=thin), Border(left=thin, right=thin, top=thick, bottom=thick)
                fill_gray = PatternFill(start_color="E9ECEF", end_color="E9ECEF", fill_type="solid")

                ws.merge_cells("A2:F2")
                ws["A2"] = "見 積 書"; ws["A2"].font = font_title; ws["A2"].alignment = Alignment(horizontal="center", vertical="center")
                for col_i in range(1, 7): ws.cell(row=3, column=col_i).border = Border(bottom=thick)

                ws["A5"] = f"〒 {q_cust_post}" if q_cust_post else ""; ws["A6"] = q_cust_addr; ws["A7"] = q_cust_comp; ws["A8"] = q_cust_person
                ws["A5"].font = ws["A6"].font = font_main; ws["A7"].font = ws["A8"].font = font_section
                ws["D7"] = "スマイルキッズ株式会社"; ws["D7"].font = font_bold
                ws["D8"] = "スマイルキッズ法人卸サイト"; ws["D9"] = "住所: 〒537-0003 大阪市東成区神路4丁目3番18号"
                ws["D10"] = "TEL: 06-6976-1371  FAX: 06-6976-8940"; ws["D11"] = "Mail: smilekids-company@smile-asahi.co.jp"; ws["D12"] = "登録番号: T8120001247513"
                for r_i in range(8, 13): ws[f"D{r_i}"].font = font_main

                if hanko_bytes:
                    try:
                        img = OpenpyxlImage(io.BytesIO(hanko_bytes))
                        img.width = img.height = 65
                        ws.add_image(img, "F7")
                    except Exception: pass

                ws.merge_cells("A14:F14")
                ws["A14"] = " 見積内容"; ws["A14"].font = font_section; ws["A14"].fill = fill_gray; ws["A14"].border = border_cell
                ws["A15"] = "■見積金額"; ws["B15"] = f":  {q_grand_total:,} 円 (税込)"; ws["A15"].font = ws["B15"].font = font_bold
                ws["A16"] = "■お支払い方法"; ws["B16"] = f":  {q_payment}"; ws["A16"].font = font_bold; ws["B16"].font = font_main

                ws.merge_cells("A18:F18")
                ws["A18"] = " お見積り明細"; ws["A18"].font = font_section; ws["A18"].fill = fill_gray; ws["A18"].border = border_cell
                ws["A19"] = "■お届け先:"; ws["B19"] = f"〒 {ship_cust_post}" if ship_cust_post else ""
                ws["B20"] = ship_cust_addr; ws["B21"] = ship_cust_comp; ws["B22"] = ship_cust_person; ws["B23"] = f"TEL: {ship_cust_tel}" if ship_cust_tel else ""
                ws["A19"].font = font_bold
                for r_i in range(19, 24): ws.cell(row=r_i, column=2).font = font_main

                ws["D19"] = "■お見積日"; ws["E19"] = f": {q_date.strftime('%Y/%m/%d')}"
                ws["D20"] = "■見積番号"; ws["E20"] = f": {q_no}"
                ws["D21"] = "■見積有効期限"; ws["E21"] = f": {q_exp_date.strftime('%Y/%m/%d')}"
                for r_i in range(19, 22):
                    ws.cell(row=r_i, column=4).font = font_bold
                    ws.cell(row=r_i, column=5).font = font_main

                start_row = 25
                for c_idx, h in enumerate(["No.", "品名・規格", "JANコード", "数量", "単価", "金額"], 1):
                    cell = ws.cell(row=start_row, column=c_idx, value=h)
                    cell.font = font_bold; cell.fill = fill_gray; cell.border = border_header; cell.alignment = Alignment(horizontal="center", vertical="center")

                cur_r = start_row + 1
                for idx_c, r in df_calc.iterrows():
                    ws.cell(row=cur_r, column=1, value=idx_c + 1).alignment = Alignment(horizontal="center")
                    ws.cell(row=cur_r, column=2, value=f"{r['商品名']} ({r['商品コード']})").alignment = Alignment(horizontal="left")
                    ws.cell(row=cur_r, column=3, value=str(r['JANコード'])).alignment = Alignment(horizontal="center")
                    ws.cell(row=cur_r, column=4, value=r['数量']).number_format = '#,##0'
                    ws.cell(row=cur_r, column=5, value=r['見積単価']).number_format = '#,##0'
                    ws.cell(row=cur_r, column=6, value=r['合計金額']).number_format = '#,##0'
                    for c_i in range(1, 7):
                        c = ws.cell(row=cur_r, column=c_i); c.font = font_main; c.border = border_cell
                    cur_r += 1

                ws.merge_cells(start_row=cur_r, start_column=1, end_row=cur_r, end_column=3)
                tot_sub_lbl = ws.cell(row=cur_r, column=1, value="小計 / 合計")
                tot_sub_lbl.alignment = Alignment(horizontal="center", vertical="center"); tot_sub_lbl.font = font_bold; tot_sub_lbl.fill = fill_gray
                for col_i, v_num in [(4, q_tot_qty), (5, q_tot_unit_price), (6, q_tot_sales)]:
                    c_t = ws.cell(row=cur_r, column=col_i, value=v_num)
                    c_t.number_format = '#,##0'; c_t.font = font_bold; c_t.fill = fill_gray
                for c_i in range(1, 7): ws.cell(row=cur_r, column=c_i).border = border_cell
                cur_r += 1

                for s_lbl, s_val in [("税抜金額合計", q_tot_sales), ("送料 (税抜)", q_shipping_fee), ("消費税 (10%)", q_tax), ("税込合計金額", q_grand_total)]:
                    ws.merge_cells(start_row=cur_r, start_column=4, end_row=cur_r, end_column=5)
                    lbl_c = ws.cell(row=cur_r, column=4, value=s_lbl)
                    lbl_c.alignment = Alignment(horizontal="center", vertical="center"); lbl_c.font = font_bold; lbl_c.fill = fill_gray; lbl_c.border = border_cell
                    ws.cell(row=cur_r, column=5).border = border_cell
                    val_c = ws.cell(row=cur_r, column=6, value=s_val)
                    val_c.number_format = '#,##0'; val_c.font = font_bold; val_c.border = border_cell
                    cur_r += 1

                for col_l, w_v in [('A', 6), ('B', 38), ('C', 16), ('D', 12), ('E', 13), ('F', 15)]:
                    ws.column_dimensions[col_l].width = w_v
                out = io.BytesIO()
                wb.save(out)
                return out.getvalue()

            hanko_base64 = base64.b64encode(hanko_bytes).decode("utf-8") if hanko_bytes else ""
            pdf_rows_html = "".join([
                f"<tr><td style='text-align:center;'>{idx + 1}</td><td style='text-align:left;'>{r['商品名']} ({r['商品コード']})</td>"
                f"<td style='text-align:center;'>{r['JANコード']}</td><td style='text-align:right;'>{r['数量']:,}</td>"
                f"<td style='text-align:right;'>¥{r['見積単価']:,}</td><td style='text-align:right;'>¥{r['合計金額']:,}</td></tr>"
                for idx, r in calc_df.iterrows()
            ])
            hanko_img_tag = f'<img src="data:image/png;base64,{hanko_base64}" style="position:absolute; right:15px; top:5px; width:75px; height:75px; opacity:0.92;" />' if hanko_bytes else ''

            pdf_template_html = f"""
            <!DOCTYPE html><html><head><meta charset="utf-8"><title>御見積書_{q_cust_comp}</title>
            <style>
                @media print {{ @page {{ size: A4 portrait; margin: 12mm; }} body {{ -webkit-print-color-adjust: exact; }} .no-print {{ display: none !important; }} }}
                body {{ font-family: "Yu Gothic", "Meiryo", sans-serif; color: #111827; padding: 10px; max-width: 800px; margin: 0 auto; }}
                .title {{ text-align: center; font-size: 26px; font-weight: bold; letter-spacing: 12px; padding-bottom: 6px; border-bottom: 2px solid #111; margin-bottom: 20px; }}
                .header-flex {{ display: flex; justify-content: space-between; margin-bottom: 15px; }}
                .cust-area {{ font-size: 14px; line-height: 1.6; width: 55%; }}
                .cust-comp, .cust-person {{ font-size: 17px; font-weight: bold; margin-top: 2px; }}
                .company-area {{ position: relative; font-size: 12px; line-height: 1.5; width: 45%; padding-left: 20px; }}
                .company-name {{ font-size: 15px; font-weight: bold; }}
                .section-header {{ background: #e5e7eb; border: 1px solid #9ca3af; padding: 4px 10px; font-weight: bold; font-size: 14px; margin-top: 15px; margin-bottom: 10px; }}
                .content-box {{ font-size: 14px; line-height: 1.8; margin-left: 10px; }}
                .quote-table {{ width: 100%; border-collapse: collapse; margin-top: 10px; font-size: 13px; }}
                .quote-table th {{ background: #f3f4f6; border: 1px solid #9ca3af; padding: 7px 6px; }}
                .quote-table td {{ border: 1px solid #d1d5db; padding: 7px 8px; font-variant-numeric: tabular-nums; }}
                .summary-table {{ width: 340px; margin-left: auto; margin-top: 10px; border-collapse: collapse; font-size: 13px; }}
                .summary-table td {{ border: 1px solid #9ca3af; padding: 6px 10px; }}
                .btn-print {{ background: #2563eb; color: #fff; padding: 10px 22px; border: none; border-radius: 6px; font-size: 15px; font-weight: bold; cursor: pointer; }}
            </style></head><body>
                <div class="no-print" style="text-align:right; margin-bottom:12px;"><button class="btn-print" onclick="window.print();">🖨️ PDFとして保存 / 印刷する</button></div>
                <div class="title">見積書</div>
                <div class="header-flex">
                    <div class="cust-area"><div>{'〒 ' + q_cust_post if q_cust_post else ''}</div><div>{q_cust_addr}</div><div class="cust-comp">{q_cust_comp}</div><div class="cust-person">{q_cust_person}</div></div>
                    <div class="company-area">{hanko_img_tag}<div class="company-name">スマイルキッズ株式会社</div><div>住所: 〒537-0003 大阪市東成区神路4丁目3番18号</div><div>TEL : 06-6976-1371 &nbsp; FAX : 06-6976-8940</div></div>
                </div>
                <div class="section-header">見積内容</div>
                <div class="content-box"><div>■見積金額 : <b>{q_grand_total:,} 円</b> (税込)</div><div>■お支払い方法 : {q_payment}</div></div>
                <table class="quote-table"><thead><tr><th style="width:35px;">No.</th><th>品名・規格</th><th style="width:140px;">JANコード</th><th style="width:75px;">数量</th><th style="width:90px;">単価</th><th style="width:105px;">金額</th></tr></thead>
                <tbody>{pdf_rows_html}<tr style="background:#f9fafb; font-weight:bold; border-top:2px solid #9ca3af;"><td colspan="3" style="text-align:center;">小計 / 合計</td><td style="text-align:right;">{q_tot_qty:,}</td><td style="text-align:right;">¥{q_tot_unit_price:,}</td><td style="text-align:right;">¥{q_tot_sales:,}</td></tr></tbody></table>
                <table class="summary-table">
                    <tr><td style="background:#f3f4f6; text-align:center; font-weight:bold;">税抜金額合計</td><td style="text-align:right; font-weight:bold;">¥{q_tot_sales:,}</td></tr>
                    <tr><td style="background:#f3f4f6; text-align:center; font-weight:bold;">送料 (税抜)</td><td style="text-align:right; font-weight:bold;">¥{q_shipping_fee:,}</td></tr>
                    <tr><td style="background:#f3f4f6; text-align:center; font-weight:bold;">消費税 (10%)</td><td style="text-align:right; font-weight:bold;">¥{q_tax:,}</td></tr>
                    <tr style="background:#f9fafb;"><td style="background:#e5e7eb; text-align:center; font-weight:bold; font-size:14px;">税込合計金額</td><td style="text-align:right; font-weight:bold; font-size:15px; color:#d93025;">¥{q_grand_total:,}</td></tr>
                </table>
            </body></html>
            """

            file_prefix = q_cust_comp.replace(' ', '_') if q_cust_comp else "御見積書"
            btn_c1, btn_c2, btn_c3 = st.columns([1.4, 1.4, 1.6])
            with btn_c1:
                if st.button("💾 この見積を保存して一覧へ登録", use_container_width=True, key=f"save_btn_{s_key}"):
                    st.session_state.saved_estimates.insert(0, current_estimate_obj)
                    save_quotes_list(st.session_state.saved_estimates)
                    st.success("✅ 保存済み見積一覧に登録しました！")
            with btn_c2:
                st.download_button(
                    label="📥 御見積書（Excel・角印付き）",
                    data=generate_quote_excel_exact_style(calc_df),
                    file_name=f"御見積書_{file_prefix}_{datetime.now().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True, key=f"dl_excel_{s_key}"
                )
            with btn_c3:
                st.download_button(
                    label="🛒 注文確定：スマレジ取込用CSV出力",
                    data=create_smaregi_order_csv_bytes(current_estimate_obj, raw_prod_title_map, jan_std_price_dict),
                    file_name=f"order_{q_no}_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                    mime="text/csv", use_container_width=True, key=f"dl_csv_{s_key}"
                )

            st.write("")
            st.markdown("##### 👁️ PDFプレビュー表示＆ワンクリックPDF保存")
            components.html(pdf_template_html, height=720, scrolling=True)

    elif tab_choice.startswith("📋"):
        st.markdown(f"##### 📋 保存済み見積一覧（担当: {est_staff}）")
        if not staff_Filter_list:
            st.info(f"担当「{est_staff}」の保存された見積もりはありません。")
        else:
            search_est_kw = st.text_input("🔍 見積番号・会社名・商品名で絞り込み", placeholder="例: アーネスト, ナスコム, ABA-105...")
            for idx_q, est in enumerate(st.session_state.saved_estimates):
                if est_staff != "（すべて表示）" and est.get("staff", "") != est_staff and (est_staff not in str(est.get("hitokoto", ""))):
                    continue
                if search_est_kw.strip() and search_est_kw.strip() not in json.dumps(est, ensure_ascii=False):
                    continue

                est_comp_name = est.get("cust_comp", "") or est.get("ship_comp", "")
                matched_grp = lookup_group_info_by_comp(est_comp_name, member_group_master_dict)
                g_no_est = matched_grp.get("group_no", "")
                m_cd_est = matched_grp.get("member_code", "")
                grp_badge_str = f" &nbsp;|&nbsp; <span style='color:#1d4ed8; font-size:12px;'>[グループ連番: <b>{g_no_est or '未検出'}</b> / 加盟店コード: <b>{m_cd_est or '未検出'}</b>]</span>"

                st.markdown(f"""
                <div class="est-card"><div class="est-header">
                    <div><span class="badge-green">店</span> &nbsp;<b>見積番号: {est.get('q_no', '')}</b> &nbsp;|&nbsp; <b>担当:</b> {est.get('staff', est_staff)} &nbsp;|&nbsp; <b>見積日時:</b> {est.get('created_at', '')}{grp_badge_str}</div>
                    <div><b>税込見積総額:</b> <span style="color:#d93025; font-size:16px; font-weight:bold;">¥{int(est.get('grand_total', 0)):,}</span></div>
                </div></div>
                """, unsafe_allow_html=True)

                c_info1, c_info2, c_items = st.columns([1.5, 1.5, 2.5])
                with c_info1:
                    st.markdown("**【宛先情報】**")
                    st.write(f"🏢 **{est.get('cust_comp', '')}** / 👤 {est.get('cust_person', '')}")
                    st.write(f"〒{est.get('cust_post', '')} {est.get('cust_addr', '')}")
                with c_info2:
                    st.markdown("**【お届け先・備考】**")
                    st.write(f"🏢 {est.get('ship_comp', '')}")
                    if est.get('memo'): st.info(f"伝達事項: {est.get('memo')}")
                with c_items:
                    items_df = pd.DataFrame(est.get("items", []))
                    if not items_df.empty:
                        st.dataframe(items_df[[c for c in ["商品コード", "商品名", "数量", "見積単価", "合計金額", "粗利額", "粗利率"] if c in items_df.columns]], use_container_width=True, hide_index=True)

                b_col1, b_col_lot_mode, b_col_lot_dl, b_col2, b_col3 = st.columns([1.8, 0.9, 1.9, 1.3, 0.8])
                with b_col1:
                    st.download_button(
                        label="🛒 注文確定：スマレジ用CSV出力",
                        data=create_smaregi_order_csv_bytes(est, raw_prod_title_map, jan_std_price_dict),
                        file_name=f"order_{est.get('q_no')}_{datetime.now().strftime('%Y%m%d')}.csv",
                        mime="text/csv", use_container_width=True, key=f"dl_saved_csv_{idx_q}_{est.get('q_no')}"
                    )
                with b_col_lot_mode:
                    sel_est_ctrl = st.selectbox(
                        "A列区分",
                        ["更新 (u)", "新規 (n)"],
                        index=0,
                        key=f"est_lot_ctrl_{idx_q}_{est.get('q_no')}",
                        label_visibility="collapsed"
                    )
                with b_col_lot_dl:
                    ctrl_code = "n" if "n" in sel_est_ctrl else "u"
                    st.download_button(
                        label="🏷️ 価格登録CSV出力 (lotmember)",
                        data=create_lotmember_csv_bytes(est.get("items", []), g_no_est, m_cd_est, lotmember_sku_meta_dict, default_ctrl=ctrl_code),
                        file_name="item_class_lotmember.csv",
                        mime="text/csv",
                        type="primary",
                        use_container_width=True,
                        key=f"dl_saved_lotcsv_{idx_q}_{est.get('q_no')}"
                    )
                with b_col2:
                    if st.button("✏️ 編集画面に読込", use_container_width=True, key=f"load_est_{idx_q}"):
                        st.session_state["load_target_estimate"] = est
                        st.rerun()
                with b_col3:
                    if st.button("🗑️ 削除", use_container_width=True, key=f"del_est_{idx_q}"):
                        st.session_state.saved_estimates.pop(idx_q)
                        save_quotes_list(st.session_state.saved_estimates)
                        st.rerun()
                st.markdown("<hr style='margin:12px 0; border:0; border-top:1px dashed #cbd5e1;'>", unsafe_allow_html=True)

    st.stop()

# ========================================================
# ★ 🏷️ ランク価格マスタ一覧（指定12ランク限定・ロット数列削除）
# ========================================================
if view_type == '🏷️ ランク価格マスタ一覧':
    raw_rank_path, raw_rank_mtime = get_rank_price_file_info()
    raw_rank_df = load_rank_price_master(raw_rank_path, raw_rank_mtime)
    if raw_rank_df.empty and member_p_path and os.path.exists(member_p_path):
        raw_rank_df = load_rank_price_master(member_p_path, member_p_mtime)
    if raw_rank_df.empty:
        st.error("ランク価格データが見つかりませんでした。")
        st.stop()

    code_to_name_map = df.dropna(subset=["商品コード", "商品名"]).drop_duplicates(subset=["商品コード"]).set_index("商品コード")["商品名"].to_dict()
    for k_t, v_t in jan_title_dict.items():
        if v_t and str(v_t).strip() not in ["", "nan"]:
            code_to_name_map[str(k_t).strip()] = str(v_t).strip()

    rank_col = next((c for c in raw_rank_df.columns if "ランク名" in str(c)), None)
    if rank_col is None:
        rank_col = next((c for c in raw_rank_df.columns if any(k in str(c) for k in ["ランク", "会社名", "取引先", "グループ"])), None)
    sku_col = next((c for c in raw_rank_df.columns if any(k in str(c) for k in ["商品番号", "商品コード", "商品管理番号", "SKU", "品番"])), raw_rank_df.columns[0])
    price_col = next((c for c in raw_rank_df.columns if any(k in str(c) for k in ["ロット単価", "単価", "下代", "価格"])), None)
    cost_col = next((c for c in raw_rank_df.columns if "原価" in str(c)), None)

    work_rk = pd.DataFrame()
    work_rk["商品番号"] = raw_rank_df[sku_col].astype(str).str.strip()
    work_rk = work_rk[~work_rk["商品番号"].isin(["", "nan", "None"])].copy()
    work_rk["商品名"] = raw_rank_df["商品名"].astype(str).str.strip() if "商品名" in raw_rank_df.columns else work_rk["商品番号"].map(lambda c: code_to_name_map.get(c, jan_title_dict.get(c, "")))
    work_rk["ランク名"] = raw_rank_df[rank_col].astype(str).str.strip() if rank_col is not None else "1.1倍"
    work_rk = work_rk[work_rk["ランク名"].isin(ALLOWED_EXACT_RANKS)].copy()
    work_rk["ロット単価"] = raw_rank_df[price_col].apply(clean_num).astype(int) if price_col is not None else 0
    work_rk["原価"] = raw_rank_df[cost_col].apply(clean_num).astype(int) if cost_col is not None else work_rk["商品番号"].map(lambda c: int(cost_master_dict.get(c, 0)))
    work_rk = work_rk[(work_rk["ロット単価"] > 0) & (work_rk["ロット単価"] < 90000)].copy()

    r_col1, r_col2, r_col3 = st.columns([1.5, 1.5, 1.0])
    with r_col1:
        avail_ranks = [rk for rk in ALLOWED_EXACT_RANKS if rk in work_rk["ランク名"].values]
        sel_rank = st.selectbox("🎯 ランクで絞り込み", ["（すべて）"] + avail_ranks)
    with r_col2:
        search_rank_prod = st.text_input("🔍 商品番号・商品名で検索", placeholder="例：ABO-03...")
    with r_col3:
        rank_limit_val = st.selectbox("表示件数", [50, 100, 200, 500, "全件"], index=0)

    filt_rank_df = work_rk.copy()
    if sel_rank != "（すべて）": filt_rank_df = filt_rank_df[filt_rank_df["ランク名"] == sel_rank]
    if search_rank_prod.strip():
        kw_rp = search_rank_prod.strip()
        filt_rank_df = filt_rank_df[filt_rank_df["商品番号"].str.contains(kw_rp, case=False, na=False) | filt_rank_df["商品名"].str.contains(kw_rp, case=False, na=False)]

    filt_rank_df["ロット粗利額"] = filt_rank_df["ロット単価"] - filt_rank_df["原価"]
    filt_rank_df["粗利率"] = filt_rank_df.apply(lambda r: f"{(r['ロット粗利額'] / r['ロット単価'] * 100):.1f}%" if r['ロット単価'] > 0 else "-", axis=1)

    disp_rank_cols = ["商品番号", "商品名", "ランク名", "ロット単価", "原価", "ロット粗利額", "粗利率"]
    out_rank_df = filt_rank_df[disp_rank_cols].reset_index(drop=True)
    out_rank_df.insert(0, "番号", range(1, len(out_rank_df) + 1))

    r_top_c1, r_top_c2 = st.columns([4, 1.5])
    with r_top_c1:
        st.markdown(f"<div class='total-card' style='margin-bottom:0;'><b>対象ランク:</b> <span style='color:#1a73e8; font-weight:bold;'>{sel_rank}</span> &nbsp;|&nbsp; <b>該当件数:</b> <span style='font-size:16px; font-weight:bold; color:#d93025;'>{len(out_rank_df):,} 件</span></div>", unsafe_allow_html=True)
    with r_top_c2:
        st.download_button(label="📥 ランク価格一覧をExcel保存", data=generate_excel_bytes(out_rank_df, "ランク価格一覧"), file_name=f"ランク価格_{sel_rank}_{datetime.now().strftime('%Y%m%d')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

    st.write("")
    if rank_limit_val != "全件": out_rank_df = out_rank_df.head(int(rank_limit_val))
    render_universal_table(out_rank_df)
    st.stop()

# ========================================================
# ★ 期間設定（売上分析用）
# ========================================================
idx = valid_ym.index(latest_ym) if latest_ym in valid_ym else len(valid_ym) - 1
curr_y, curr_m = int(latest_ym[:4]), int(latest_ym[5:7])
fiscal_year = curr_y if curr_m >= 7 else curr_y - 1

if quick_choice == '当月':
    s_def, e_def = latest_ym, latest_ym
    s_date_calc, e_date_calc = datetime(curr_y, curr_m, 1).date(), max_date
elif quick_choice == '前月':
    prev = valid_ym[max(0, idx - 1)]
    s_def, e_def = prev, prev
    py, pm = int(prev[:4]), int(prev[5:7])
    s_date_calc = datetime(py, pm, 1).date()
    e_date_calc = (pd.Timestamp(year=py, month=pm, day=1) + pd.offsets.MonthEnd(1)).date()
elif quick_choice == '当期累計':
    s_def, e_def = f"{fiscal_year}年07月", latest_ym
    s_date_calc, e_date_calc = datetime(fiscal_year, 7, 1).date(), max_date
elif quick_choice == '前期累計':
    s_def, e_def = f"{fiscal_year - 1}年07月", f"{fiscal_year}年06月"
    s_date_calc, e_date_calc = datetime(fiscal_year - 1, 7, 1).date(), datetime(fiscal_year, 6, 30).date()
elif quick_choice == '全データ累計':
    s_def, e_def = valid_ym[0], valid_ym[-1]
    s_date_calc, e_date_calc = min_date, max_date
else:
    s_def, e_def = valid_ym[0], latest_ym
    s_date_calc, e_date_calc = min_date, max_date

if "prev_quick_choice" not in st.session_state or st.session_state.prev_quick_choice != quick_choice:
    st.session_state.prev_quick_choice = quick_choice
    if quick_choice != "手動指定":
        sync_ym = s_def if s_def in desc_valid_ym else desc_valid_ym[0]
        st.session_state["analysis_ym_start_v3"] = sync_ym
        st.session_state["analysis_ym_end_v3"] = e_def if e_def in desc_valid_ym else desc_valid_ym[0]
        st.session_state["tgt_ym_meeting_v3"] = sync_ym
        st.session_state["tgt_ym_topcust_v3"] = sync_ym

if "analysis_ym_start_v3" not in st.session_state:
    st.session_state["analysis_ym_start_v3"] = s_def if s_def in desc_valid_ym else desc_valid_ym[0]
if "analysis_ym_end_v3" not in st.session_state:
    st.session_state["analysis_ym_end_v3"] = e_def if e_def in desc_valid_ym else desc_valid_ym[0]

if view_type not in ['📑 営業会議用：上位客先別売上実績', '📊 営業会議用：上位社 売上比較シート']:
    col_mode, col_p1, col_p2, col_staff = st.columns([1, 1.2, 1.2, 1.6])
    with col_mode:
        date_mode = st.radio("指定単位", ["年月で指定", "日で指定"], horizontal=True, key="date_mode_radio_analysis")
    if date_mode == "年月で指定":
        with col_p1: ym_start = st.selectbox("開始年月", desc_valid_ym, key="analysis_ym_start_v3")
        with col_p2: ym_end = st.selectbox("終了年月", desc_valid_ym, key="analysis_ym_end_v3")
        st.session_state["tgt_ym_meeting_v3"] = ym_start
        st.session_state["tgt_ym_topcust_v3"] = ym_start
    else:
        with col_p1: d_start = st.date_input("開始日", value=s_date_calc, min_value=min_date, max_value=max_date)
        with col_p2: d_end = st.date_input("終了日", value=e_date_calc, min_value=min_date, max_value=max_date)

    with col_staff:
        staff_options = ["（すべて）"] + sorted([str(x).strip() for x in df["担当者"].dropna().unique() if str(x).strip() not in ["", "nan", "None"]])
        sel_staff = st.selectbox("担当者指定", staff_options)

    all_unique_custs = sorted([str(x).strip() for x in df["得意先名"].dropna().unique() if str(x).strip() not in ["", "nan", "None"]])
    c1, c2, c3, c4 = st.columns(4)
    with c1: sel_cust = st.selectbox("得意先名", ["（すべて）"] + all_unique_custs)
    with c2: search_cust = st.text_input("得意先検索", placeholder="会社名の一部を入力（例：丸藤、アマゾン）...")
    with c3:
        prod_pairs = df[["商品コード", "商品名"]].dropna().drop_duplicates()
        prod_list = ["（すべて）"] + sorted([f"{str(r['商品コード']).strip()} : {str(r['商品名']).strip()}" for _, r in prod_pairs.iterrows() if str(r['商品コード']).strip() not in ["", "nan", "None"]])
        sel_prod = st.selectbox("商品選択", prod_list)
    with c4: search_prod = st.text_input("商品検索", placeholder="商品コード・品名の一部を入力...")

    matched_custs = []
    chosen_sub_branch = "（すべて合算）"
    if search_cust.strip():
        kw_clean = re.sub(r'[\s\u3000\xa0]+', '', search_cust.strip())
        matched_custs = [c for c in all_unique_custs if kw_clean in re.sub(r'[\s\u3000\xa0]+', '', c)]
    if len(matched_custs) > 1:
        st.info(f"💡 「{search_cust.strip()}」で **{len(matched_custs)} 件** の得意先が見つかりました：")
        chosen_sub_branch = st.selectbox("🎯 対象の取引先を選択", [f"（該当する全{len(matched_custs)}件を合算表示）"] + matched_custs)

    limit_col, _, _, _ = st.columns(4)
    with limit_col: limit_val = st.selectbox("表示件数", [50, 100, 200, 500, "全件"], index=0)
else:
    date_mode = "年月で指定"
    ym_start = st.session_state.get("analysis_ym_start_v3", latest_ym)
    ym_end = st.session_state.get("analysis_ym_end_v3", latest_ym)
    sel_staff, sel_cust, search_cust, sel_prod, search_prod, limit_val, matched_custs, chosen_sub_branch = "（すべて）", "（すべて）", "", "（すべて）", "", 50, [], ""

prod_first_seen_dt = df.groupby("商品コード")["日付_dt"].min().to_dict()
code_to_name = df.dropna(subset=["商品コード", "商品名"]).drop_duplicates(subset=["商品コード"]).set_index("商品コード")["商品名"].to_dict()
for k_t, v_t in (jan_title_dict or {}).items():
    if v_t and str(v_t).strip() not in ["", "nan"]:
        code_to_name[str(k_t).strip()] = str(v_t).strip()

one_year_ago_date = pd.Timestamp(max_date) - pd.DateOffset(years=1)
active_months_list = []

if view_type in ['📑 営業会議用：上位客先別売上実績', '📊 営業会議用：上位社 売上比較シート']:
    filtered, ly_filtered, period_text = df.copy(), pd.DataFrame(), ""
elif view_type == '✨ 新商品売上分析（直近1年）':
    true_new_skus = {sku for sku, f_dt in prod_first_seen_dt.items() if pd.notna(f_dt) and f_dt >= one_year_ago_date}
    filtered = df[(df["日付_dt"] >= one_year_ago_date) & (df["商品コード"].isin(true_new_skus))].copy()
    period_text = f"{one_year_ago_date.strftime('%Y/%m/%d')} 〜 {max_date.strftime('%Y/%m/%d')}（直近1年間に初登場した新商品）"
    ly_filtered = pd.DataFrame()
else:
    if date_mode == "年月で指定":
        s_ym_calc, e_ym_calc = (ym_end, ym_start) if ym_start > ym_end else (ym_start, ym_end)
        filtered = df[(df["年月"] >= s_ym_calc) & (df["年月"] <= e_ym_calc)]
        period_text = f"{s_ym_calc} 〜 {e_ym_calc}" if s_ym_calc != e_ym_calc else f"{s_ym_calc}（単月）"
        s_y, s_m = int(s_ym_calc[:4]), int(s_ym_calc[5:7])
        e_y, e_m = int(e_ym_calc[:4]), int(e_ym_calc[5:7])
        ly_filtered = df[(df["年月"] >= f"{s_y - 1}年{s_m:02d}月") & (df["年月"] <= f"{e_y - 1}年{e_m:02d}月")]
        active_months_list = [int(ym[5:7]) for ym in valid_ym if s_ym_calc <= ym <= e_ym_calc]
    else:
        dt_s, dt_e = pd.to_datetime(d_start), pd.to_datetime(d_end)
        if dt_s > dt_e: dt_s, dt_e = dt_e, dt_s
        filtered = df[(df["日付_dt"] >= dt_s) & (df["日付_dt"] <= dt_e)]
        period_text = f"{dt_s.strftime('%Y/%m/%d')} 〜 {dt_e.strftime('%Y/%m/%d')}"
        ly_filtered = df[(df["日付_dt"] >= (dt_s - pd.DateOffset(years=1))) & (df["日付_dt"] <= (dt_e - pd.DateOffset(years=1)))]
        active_months_list = [int(ym[5:7]) for ym in sorted(filtered["年月"].dropna().unique()) if len(ym) >= 7]

if view_type not in ['📑 営業会議用：上位客先別売上実績', '📊 営業会議用：上位社 売上比較シート', '🛒 アマゾン日別注文数＆利益計算出力']:
    if sel_cust != "（すべて）":
        filtered = filtered[filtered["得意先名"] == sel_cust]
        if not ly_filtered.empty: ly_filtered = ly_filtered[ly_filtered["得意先名"] == sel_cust]
    if search_cust.strip():
        kw_clean = re.sub(r'[\s\u3000\xa0]+', '', search_cust.strip())
        if len(matched_custs) > 1 and not chosen_sub_branch.startswith("（該当する全"):
            filtered = filtered[filtered["得意先名"] == chosen_sub_branch]
            if not ly_filtered.empty: ly_filtered = ly_filtered[ly_filtered["得意先名"] == chosen_sub_branch]
        else:
            filtered = filtered[filtered["得意先名"].map(lambda x: kw_clean in re.sub(r'[\s\u3000\xa0]+', '', str(x)))]
            if not ly_filtered.empty: ly_filtered = ly_filtered[ly_filtered["得意先名"].map(lambda x: kw_clean in re.sub(r'[\s\u3000\xa0]+', '', str(x)))]
    if sel_staff != "（すべて）":
        filtered = filtered[filtered["担当者"] == sel_staff]
        if not ly_filtered.empty: ly_filtered = ly_filtered[ly_filtered["担当者"] == sel_staff]
    if sel_prod != "（すべて）":
        code = sel_prod.split(" : ")[0]
        filtered = filtered[filtered["商品コード"] == code]
        if not ly_filtered.empty: ly_filtered = ly_filtered[ly_filtered["商品コード"] == code]
    if search_prod.strip():
        kw_p = search_prod.strip()
        filtered = filtered[filtered["商品名"].str.contains(kw_p, na=False) | filtered["商品コード"].str.contains(kw_p, na=False)]
        if not ly_filtered.empty:
            ly_filtered = ly_filtered[ly_filtered["商品名"].str.contains(kw_p, na=False) | ly_filtered["商品コード"].str.contains(kw_p, na=False)]

    total_sales = int(filtered["金額"].sum()) if not filtered.empty else 0

# ========================================================
# ★ 🏢 得意先別ランキング
# ========================================================
if view_type == '🏢 得意先別ランキング':
    group_cols = [c for c in ["得意先名", "担当者"] if c in filtered.columns]
    agg = filtered.groupby(group_cols).agg(販売金額=("金額", "sum"), 数量=("数量", "sum"), 原価合計=("原価合計", "sum"), 粗利額=("粗利額", "sum"), 取扱品目数=("商品コード", "nunique")).reset_index() if not filtered.empty else pd.DataFrame()
    if not agg.empty:
        agg["粗利率"] = agg.apply(lambda r: f"{(r['粗利額'] / r['販売金額'] * 100):.1f}%" if r['販売金額'] > 0 else "-", axis=1)
        agg["構成比"] = (agg["販売金額"] / total_sales * 100).round(2).map(lambda x: f"{x:.2f}%") if total_sales > 0 else "0.00%"
        agg = agg.sort_values(by="販売金額", ascending=False)
        agg.insert(0, "順位", range(1, len(agg) + 1))
        if limit_val != "全件": agg = agg.head(int(limit_val))
        full_cust_df = agg[["順位", "得意先名", "担当者", "販売金額", "粗利額", "粗利率", "原価合計", "数量", "取扱品目数", "構成比"]].copy()
        show_df = agg[["順位", "得意先名", "販売金額", "粗利額", "粗利率"]].copy() if is_mobile_view else full_cust_df
    else:
        full_cust_df, show_df = pd.DataFrame(), pd.DataFrame()

    c_sum1, c_sum2 = st.columns([4, 1.2])
    with c_sum1:
        st.markdown(f"<div class='total-card' style='margin-bottom:0;'><b>抽出期間:</b> <span style='color:#1a73e8; font-weight:bold;'>{period_text}</span> &nbsp;|&nbsp; <b>売上合計:</b> <span style='color:#d93025; font-size:16px; font-weight:bold;'>{int(total_sales):,} 円</span></div>", unsafe_allow_html=True)
    with c_sum2:
        if not full_cust_df.empty:
            st.download_button(label="📥 ランキングをExcel保存", data=generate_excel_bytes(full_cust_df, "得意先別ランキング"), file_name=f"得意先別ランキング_{datetime.now().strftime('%Y%m%d')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

    st.write("")
    if not show_df.empty:
        default_cust = show_df.iloc[0]["得意先名"]
        if "selected_side_cust" not in st.session_state or st.session_state["selected_side_cust"] not in show_df["得意先名"].values:
            st.session_state["selected_side_cust"] = default_cust

        if is_mobile_view:
            st.caption("💡 表内の **「🔍 得意先名」** をタップすると、その得意先の購入商品一覧がポップアップ表示されます。")
            clicked_res = render_clickable_customer_table(show_df, selected_cust="")
            if clicked_res and isinstance(clicked_res, dict) and clicked_res.get("cust"):
                last_ts = st.session_state.get("_last_cust_click_ts", 0)
                cur_ts = clicked_res.get("ts", 0)
                if cur_ts != last_ts:
                    st.session_state["_last_cust_click_ts"] = cur_ts
                    clicked_cust = clicked_res["cust"]
                    show_customer_top10_dialog(clicked_cust, filtered[filtered["得意先名"] == clicked_cust], period_text)
        else:
            col_left_rank, col_right_detail = st.columns([1.15, 1.0], gap="medium")
            with col_left_rank:
                st.caption("💡 表内の **「🔍 得意先名」** をクリックすると、右側のパネルにその得意先の全購入商品が表示されます。")
                clicked_res = render_clickable_customer_table(show_df, selected_cust=st.session_state["selected_side_cust"])
                if clicked_res and isinstance(clicked_res, dict) and clicked_res.get("cust"):
                    last_ts = st.session_state.get("_last_cust_click_ts", 0)
                    cur_ts = clicked_res.get("ts", 0)
                    if cur_ts != last_ts:
                        st.session_state["_last_cust_click_ts"] = cur_ts
                        clicked_cust = clicked_res["cust"]
                        st.session_state["selected_side_cust"] = clicked_cust
                        if clicked_res.get("is_mobile"):
                            show_customer_top10_dialog(clicked_cust, filtered[filtered["得意先名"] == clicked_cust], period_text)
                        else:
                            st.rerun()

            with col_right_detail:
                target_side_cust = st.session_state["selected_side_cust"]
                side_prod_df = build_customer_all_products_df(filtered[filtered["得意先名"] == target_side_cust])
                st.markdown(f"<div class='side-panel-header'>🎯 {target_side_cust} の購入商品内訳（全件）</div>", unsafe_allow_html=True)
                if not side_prod_df.empty:
                    s_cnt, s_qty, s_amt = len(side_prod_df), int(side_prod_df["数量"].sum()), int(side_prod_df["販売金額"].sum())
                    sc1, sc2 = st.columns([2.6, 1.4])
                    with sc1:
                        st.markdown(f"<div class='total-card' style='margin-bottom:6px; padding:7px 12px; font-size:13px;'><b>該当品目数:</b> <span style='color:#1a73e8; font-weight:bold;'>{s_cnt:,} 品目</span> &nbsp;|&nbsp; <b>数量合計:</b> {s_qty:,} 個 &nbsp;|&nbsp; <b>売上合計:</b> <span style='color:#d93025; font-weight:bold;'>¥{s_amt:,}</span></div>", unsafe_allow_html=True)
                    with sc2:
                        st.download_button(label="📥 この得意先の内訳をExcel保存", data=generate_excel_bytes(side_prod_df, f"{target_side_cust}_商品内訳"), file_name=f"得意先別商品内訳_{target_side_cust}_{datetime.now().strftime('%Y%m%d')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, key=f"dl_side_cust_{target_side_cust}")
                    clicked_side_prod = render_clickable_product_table(side_prod_df, comp_key=f"side_prod_tbl_{target_side_cust}", max_h=525)
                    if clicked_side_prod and isinstance(clicked_side_prod, dict) and clicked_side_prod.get("sku"):
                        last_sp_ts = st.session_state.get("_last_side_prod_click_ts", 0)
                        cur_sp_ts = clicked_side_prod.get("ts", 0)
                        if cur_sp_ts != last_sp_ts:
                            st.session_state["_last_side_prod_click_ts"] = cur_sp_ts
                            c_sku = clicked_side_prod["sku"]
                            show_product_inventory_dialog(c_sku, code_to_name.get(c_sku, c_sku))
                else:
                    st.info("該当期間の商品売上データはありません。")
    else:
        st.info("該当期間のデータがありません。")
    st.stop()

# ========================================================
# ★ 📦 商品別売上ランキング
# ========================================================
elif view_type == '📦 商品別売上ランキング':
    group_cols = [c for c in ["得意先名", "担当者", "商品コード", "商品名"] if c in filtered.columns]
    if not filtered.empty:
        valid_date_df = filtered[filtered["日付"].str.strip() != ""]
        date_agg = valid_date_df.groupby(group_cols).agg(開始日=("日付", "min"), 終了日=("日付", "max")).reset_index()
        agg = filtered.groupby(group_cols).agg(販売金額=("金額", "sum"), 数量=("数量", "sum"), 単価原価=("単価原価", "first"), 原価合計=("原価合計", "sum"), 粗利額=("粗利額", "sum"), 単価_min=("単価", "min"), 単価_max=("単価", "max")).reset_index()
        agg = pd.merge(agg, date_agg, on=group_cols, how="left")
        agg["粗利率"] = agg.apply(lambda r: f"{(r['粗利額'] / r['販売金額'] * 100):.1f}%" if r['販売金額'] > 0 else "-", axis=1)
        agg["販売価格帯"] = agg.apply(lambda r: f"{int(r['単価_min']):,} 〜 {int(r['単価_max']):,}" if pd.notna(r['単価_min']) else "-", axis=1)
        agg["販売期間"] = agg.apply(lambda r: f"{r['開始日']} 〜 {r['終了日']}" if pd.notna(r['開始日']) else "-", axis=1)
        agg["構成比"] = (agg["販売金額"] / total_sales * 100).round(2).map(lambda x: f"{x:.2f}%") if total_sales > 0 else "0.00%"
        agg = agg.sort_values(by="販売金額", ascending=False)
        agg.insert(0, "番号", range(1, len(agg) + 1))
        if limit_val != "全件": agg = agg.head(int(limit_val))
        full_cols = ["番号", "得意先名", "担当者", "商品コード", "商品名", "数量", "販売価格帯", "単価原価", "販売金額", "原価合計", "粗利額", "粗利率", "構成比", "販売期間"]
        full_prod_df = agg[[c for c in full_cols if c in agg.columns]].copy()
        mob_cols = ["番号", "得意先名", "商品コード", "商品名", "数量", "販売金額", "粗利率"]
        show_df = agg[[c for c in mob_cols if c in agg.columns]].copy() if is_mobile_view else full_prod_df
    else:
        full_prod_df, show_df = pd.DataFrame(), pd.DataFrame()

    c_sum1, c_sum2 = st.columns([4, 1.2])
    with c_sum1:
        st.markdown(f"<div class='total-card' style='margin-bottom:0;'><b>抽出期間:</b> <span style='color:#1a73e8; font-weight:bold;'>{period_text}</span> &nbsp;|&nbsp; <b>販売金額合計:</b> <span style='color:#d93025; font-size:16px; font-weight:bold;'>{int(total_sales):,} 円</span></div>", unsafe_allow_html=True)
    with c_sum2:
        if not full_prod_df.empty:
            st.download_button(label="📥 Excelで保存", data=generate_excel_bytes(full_prod_df, "商品別売上ランキング"), file_name=f"商品別売上ランキング_{datetime.now().strftime('%Y%m%d')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

    st.write("")
    if not show_df.empty:
        st.caption("💡 表内の **「🔍 商品コード」** をクリック（またはタップ）すると、その商品の **「ロジザード現在庫・弥生確保数・発注検討ステータス詳細」** がポップアップ表示されます。")
        clicked_prod = render_clickable_product_table(show_df, comp_key="prod_rank_clickable_tbl_v1", max_h=600)
        if clicked_prod and isinstance(clicked_prod, dict) and clicked_prod.get("sku"):
            last_p_ts = st.session_state.get("_last_prod_click_ts", 0)
            cur_p_ts = clicked_prod.get("ts", 0)
            if cur_p_ts != last_p_ts:
                st.session_state["_last_prod_click_ts"] = cur_p_ts
                c_sku = clicked_prod["sku"]
                c_pname = code_to_name.get(c_sku, c_sku)
                show_product_inventory_dialog(c_sku, c_pname)
    else:
        st.info("該当期間のデータがありません。")
    st.stop()

# ========================================================
# ★ 👤 担当者別 売上＆目標達成分析
# ========================================================
elif view_type == '👤 担当者別 売上＆目標達成分析':
    st.markdown(f"#### 👤 担当者別 売上＆目標達成分析（対象期間: {period_text}）")
    if not filtered.empty:
        staff_agg = filtered.groupby("担当者").agg(
            販売金額=("金額", "sum"), 数量=("数量", "sum"), 原価合計=("原価合計", "sum"),
            粗利額=("粗利額", "sum"), 担当得意先数=("得意先名", "nunique"), 取扱品目数=("商品コード", "nunique")
        ).reset_index()
        staff_agg = staff_agg[staff_agg["担当者"].astype(str).str.strip() != ""]

        if not ly_filtered.empty:
            ly_st = ly_filtered.groupby("担当者").agg(前年売上額=("金額", "sum"), 前年粗利額=("粗利額", "sum")).reset_index()
            staff_agg = pd.merge(staff_agg, ly_st, on="担当者", how="left").fillna(0)
        else:
            staff_agg["前年売上額"], staff_agg["前年粗利額"] = 0, 0

        def calc_period_target(st_name):
            norm_s = normalize_target_staff_name(st_name)
            t_info = staff_target_master.get(norm_s, {})
            if not t_info: return 0
            m_map = t_info.get("monthly", {})
            return int(t_info.get("annual", 0)) if not active_months_list else int(sum(m_map.get(m, 0) for m in active_months_list))

        def calc_annual_target(st_name):
            return int(staff_target_master.get(normalize_target_staff_name(st_name), {}).get("annual", 0))

        staff_agg["期間目標額"] = staff_agg["担当者"].map(calc_period_target)
        staff_agg["目標差額"] = staff_agg["販売金額"] - staff_agg["期間目標額"]
        staff_agg["目標達成率"] = staff_agg.apply(lambda r: f"{(r['販売金額'] / r['期間目標額'] * 100):.1f}%" if r["期間目標額"] > 0 else "-", axis=1)
        staff_agg["通期目標額"] = staff_agg["担当者"].map(calc_annual_target)
        staff_agg["通期進捗率"] = staff_agg.apply(lambda r: f"{(r['販売金額'] / r['通期目標額'] * 100):.1f}%" if r["通期目標額"] > 0 else "-", axis=1)

        staff_agg["前年売上額"] = staff_agg["前年売上額"].astype(int)
        staff_agg["増減額"] = staff_agg["販売金額"] - staff_agg["前年売上額"]
        staff_agg["前年比"] = staff_agg.apply(lambda r: f"{(r['販売金額'] / r['前年売上額'] * 100):.1f}%" if r["前年売上額"] > 0 else "-", axis=1)
        staff_agg["粗利率"] = staff_agg.apply(lambda r: f"{(r['粗利額'] / r['販売金額'] * 100):.1f}%" if r['販売金額'] > 0 else "-", axis=1)
        staff_agg["構成比"] = (staff_agg["販売金額"] / total_sales * 100).round(1).map(lambda x: f"{x:.1f}%") if total_sales > 0 else "0.0%"
        staff_agg = staff_agg.sort_values(by="販売金額", ascending=False)
        staff_agg.insert(0, "順位", range(1, len(staff_agg) + 1))

        cols_st = ["順位", "担当者", "販売金額", "期間目標額", "目標達成率", "目標差額", "通期目標額", "通期進捗率", "前年売上額", "前年比", "増減額", "粗利額", "粗利率", "数量", "担当得意先数", "取扱品目数", "構成比"]
        show_st_df = staff_agg[[c for c in cols_st if c in staff_agg.columns]].copy()

        tot_prof, tot_tgt, tot_ly = int(staff_agg["粗利額"].sum()), int(staff_agg["期間目標額"].sum()), int(staff_agg["前年売上額"].sum())
        tot_ach = f"{(total_sales / tot_tgt * 100):.1f}%" if tot_tgt > 0 else "-"
        yoy_tot_str = f"{(total_sales / tot_ly * 100):.1f}%" if tot_ly > 0 else "-"

        c_sum1, c_sum2 = st.columns([4, 1.2])
        with c_sum1:
            st.markdown(f"<div class='total-card' style='margin-bottom:0;'><b>期間売上合計:</b> <span style='color:#d93025; font-size:16px; font-weight:bold;'>¥{total_sales:,}</span> &nbsp;|&nbsp; <b>期間目標合計:</b> ¥{tot_tgt:,} (達成率: <b style='color:#1a73e8;'>{tot_ach}</b>) &nbsp;|&nbsp; <b>前年同期合計:</b> ¥{tot_ly:,} (前年比: <b>{yoy_tot_str}</b>) &nbsp;|&nbsp; <b>粗利合計:</b> <span style='color:#188038; font-weight:bold;'>¥{tot_prof:,}</span></div>", unsafe_allow_html=True)
        with c_sum2:
            st.download_button(label="📥 Excelで保存", data=generate_excel_bytes(show_st_df, "担当者別分析"), file_name=f"担当者別売上分析_{datetime.now().strftime('%Y%m%d')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        st.write("")
        render_universal_table(show_st_df)
    else:
        st.info("該当期間のデータがありません。")
    st.stop()

# ========================================================
# ★ ✨ 新商品売上分析（直近1年）
# ========================================================
elif view_type == '✨ 新商品売上分析（直近1年）':
    st.markdown(f"#### ✨ 新商品売上分析（{period_text}）")
    st.caption("※全データ履歴の中で「初回売上日」が直近1年以内の新商品のみを自動抽出しています。")
    if not filtered.empty:
        c_sort1, c_sort2, _ = st.columns([1.8, 1.8, 2.0])
        with c_sort1: np_sort_col = st.radio("📊 並び替え項目", ["販売金額", "数量", "導入社数"], horizontal=True, key="np_sort_col_radio")
        with c_sort2: np_sort_order = st.radio("🔃 並び順", ["⬇️ 降順（高い順）", "⬆️ 昇順（低い順）"], horizontal=True, key="np_sort_order_radio")
        is_asc = np_sort_order.startswith("⬆️")

        np_agg = filtered.groupby(["商品コード", "商品名"]).agg(
            販売金額=("金額", "sum"), 数量=("数量", "sum"), 単価原価=("単価原価", "first"),
            原価合計=("原価合計", "sum"), 粗利額=("粗利額", "sum"), 導入社数=("得意先名", "nunique"), 最終売上日=("日付", "max")
        ).reset_index()
        np_agg["初回売上日"] = np_agg["商品コード"].map(lambda c: prod_first_seen_dt.get(c).strftime('%Y/%m/%d') if pd.notna(prod_first_seen_dt.get(c)) else "-")
        np_agg["粗利率"] = np_agg.apply(lambda r: f"{(r['粗利額'] / r['販売金額'] * 100):.1f}%" if r['販売金額'] > 0 else "-", axis=1)
        np_agg = np_agg.sort_values(by=[np_sort_col, "販売金額"], ascending=[is_asc, False]).reset_index(drop=True)
        np_agg.insert(0, "順位", range(1, len(np_agg) + 1))
        if limit_val != "全件": np_agg = np_agg.head(int(limit_val))
        cols_np = ["順位", "商品コード", "商品名", "初回売上日", "最終売上日", "導入社数", "数量", "単価原価", "販売金額", "粗利額", "粗利率"]
        show_np_df = np_agg[[c for c in cols_np if c in np_agg.columns]].copy()

        c_sum1, c_sum2 = st.columns([4, 1.2])
        with c_sum1:
            st.markdown(f"<div class='total-card' style='margin-bottom:0;'><b>新商品 該当品目数:</b> <span style='color:#1a73e8; font-weight:bold;'>{len(show_np_df):,} 品目</span> &nbsp;|&nbsp; <b>新商品 売上合計:</b> <span style='color:#d93025; font-size:16px; font-weight:bold;'>¥{int(show_np_df['販売金額'].sum()):,}</span> &nbsp;|&nbsp; <b>現在の並び順:</b> <span style='color:#188038; font-weight:bold;'>{np_sort_col}（{'昇順' if is_asc else '降順'}）</span></div>", unsafe_allow_html=True)
        with c_sum2:
            st.download_button(label="📥 Excelで保存", data=generate_excel_bytes(show_np_df, "新商品売上分析"), file_name=f"新商品売上分析_{datetime.now().strftime('%Y%m%d')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        st.write("")
        render_universal_table(show_np_df)
    else:
        st.info("直近1年間に初登場した新商品の売上データはありません。")
    st.stop()

# ========================================================
# ★ 📊 営業会議用：上位社 売上比較シート
# ========================================================
elif view_type == '📊 営業会議用：上位社 売上比較シート':
    if "tgt_ym_meeting_v3" not in st.session_state:
        st.session_state["tgt_ym_meeting_v3"] = ym_start if ym_start in desc_valid_ym else desc_valid_ym[0]
    if "m_ed_ver" not in st.session_state:
        st.session_state["m_ed_ver"] = 0

    col_m1, col_m2, col_m3 = st.columns([1.5, 1.5, 1.5])
    with col_m1: target_meeting_ym = st.selectbox("対象月度（期間ショートカット連動）", desc_valid_ym, key="tgt_ym_meeting_v3")
    with col_m2: compare_cust_count = st.selectbox("比較上位社数", [4, 6, 8, 10, 20, 30], index=0, key="compare_cust_count_meeting_v3")
    with col_m3:
        staff_options_meeting = ["（すべて）"] + sorted([str(x).strip() for x in df["担当者"].dropna().unique() if str(x).strip() not in ["", "nan", "None"]])
        sel_staff_m = st.selectbox("担当者指定", staff_options_meeting, key="meeting_staff_sel_v3")

    tgt_y, tgt_m = int(target_meeting_ym[:4]), int(target_meeting_ym[5:7])
    ly_ym = f"{tgt_y - 1}年{tgt_m:02d}月"
    next_y = tgt_y + 1 if tgt_m == 12 else tgt_y
    next_m = 1 if tgt_m == 12 else tgt_m + 1
    next_ym = f"{next_y}年{next_m:02d}月"
    f_year = tgt_y if tgt_m >= 7 else tgt_y - 1
    prev_fy_ym = f"{next_y - 1}年{next_m:02d}月"

    df_scope = df if sel_staff_m == "（すべて）" else df[df["担当者"] == sel_staff_m]
    cur_m_df = df_scope[df_scope["年月"] == target_meeting_ym] if not df_scope.empty else pd.DataFrame()
    ly_m_df = df_scope[df_scope["年月"] == ly_ym] if not df_scope.empty else pd.DataFrame()
    next_m_df = df_scope[df_scope["年月"] == next_ym] if not df_scope.empty else pd.DataFrame()
    prev_fy_df = df_scope[df_scope["年月"] == prev_fy_ym] if not df_scope.empty else pd.DataFrame()

    cust_id_lookup = {}
    if not customer_master_df.empty:
        c_name_col = "加盟店名" if "加盟店名" in customer_master_df.columns else customer_master_df.columns[0]
        c_id_col = "加盟店ID" if "加盟店ID" in customer_master_df.columns else ("顧客ID" if "顧客ID" in customer_master_df.columns else None)
        if c_id_col:
            for _, r_c in customer_master_df.iterrows():
                cn = clean_cust_name(str(r_c[c_name_col]))
                cid = str(r_c[c_id_col]).replace(".0", "").strip()
                if cn and cid != "nan": cust_id_lookup[cn] = cid

    cur_cust_sales = cur_m_df.groupby("得意先名")["金額"].sum() if not cur_m_df.empty else pd.Series()
    top_custs = cur_cust_sales.nlargest(compare_cust_count).index.tolist() if not cur_cust_sales.empty else []

    all_saved_meeting_dict = load_saved_meeting_sheet_data()
    month_saved_dict = all_saved_meeting_dict.get(target_meeting_ym, {})

    cards_data = [build_meeting_comparison_card_data(cust, cust_id_lookup.get(cust, ""), cur_m_df, ly_m_df, next_m_df, prev_fy_df, tgt_y, tgt_m, next_y, next_m, f_year, month_saved_dict) for cust in top_custs]

    m_ver = st.session_state["m_ed_ver"]
    with st.expander("📝 各社の「主な営業活動」「今後の予定」「売上進捗手動入力」を編集・保存する（帳票＆Excelに完全反映）", expanded=False):
        st.caption("※品番・単価・数量・金額は何度でも書き換え・削除（0または空欄化）が可能です。入力後は一番下の「💾 この月度の入力内容を保存する」を押してください。")
        current_edit_results = {}
        for c_item in cards_data:
            cn = c_item["cust_name"]
            col_title, col_rst = st.columns([3.5, 1.5])
            with col_title: st.markdown(f"##### 🏢 **{cn}**")
            with col_rst:
                if st.button(f"🔄 {cn} の進捗表を自動集計に戻す", key=f"rst_btn_{target_meeting_ym}_{cn}_{m_ver}"):
                    if target_meeting_ym in all_saved_meeting_dict and cn in all_saved_meeting_dict[target_meeting_ym]:
                        all_saved_meeting_dict[target_meeting_ym][cn].pop("custom_next_rows", None)
                        all_saved_meeting_dict[target_meeting_ym][cn].pop("custom_next_other", None)
                        save_meeting_sheet_data(all_saved_meeting_dict)
                    st.session_state["m_ed_ver"] += 1
                    st.rerun()

            ca1, ca2 = st.columns(2)
            with ca1: st.text_area("この1ヶ月の主な営業活動", value=c_item.get("activity_text", ""), key=f"inp_act_{target_meeting_ym}_{cn}_{m_ver}", height=70)
            with ca2: st.text_area("【今後の予定】", value=c_item.get("plan_text", ""), key=f"inp_pln_{target_meeting_ym}_{cn}_{m_ver}", height=70)

            st.caption("▼ 翌月売上進捗の上書き・手動修正（行を削除したい場合は品番を消すか数値を0にして保存してください）")
            edit_items_df = pd.DataFrame(c_item.get("next_rows", []))
            for c_col in ["sku", "price", "qty", "amt"]:
                if c_col not in edit_items_df.columns: edit_items_df[c_col] = "" if c_col == "sku" else 0
            edit_items_df["sku"] = edit_items_df["sku"].fillna("").astype(str)
            for n_col in ["price", "qty", "amt"]: edit_items_df[n_col] = edit_items_df[n_col].apply(clean_num).astype(int)

            edit_res = st.data_editor(
                edit_items_df[["sku", "price", "qty", "amt"]],
                column_config={
                    "sku": st.column_config.TextColumn("品番", width="medium"),
                    "price": st.column_config.NumberColumn("単価", min_value=0, step=1, format="¥%d"),
                    "qty": st.column_config.NumberColumn("数量", min_value=0, step=1, format="%d"),
                    "amt": st.column_config.NumberColumn("金額", min_value=0, step=1, format="¥%d")
                },
                num_rows="dynamic", use_container_width=True, hide_index=True, key=f"next_edit_tbl_{target_meeting_ym}_{cn}_{m_ver}"
            )
            current_edit_results[cn] = edit_res
            st.number_input("「その他」金額", value=int(c_item.get("next_other", 0)), step=1000, key=f"inp_oth_{target_meeting_ym}_{cn}_{m_ver}")
            st.markdown("---")

        if st.button("💾 この月度の入力内容を保存する", type="primary", use_container_width=True):
            if target_meeting_ym not in all_saved_meeting_dict: all_saved_meeting_dict[target_meeting_ym] = {}
            for c_item in cards_data:
                cn = c_item["cust_name"]
                t_df = current_edit_results.get(cn, pd.DataFrame())
                o_key, act_k, pln_k = f"inp_oth_{target_meeting_ym}_{cn}_{m_ver}", f"inp_act_{target_meeting_ym}_{cn}_{m_ver}", f"inp_pln_{target_meeting_ym}_{cn}_{m_ver}"
                rows_to_save = []
                if isinstance(t_df, pd.DataFrame) and not t_df.empty:
                    for _, r in t_df.iterrows():
                        sku_v = str(r.get("sku", "") or "").strip()
                        if sku_v in ["None", "nan"]: sku_v = ""
                        p_val, q_val, a_val = int(clean_num(r.get("price", 0))), int(clean_num(r.get("qty", 0))), int(clean_num(r.get("amt", 0)))
                        if not sku_v and p_val == 0 and q_val == 0 and a_val == 0: continue
                        if a_val == 0 and p_val > 0 and q_val > 0: a_val = p_val * q_val
                        rows_to_save.append({"sku": sku_v, "price": p_val, "qty": q_val, "amt": a_val})
                while len(rows_to_save) < 5: rows_to_save.append({"sku": "", "price": 0, "qty": 0, "amt": 0})
                all_saved_meeting_dict[target_meeting_ym][cn] = {
                    "activity": st.session_state.get(act_k, ""),
                    "plan": st.session_state.get(pln_k, ""),
                    "custom_next_rows": rows_to_save[:5],
                    "custom_next_other": int(clean_num(st.session_state.get(o_key, 0)))
                }
            save_meeting_sheet_data(all_saved_meeting_dict)
            st.session_state["m_ed_ver"] += 1
            st.success("✅ 入力内容を保存しました！カード表示とExcel出力に反映されました。")
            time.sleep(0.4)
            st.rerun()

    st.download_button(label=f"📥 上位社 売上比較シート（{target_meeting_ym}）をExcel保存", data=generate_meeting_comparison_excel(cards_data, target_meeting_ym), file_name=f"営業会議_上位社売上比較シート_{target_meeting_ym}_{datetime.now().strftime('%Y%m%d')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_btn_meeting_comp_v3")
    st.markdown(f"#### 📊 {tgt_y}年{tgt_m}月度 上位{len(top_custs)}社 売上比較シート（担当: {sel_staff_m}）")
    if not cards_data:
        st.info(f"{target_meeting_ym}（担当: {sel_staff_m}）の売上データがありません。")
    else:
        for idx_c in range(0, len(cards_data), 2):
            cols_card = st.columns(2)
            for sub_i in range(2):
                if idx_c + sub_i < len(cards_data):
                    with cols_card[sub_i]: st.markdown(render_meeting_card_html(cards_data[idx_c + sub_i]), unsafe_allow_html=True)
    st.stop()

# ========================================================
# ★ 📑 営業会議用：上位客先別売上実績
# ========================================================
elif view_type == '📑 営業会議用：上位客先別売上実績':
    if "tgt_ym_topcust_v3" not in st.session_state:
        st.session_state["tgt_ym_topcust_v3"] = ym_start if ym_start in desc_valid_ym else desc_valid_ym[0]

    col_m1, col_m2, col_m3 = st.columns([1.5, 1.5, 1.5])
    with col_m1: target_meeting_ym = st.selectbox("対象月度（期間ショートカット連動）", desc_valid_ym, key="tgt_ym_topcust_v3")
    with col_m2: meeting_cust_count = st.selectbox("表示上位数", [10, 20, 30, 50, 100], index=0, key="meeting_cust_count_v3")
    with col_m3:
        staff_options_topcust = ["（すべて）"] + sorted([str(x).strip() for x in df["担当者"].dropna().unique() if str(x).strip() not in ["", "nan", "None"]])
        sel_staff_tc = st.selectbox("担当者指定", staff_options_topcust, key="topcust_staff_sel_v3")

    tgt_y, tgt_m = int(target_meeting_ym[:4]), int(target_meeting_ym[5:7])
    ly_ym = f"{tgt_y - 1}年{tgt_m:02d}月"
    f_year = tgt_y if tgt_m >= 7 else tgt_y - 1
    term_start_ym, ly_term_start_ym, ly_term_end_ym = f"{f_year}年07月", f"{f_year - 1}年07月", f"{tgt_y - 1}年{tgt_m:02d}月"

    df_scope = df if sel_staff_tc == "（すべて）" else df[df["担当者"] == sel_staff_tc]
    cur_m_df = df_scope[df_scope["年月"] == target_meeting_ym] if not df_scope.empty else pd.DataFrame()
    ly_m_df = df_scope[df_scope["年月"] == ly_ym] if not df_scope.empty else pd.DataFrame()
    cur_ytd_df = df_scope[(df_scope["年月"] >= term_start_ym) & (df_scope["年月"] <= target_meeting_ym)] if not df_scope.empty else pd.DataFrame()
    ly_ytd_df = df_scope[(df_scope["年月"] >= ly_term_start_ym) & (df_scope["年月"] <= ly_term_end_ym)] if not df_scope.empty else pd.DataFrame()

    cur_cust_sales = cur_m_df.groupby("得意先名")["金額"].sum() if not cur_m_df.empty else pd.Series()
    ly_cust_sales = ly_m_df.groupby("得意先名")["金額"].sum() if not ly_m_df.empty else pd.Series()
    cur_ytd_sales = cur_ytd_df.groupby("得意先名")["金額"].sum() if not cur_ytd_df.empty else pd.Series()
    ly_ytd_sales = ly_ytd_df.groupby("得意先名")["金額"].sum() if not ly_ytd_df.empty else pd.Series()

    target_cust_list = cur_cust_sales.nlargest(meeting_cust_count).index.tolist() if not cur_cust_sales.empty else []
    m_rows, tot_cur_m, tot_ly_m, tot_cur_ytd, tot_ly_ytd = [], 0, 0, 0, 0

    for rank, cust in enumerate(target_cust_list, 1):
        s_cur, s_ly, ytd_cur, ytd_ly = int(cur_cust_sales.get(cust, 0)), int(ly_cust_sales.get(cust, 0)), int(cur_ytd_sales.get(cust, 0)), int(ly_ytd_sales.get(cust, 0))
        tot_cur_m += s_cur; tot_ly_m += s_ly; tot_cur_ytd += ytd_cur; tot_ly_ytd += ytd_ly
        m_rows.append({
            "順位": str(rank), "客先名": cust, "売上実績": s_cur, "前年実績": s_ly,
            "前年比": f"{(s_cur / s_ly * 100):.0f}%" if s_ly != 0 else "-", "増減": s_cur - s_ly,
            "今期累計実績": ytd_cur, "前期累計実績": ytd_ly,
            "累計前年比": f"{(ytd_cur / ytd_ly * 100):.0f}%" if ytd_ly != 0 else "-"
        })

    res_m_df = pd.DataFrame(m_rows)
    tot_row = pd.DataFrame([{
        "順位": "", "客先名": "合計", "売上実績": tot_cur_m, "前年実績": tot_ly_m,
        "前年比": f"{(tot_cur_m / tot_ly_m * 100):.0f}%" if tot_ly_m != 0 else "-", "増減": tot_cur_m - tot_ly_m,
        "今期累計実績": tot_cur_ytd, "前期累計実績": tot_ly_ytd,
        "累計前年比": f"{(tot_cur_ytd / tot_ly_ytd * 100):.0f}%" if tot_ly_ytd != 0 else "-"
    }])
    res_m_df_tot = pd.concat([res_m_df, tot_row], ignore_index=True)
    title_sheet = f"{tgt_m}月度上位{len(target_cust_list)}社客先別売上実績"

    st.download_button(label=f"📥 上位客先別売上実績（{target_meeting_ym}）をExcel保存", data=generate_meeting_top_cust_excel(res_m_df_tot, title_sheet), file_name=f"営業会議_{title_sheet}_{target_meeting_ym}_{datetime.now().strftime('%Y%m%d')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_btn_top_cust_v3")
    st.markdown(f"### {title_sheet}（担当: {sel_staff_tc}）")
    render_universal_table(res_m_df_tot)
    st.stop()

# ========================================================
# ★ 📈 商品別 前年実績比較
# ========================================================
elif view_type == '📈 商品別 前年実績比較':
    st.markdown(f"#### 📈 商品別 前年実績比較（対象期間: {period_text}）")
    cur_p = filtered.groupby(["商品コード", "商品名"]).agg(当年数量=("数量", "sum"), 当年売上=("金額", "sum"), 当年粗利=("粗利額", "sum")).reset_index() if not filtered.empty else pd.DataFrame(columns=["商品コード", "商品名", "当年数量", "当年売上", "当年粗利"])
    ly_p = ly_filtered.groupby(["商品コード"]).agg(前年数量=("数量", "sum"), 前年売上=("金額", "sum"), 前年粗利=("粗利額", "sum")).reset_index() if not ly_filtered.empty else pd.DataFrame(columns=["商品コード", "前年数量", "前年売上", "前年粗利"])
    comp_p = pd.merge(cur_p, ly_p, on="商品コード", how="outer").fillna(0)
    comp_p["商品名"] = comp_p.apply(lambda r: r["商品名"] if r["商品名"] not in [0, "0", ""] else code_to_name.get(r["商品コード"], r["商品コード"]), axis=1)
    for col_int in ["当年数量", "前年数量", "当年売上", "前年売上", "当年粗利", "前年粗利"]: comp_p[col_int] = comp_p[col_int].astype(int)
    comp_p["増減額"] = comp_p["当年売上"] - comp_p["前年売上"]
    comp_p["前年比"] = comp_p.apply(lambda r: f"{(r['当年売上'] / r['前年売上'] * 100):.1f}%" if r["前年売上"] > 0 else "-", axis=1)
    comp_p = comp_p.sort_values(by="当年売上", ascending=False)
    comp_p.insert(0, "順位", range(1, len(comp_p) + 1))
    if limit_val != "全件": comp_p = comp_p.head(int(limit_val))
    show_cp = comp_p[["順位", "商品コード", "商品名", "当年数量", "前年数量", "当年売上", "前年売上", "前年比", "増減額", "当年粗利"]].copy()
    render_universal_table(show_cp)
    st.stop()

# ========================================================
# ★ 📊 得意先別 前年実績比較
# ========================================================
elif view_type == '📊 得意先別 前年実績比較':
    st.markdown(f"#### 📊 得意先別 前年実績比較（対象期間: {period_text}）")
    cur_c = filtered.groupby(["得意先名"]).agg(担当者=("担当者", "last"), 当年売上=("金額", "sum"), 当年粗利=("粗利額", "sum"), 当年数量=("数量", "sum")).reset_index() if not filtered.empty else pd.DataFrame(columns=["得意先名", "担当者", "当年売上", "当年粗利", "当年数量"])
    ly_c = ly_filtered.groupby(["得意先名"]).agg(前年売上=("金額", "sum"), 前年粗利=("粗利額", "sum"), 前年数量=("数量", "sum")).reset_index() if not ly_filtered.empty else pd.DataFrame(columns=["得意先名", "前年売上", "前年粗利", "前年数量"])
    comp_c = pd.merge(cur_c, ly_c, on="得意先名", how="outer").fillna(0)
    comp_c["担当者"] = comp_c["担当者"].replace(0, "")
    for col_int in ["当年売上", "前年売上", "当年粗利", "前年粗利", "当年数量", "前年数量"]: comp_c[col_int] = comp_c[col_int].astype(int)
    comp_c["増減額"] = comp_c["当年売上"] - comp_c["前年売上"]
    comp_c["前年比"] = comp_c.apply(lambda r: f"{(r['当年売上'] / r['前年売上'] * 100):.1f}%" if r["前年売上"] > 0 else "-", axis=1)
    comp_c = comp_c.sort_values(by="当年売上", ascending=False)
    comp_c.insert(0, "順位", range(1, len(comp_c) + 1))
    if limit_val != "全件": comp_c = comp_c.head(int(limit_val))
    show_cc = comp_c[["順位", "得意先名", "担当者", "当年売上", "前年売上", "前年比", "増減額", "当年粗利", "当年数量"]].copy()
    render_universal_table(show_cc)
    st.stop()
