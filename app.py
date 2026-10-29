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
import plotly.express as px
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.drawing.image import Image as OpenpyxlImage

warnings.filterwarnings("ignore")

st.set_page_config(page_title="売上管理＆営業分析表", layout="wide")

st.markdown("""
<style>
    .main-header {
        background-color: #2a75d3;
        color: white;
        padding: 8px 15px;
        font-size: 19px;
        font-weight: bold;
        border-radius: 4px;
        margin-bottom: 10px;
    }
    .total-card {
        background-color: #f0f4f8;
        padding: 10px 14px;
        border-radius: 6px;
        border-left: 5px solid #2a75d3;
        margin-bottom: 8px;
    }
    .est-card {
        border: 1px solid #cbd5e1;
        border-radius: 6px;
        background-color: #ffffff;
        margin-bottom: 16px;
        overflow: hidden;
    }
    .est-header {
        background-color: #f8fafc;
        padding: 10px 14px;
        border-bottom: 1px solid #cbd5e1;
        display: flex;
        justify-content: space-between;
        align-items: center;
    }
    .badge-green {
        background-color: #16a34a;
        color: white;
        padding: 2px 6px;
        border-radius: 3px;
        font-size: 11px;
        font-weight: bold;
    }
    .modal-box-header {
        background: linear-gradient(135deg, #1e3a8a, #2563eb);
        color: white;
        padding: 12px 18px;
        border-radius: 6px 6px 0 0;
        font-size: 16px;
        font-weight: bold;
        margin-bottom: 10px;
    }
    .modal-summary-box {
        background-color: #fefce8;
        border: 1px solid #fde047;
        border-left: 6px solid #eab308;
        padding: 10px 16px;
        font-size: 13px;
        margin-bottom: 12px;
        border-radius: 4px;
    }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-header">売上管理＆営業分析表</div>', unsafe_allow_html=True)

# クラウドやGitHub上では、app.pyと同じ場所にある「その他データ」または同じ階層のExcelを読むようにする
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OTHER_DATA_DIR = BASE_DIR

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CACHE_FILE = os.path.join(BASE_DIR, "sales_cache.parquet")
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SAVED_QUOTES_FILE = os.path.join(BASE_DIR, "saved_estimates.json")

STAFF_MAP = {
    "1": "小泉", "2": "河端", "3": "安部", "4": "原",
    "5": "西山", "6": "萩原", "7": "田中", "8": "水原",
    "9": "川村", "10": "穴山", "11": "鈴木", "12": "松平"
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

EXCLUDE_RANK_PATTERN = re.compile(r"下代変更前|オ.*プン|ｵ.*ﾌﾟﾝ|open|サンプル|ｻﾝﾌﾟﾙ|99,?999", re.IGNORECASE)

def is_invalid_rank_name(name_str):
    if not name_str: return True
    s = str(name_str).strip()
    if s in ["", "nan", "None", "#N/A"]: return True
    return bool(EXCLUDE_RANK_PATTERN.search(s))

def clean_cust_name(val):
    if not isinstance(val, str) or not val.strip(): return ""
    raw = str(val).strip()
    norm = re.sub(r'[\s ]+', '', raw)
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
    except: return 0.0

def convert_staff_name(val):
    if pd.isna(val): return ""
    v = str(val).strip().replace(".0", "")
    return STAFF_MAP.get(v, v)

def clean_product_name(val):
    if not isinstance(val, str) or not val.strip(): return ""
    if "】" in val: val = val.split("】", 1)[-1]
    return re.sub(r'[\(（]\s*[\d０-９]{7,16}\s*[\)）]', '', val).strip()

def find_file_multidir(filename_pattern):
    # クラウド（GitHub）上のファイルを探すため、BASE_DIRのみを対象にする
    search_dirs = [BASE_DIR]
    for d in search_dirs:
        if os.path.exists(d):
            fl = glob.glob(os.path.join(d, filename_pattern))
            if fl: return fl[0], os.path.getmtime(fl[0])
    return None, 0

def get_cost_master_file_info():
    return find_file_multidir("*原価*.xls*")

@st.cache_data(show_spinner=False)
def load_cost_master(file_path, file_mtime):
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
                    sk = str(r[sku_col]).strip()
                    co = clean_num(r[cost_col])
                    if sk: cost_dict[sk] = co
        except Exception: pass
    return cost_dict

def get_member_price_file_info():
    return find_file_multidir("*加盟店価格*.xls*")

@st.cache_data(show_spinner=False)
def load_member_price_master_v2(file_path, file_mtime):
    cust_sku_map = {}
    default_min_map = {}
    if file_path and os.path.exists(file_path):
        try:
            mdf = pd.read_excel(file_path)
            for _, r in mdf.iterrows():
                sku = str(r.iloc[3]).strip() if len(r) > 3 else ""
                if (not sku or sku == "nan") and len(r) > 4:
                    sku = str(r.iloc[4]).strip()
                if not sku or sku == "nan": continue

                j_price = int(clean_num(r.iloc[9])) if len(r) > 9 else 0
                n_min_price = int(clean_num(r.iloc[13])) if len(r) > 13 else 0
                m_comp = str(r.iloc[12]).strip() if len(r) > 12 else ""

                if 50 <= n_min_price < 90000:
                    default_min_map[sku] = n_min_price

                if m_comp and m_comp not in ["nan", "None", "#N/A", ""]:
                    norm_comp = re.sub(r'[\s ]+', '', m_comp)
                    if 50 <= j_price < 90000:
                        cust_sku_map[(m_comp, sku)] = j_price
                        cust_sku_map[(norm_comp, sku)] = j_price
        except Exception: pass
    return cust_sku_map, default_min_map

def get_jan_file_info():
    return find_file_multidir("*JAN*.xls*")

@st.cache_data(show_spinner=False)
def load_jan_master_v2(file_path, file_mtime):
    jan_dict = {}
    full_title_dict = {}
    std_price_dict = {}
    if file_path and os.path.exists(file_path):
        try:
            jdf = pd.read_excel(file_path)
            c_sku = [c for c in jdf.columns if any(k in str(c) for k in ['商品番号', '商品コード', 'SKU', '品番'])]
            c_jan = [c for c in jdf.columns if any(k in str(c) for k in ['カタログID', 'JAN', 'JANコード'])]
            c_name = [c for c in jdf.columns if any(k in str(c) for k in ['商品名', '品名'])]
            c_std = [c for c in jdf.columns if any(k in str(c) for k in ['上代', '標準', '定価', '税抜上代'])]
            s_col = c_sku[0] if c_sku else jdf.columns[0]
            j_col = c_jan[0] if c_jan else (jdf.columns[1] if len(jdf.columns) > 1 else None)
            n_col = c_name[0] if c_name else None
            p_col = c_std[0] if c_std else None

            for _, r in jdf.iterrows():
                sk = str(r[s_col]).strip()
                if not sk or sk == "nan": continue
                if j_col and pd.notna(r[j_col]):
                    try: jan_dict[sk] = str(int(float(r[j_col])))
                    except: jan_dict[sk] = str(r[j_col]).replace(".0", "").strip()
                if n_col and pd.notna(r[n_col]):
                    full_title_dict[sk] = str(r[n_col]).strip()
                if p_col and pd.notna(r[p_col]):
                    std_price_dict[sk] = int(clean_num(r[p_col]))
        except Exception: pass
    return jan_dict, full_title_dict, std_price_dict

def get_customer_file_info():
    return find_file_multidir("*注文者会社名*.xls*")

@st.cache_data(show_spinner=False)
def load_customer_master(file_path, file_mtime):
    if file_path and os.path.exists(file_path):
        try: return pd.read_excel(file_path)
        except Exception: pass
    return pd.DataFrame()

def get_hanko_bytes():
    # クラウド環境に合わせてBASE_SHAREを除外し、BASE_DIRとOTHER_DATA_DIRのみを対象にする
    search_dirs = [BASE_DIR]
    patterns = ["*hanko*.png", "*05171642*.png", "*印鑑*.png", "input_file_5.png"]
    for d in search_dirs:
        if os.path.exists(d):
            for p in patterns:
                fl = glob.glob(os.path.join(d, p))
                if fl:
                    try:
                        with open(fl[0], "rb") as f: return f.read(), fl[0]
                    except: pass
    return None, None

def load_saved_quotes():
    if os.path.exists(SAVED_QUOTES_FILE):
        try:
            with open(SAVED_QUOTES_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception: return []
    return []

def save_quotes_list(q_list):
    try:
        with open(SAVED_QUOTES_FILE, "w", encoding="utf-8") as f:
            json.dump(q_list, f, ensure_ascii=False, indent=2)
    except Exception: pass

def create_smaregi_order_csv_bytes(est_data, raw_title_map, std_price_map):
    rows = []
    pay_map = {"銀行振込": 3, "代金引換": 2, "現行通り": 3, "掛け払い（月末締め翌月末払い）": 3}
    pay_id = pay_map.get(est_data.get("payment", ""), 3)
    now_dt_str = datetime.now().strftime("%Y/%m/%d %H:%M")
    ship_dt_str = datetime.now().strftime("%Y/%m/%d")
    deliv_dt_str = (datetime.now() + timedelta(days=1)).strftime("%Y/%m/%d")

    for item in est_data.get("items", []):
        sku = item.get("商品コード", "")
        qty = int(item.get("数量", 1))
        unit_p = int(item.get("見積単価", 0))
        line_amt = unit_p * qty
        unit_tax = int(round(unit_p * 0.10))
        unit_tax_in = unit_p + unit_tax
        jan = item.get("JANコード", "")
        clean_name = item.get("商品名", "")
        full_prod_name = raw_title_map.get(sku, f"【{sku}】{clean_name}（{jan}）" if jan else f"【{sku}】{clean_name}")
        std_p = std_price_map.get(sku, "")

        r_dict = {col: "" for col in ORDER_CSV_COLUMNS}
        r_dict['顧客ID'] = est_data.get("cust_id", -1)
        r_dict['要望等'] = est_data.get("memo", "")
        r_dict['注文者お名前(姓)'] = est_data.get("cust_sei", "ご担当者")
        r_dict['注文者お名前(名)'] = est_data.get("cust_mei", "様")
        r_dict['注文者お名前(フリガナ・姓)'] = est_data.get("cust_kana_sei", "")
        r_dict['注文者お名前(フリガナ名)'] = est_data.get("cust_kana_mei", "")
        r_dict['注文者会社名'] = est_data.get("cust_comp", "")
        r_dict['注文者メールアドレス'] = est_data.get("cust_email", "")
        r_dict['注文者メールアドレスタイプ'] = 1
        r_dict['注文者電話番号'] = est_data.get("cust_tel", "")
        r_dict['注文者FAX'] = est_data.get("cust_fax", "")
        r_dict['注文者郵便番号'] = est_data.get("cust_post", "")
        r_dict['注文者都道府県'] = est_data.get("cust_pref", "")
        r_dict['注文者住所1'] = est_data.get("cust_addr1", "")
        r_dict['注文者住所2'] = est_data.get("cust_addr2", "")
        r_dict['注文者性別ID'] = 1
        r_dict['注文者性別'] = "男性"
        r_dict['注文者職種'] = est_data.get("cust_job_id", "")
        r_dict['商品合計'] = line_amt
        r_dict['調整額'] = 0
        r_dict['うち割引額１'] = 0
        r_dict['うち割引額２'] = 0
        r_dict['うち割引額３'] = 0
        r_dict['うちその他調整額'] = 0
        r_dict['合計額'] = line_amt
        r_dict['送料（税別）'] = est_data.get("shipping_fee", 0)
        r_dict['手数料（税別）'] = est_data.get("cod_fee", 0)
        r_dict['クーポン名'] = str(est_data.get("rank_name", "")).split(":")[0].strip()
        r_dict['クーポン利用'] = 0
        r_dict['ポイント利用'] = 0
        r_dict['請求額'] = line_amt
        r_dict['消費税額'] = unit_tax
        r_dict['入金額'] = 0
        r_dict['EC種類'] = 16
        r_dict['EC注文番号'] = est_data.get("q_no", "").replace("-", "")
        r_dict['EC店舗ID'] = 1000
        r_dict['EC店舗名'] = "旭電機化成"
        r_dict['販売ルートID'] = 1
        r_dict['受注ステータス'] = 1
        r_dict['支払方法ID'] = pay_id
        r_dict['決済ステータス'] = 0
        r_dict['配送方法ID'] = 1
        r_dict['配送区分'] = 0
        r_dict['申込番号種類'] = 0
        r_dict['注文日時'] = est_data.get("created_at", now_dt_str)
        r_dict['出荷予定日'] = ship_dt_str
        r_dict['ひとことメモ'] = est_data.get("hitokoto", "")
        r_dict['伝達事項'] = est_data.get("memo", "")
        r_dict['再計算フラグ'] = 1
        r_dict['送付先お名前(姓)'] = est_data.get("ship_sei", est_data.get("cust_sei", "ご担当者"))
        r_dict['送付先お名前(名)'] = est_data.get("ship_mei", est_data.get("cust_mei", "様"))
        r_dict['送付先会社名'] = est_data.get("ship_comp", est_data.get("cust_comp", ""))
        r_dict['送付先電話番号'] = est_data.get("ship_tel", est_data.get("cust_tel", ""))
        r_dict['送付先郵便番号'] = est_data.get("ship_post", est_data.get("cust_post", ""))
        r_dict['送付先都道府県'] = est_data.get("ship_pref", est_data.get("cust_pref", ""))
        r_dict['送付先住所1'] = est_data.get("ship_addr1", est_data.get("cust_addr1", ""))
        r_dict['送付先住所2'] = est_data.get("ship_addr2", est_data.get("cust_addr2", ""))
        r_dict['お届け日'] = deliv_dt_str
        r_dict['出力フラグ(受注伝票)'] = 0
        r_dict['出力フラグ(お手紙)'] = 0
        r_dict['出力フラグ(納品書)'] = 1
        r_dict['出力フラグ(領収書)'] = 0
        r_dict['出力フラグ(請求書)'] = 0
        r_dict['出力フラグ(EMS)'] = 0
        r_dict['出力フラグ(送り状)'] = 0
        for flag_col in [
            '出力済(受注伝票)', '出力済(お手紙)', '出力済(納品書)', '出力済(領収書)', '出力済(請求書)',
            '出力済(EMS)', '出力済(送り状)', '出力済(後払い取引登録)', '出力済(後払い出荷登録)',
            '出力済(配送指示伝票)', '出力済(モール連携出荷)', '出力済(モール連携処理済)',
            'メール送信済(サンクス)', 'メール送信済(入金確認)', 'メール送信済(発送完了)',
            'メール送信済(フォロー)', 'メール送信済(その他１)', 'メール送信済(その他２)'
        ]:
            r_dict[flag_col] = 0

        r_dict['ご購入明細商品コード'] = sku
        r_dict['ご購入明細商品名'] = full_prod_name
        r_dict['ご購入明細消費税区分'] = "別"
        r_dict['ご購入明細消費税課税規則'] = 3
        r_dict['ご購入明細送料区分'] = "別"
        r_dict['ご購入明細手数料区分'] = "別"
        r_dict['ご購入明細税抜単価'] = unit_p
        r_dict['ご購入明細税込単価'] = unit_tax_in
        r_dict['ご購入明細消費税'] = unit_tax
        r_dict['ご購入明細消費税率'] = 10
        r_dict['ご購入明細数量'] = qty
        r_dict['ご購入明細金額'] = line_amt
        r_dict['ご購入明細定期区分'] = "商品"
        r_dict['ご購入明細標準販売価格'] = std_p
        r_dict['消費税計算区分'] = "伝票単位(B)"
        rows.append(r_dict)

    out_df = pd.DataFrame(rows, columns=ORDER_CSV_COLUMNS)
    return out_df.to_csv(index=False, encoding="cp932", errors="replace").encode("cp932")

def _read_single_csv(f):
    path_date = None
    ym_match = re.search(r'(\d{4})年?[/\\_](\d{1,2})月?[/\\_]yayoi_(\d{2})(\d{2})', f, re.IGNORECASE)
    if ym_match:
        y, m, md_m, md_d = ym_match.groups()
        path_date = f"{y}{md_m}{md_d}"
    else:
        y_m = re.search(r'(\d{4})年', f)
        d_m = re.search(r'yayoi_(\d{4})', f, re.IGNORECASE)
        if y_m and d_m: path_date = f"{y_m.group(1)}{d_m.group(1)}"

    for enc in ["cp932", "shift_jis", "utf-8"]:
        try:
            d = pd.read_csv(f, encoding=enc, encoding_errors="replace", low_memory=False)
            cols = [str(c).strip() for c in d.columns]
            qty_cols = [c for c in cols if "数量" in c]
            d["_auto_qty"] = d[qty_cols[0]].apply(clean_num) if qty_cols else (
                d.iloc[:, 21].apply(clean_num) if len(cols) >= 22 else 0.0
            )
            price_cols = [c for c in cols if "単価" in c]
            d["_auto_price"] = d[price_cols[0]].apply(clean_num) if price_cols else (
                d.iloc[:, 20].apply(clean_num) if len(cols) >= 21 else 0.0
            )
            amt_cols = [c for c in cols if c in ["金額", "売上金額", "販売金額"]]
            if amt_cols:
                d["_auto_amt"] = d[amt_cols[0]].apply(clean_num)
                d["_auto_amt"] = d.apply(lambda r: (r["_auto_qty"] * r["_auto_price"]) if r["_auto_amt"] == 0 and r["_auto_price"] > 0 else r["_auto_amt"], axis=1)
            else:
                d["_auto_amt"] = d["_auto_qty"] * d["_auto_price"]
            d["_fallback_date"] = path_date
            d["_source_file"] = os.path.basename(f)
            return d
        except Exception: continue
    return None

def _read_amazon_seller_csv(f):
    norm_f = f.replace('\\', '/')
    ym_match = re.search(r'(\d{4})年?[/](\d{1,2})月?', norm_f)
    if ym_match:
        y_str = ym_match.group(1)
        m_str = f"{int(ym_match.group(2)):02d}"
        target_ym = f"{y_str}年{m_str}月"
        try: dt_last = pd.Timestamp(year=int(y_str), month=int(m_str), day=1) + pd.offsets.MonthEnd(1)
        except Exception: dt_last = datetime(int(y_str), int(m_str), 28)
    else:
        target_ym = "2026年09月"
        dt_last = datetime(2026, 9, 30)

    for enc in ["utf-8-sig", "utf-8", "cp932", "shift_jis"]:
        try:
            d = pd.read_csv(f, encoding=enc, encoding_errors="replace", low_memory=False)
            if d.empty: continue
            sku_series = None
            for c in d.columns:
                if str(c).strip().upper() == "SKU":
                    sku_series = d[c].astype(str).str.strip()
                    break
            if sku_series is None:
                sku_series = d.iloc[:, 3].astype(str).str.strip() if len(d.columns) >= 4 else d.iloc[:, 0].astype(str).str.strip()

            qty_series = None
            for c in d.columns:
                if any(k in str(c) for k in ['点数', '數量', '個数']):
                    qty_series = d[c].apply(clean_num).astype(int)
                    break
            if qty_series is None:
                if len(d.columns) >= 15: qty_series = d.iloc[:, 14].apply(clean_num).astype(int)
                elif len(d.columns) >= 3: qty_series = d.iloc[:, 2].apply(clean_num).astype(int)
                else: qty_series = pd.Series([1] * len(d))

            amt_series = None
            for c in d.columns:
                if any(k in str(c) for k in ['売上額', '売上金額', '売上', '金額']):
                    gross = d[c].apply(clean_num)
                    amt_series = (gross / 1.10).round().astype(int)
                    break
            if amt_series is None:
                if len(d.columns) >= 19: gross = d.iloc[:, 18].apply(clean_num); amt_series = (gross / 1.10).round().astype(int)
                elif len(d.columns) >= 2: gross = d.iloc[:, 1].apply(clean_num); amt_series = (gross / 1.10).round().astype(int)
                else: continue

            res = pd.DataFrame({
                "商品コード": sku_series, "商品名": sku_series, "得意先名": "アマゾンセラー", "担当者": "水原",
                "数量": qty_series, "金額": amt_series, "日付_dt": dt_last,
                "日付": dt_last.strftime('%Y/%m/%d'), "年月": target_ym,
                "_source_file": os.path.basename(f)
            })
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
    if c_pcode: res["商品コード"] = df_all[c_pcode].astype(str).str.strip()
    if c_pname:
        res["元商品名"] = df_all[c_pname].astype(str).str.strip()
        res["商品名"] = df_all[c_pname].astype(str).map(clean_product_name)
    if c_cname: res["得意先名"] = df_all[c_cname].astype(str).str.strip().map(clean_cust_name)
    else: res["得意先名"] = ""

    if c_staff: res["担当者"] = df_all[c_staff].map(convert_staff_name).fillna("").replace("nan", "")
    else: res["担当者"] = ""

    res["数量"] = df_all["_auto_qty"].fillna(0).astype(int)
    res["単価"] = df_all["_auto_price"].fillna(0).astype(int)
    res["金額"] = df_all["_auto_amt"].fillna(res["数量"] * res["単価"]).astype(int)
    if "_source_file" in df_all.columns: res["_source_file"] = df_all["_source_file"]

    raw_s = df_all[c_date].astype(str).str.strip() if c_date else pd.Series([""] * len(df_all))
    clean_date = raw_s.str.replace(r'\.0$', '', regex=True).str.replace(r'[-/.\s]', '', regex=True)
    if "_fallback_date" in df_all.columns:
        clean_date = clean_date.where(clean_date.str.len() == 8, df_all["_fallback_date"].astype(str))

    dt = pd.to_datetime(clean_date, format="%Y%m%d", errors="coerce").fillna(pd.to_datetime(raw_s, errors="coerce"))
    res["日付_dt"] = dt
    res["日付"] = dt.dt.strftime('%Y/%m/%d').fillna("")
    res["年月"] = dt.dt.strftime('%Y年%m月').fillna("")
    if "商品名" in res.columns: res = res[~res["商品名"].str.contains("送料", na=False)].copy()
    return res

@st.cache_data(show_spinner=False)
def load_and_prep_data(force_reload=False):
    cost_path, cost_mtime = get_cost_master_file_info()
    cost_master = load_cost_master(cost_path, cost_mtime)

    if not force_reload and os.path.exists(CACHE_FILE):
        try:
            cached_df = pd.read_parquet(CACHE_FILE)
            if not cached_df.empty:
                cached_df["日付_dt"] = pd.to_datetime(cached_df["日付_dt"])
                cached_df["単価原価"] = cached_df["商品コード"].map(lambda c: cost_master.get(str(c).strip(), 0.0))
                cached_df["原価合計"] = (cached_df["数量"] * cached_df["単価原価"]).round().astype(int)
                cached_df["粗利額"] = cached_df["金額"] - cached_df["原価合計"]
                return cached_df, 1, cost_master, 0
        except Exception: pass

    csv_candidates = glob.glob(os.path.join(BASE_SHARE, "**", "*.csv"), recursive=True)
    csv_candidates = sorted(list(set(csv_candidates)))

    yayoi_files = [f for f in csv_candidates if "yayoi" in f.lower()]
    amz_files = [f for f in csv_candidates if "アマゾン売上" in f or "businessreport" in f.lower()]
    if not yayoi_files: yayoi_files = [f for f in csv_candidates if f not in amz_files]

    with ThreadPoolExecutor(max_workers=16) as executor:
        res_list = list(executor.map(_read_single_csv, yayoi_files))
    valid_res = [d for d in res_list if d is not None and not d.empty]
    new_df = _parse_raw_df_to_standard(pd.concat(valid_res, ignore_index=True)) if valid_res else pd.DataFrame()

    amz_res = [_read_amazon_seller_csv(af) for af in amz_files]
    valid_amz = [ad for ad in amz_res if ad is not None and not ad.empty]
    new_amz_df = pd.concat(valid_amz, ignore_index=True) if valid_amz else pd.DataFrame()

    valid_dfs = [d for d in [new_df, new_amz_df] if d is not None and not d.empty]
    df = pd.concat(valid_dfs, ignore_index=True) if valid_dfs else pd.DataFrame()

    if df.empty or ("アマゾンセラー" not in df.get("得意先名", pd.Series()).values):
        amz_rows = []
        sep_date = pd.Timestamp(2026, 9, 30)
        for r in DEFAULT_AMZ_SEP_2026:
            u_p = int(r["amount"] / r["qty"]) if r["qty"] > 0 else r["amount"]
            amz_rows.append({
                "商品コード": r["sku"], "商品名": r["sku"], "得意先名": "アマゾンセラー", "担当者": "水原",
                "数量": r["qty"], "単価": u_p, "金額": r["amount"], "日付_dt": sep_date,
                "日付": "2026/09/30", "年月": "2026年09月", "_source_file": "default_amz"
            })
        df = pd.concat([df, pd.DataFrame(amz_rows)], ignore_index=True)

    if not df.empty and "担当者" in df.columns and "得意先名" in df.columns:
        valid_staff_df = df[df["担当者"].astype(str).str.strip().ne("")].sort_values("日付_dt")
        if not valid_staff_df.empty:
            cust_current_staff_map = valid_staff_df.groupby("得意先名")["担当者"].last().to_dict()
            df["担当者"] = df["得意先名"].map(lambda c: cust_current_staff_map.get(c, ""))

    if not df.empty and "商品コード" in df.columns:
        df["単価原価"] = df["商品コード"].map(lambda c: cost_master.get(str(c).strip(), 0.0))
        df["原価合計"] = (df["数量"] * df["単価原価"]).round().astype(int)
        df["粗利額"] = df["金額"] - df["原価合計"]

    try:
        if not df.empty: df.to_parquet(CACHE_FILE, index=False)
    except Exception: pass

    return df, len(csv_candidates), cost_master, len(csv_candidates)

def render_universal_table(df_data):
    center_cols = {"番号", "順位", "日付", "販売期間", "最終売上日", "初回売上日"}
    left_cols = {"得意先名", "商品コード", "商品名", "商品番号", "ランク名", "担当者", "客先名", "販売価格帯"}

    header_html = "".join([f"<th style='border:1px solid #cbd5e1; padding:8px 10px; background:#f1f5f9; text-align:{('center' if c in center_cols else ('left' if c in left_cols else 'right'))}; white-space:nowrap;'>{c}</th>" for c in df_data.columns])
    body_rows = []
    for r_i, row in df_data.iterrows():
        bg = "#ffffff" if r_i % 2 == 0 else "#f8fafc"
        cells = []
        for c in df_data.columns:
            val = row[c]
            val_str = "" if pd.isna(val) else str(val)
            align = "center" if c in center_cols else ("left" if c in left_cols else "right")
            if isinstance(val, (int, float)) and c not in center_cols:
                val_str = f"{int(val):+,}" if c in ["増減額", "目標差額"] else f"{int(val):,}"

            is_neg = False
            if isinstance(val, (int, float)) and val < 0: is_neg = True
            elif isinstance(val_str, str) and (val_str.startswith("-") or val_str.startswith("▲-")): is_neg = True

            style_extra = "color: #dc2626; font-weight: bold;" if is_neg else "color: #0f172a;"
            cells.append(f"<td style='border:1px solid #e2e8f0; padding:6px 10px; text-align:{align}; font-variant-numeric:tabular-nums; {style_extra} white-space:nowrap;'>{val_str}</td>")
        body_rows.append(f"<tr style='background:{bg};'>{''.join(cells)}</tr>")

    table_html = f"""
    <div style="overflow-x: auto; max-height: 600px; border: 1px solid #cbd5e1; border-radius: 6px; background-color: #ffffff;">
    <table style="width: 100%; border-collapse: collapse; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; font-size: 13px;">
    <thead><tr style="position: sticky; top: 0; z-index: 1;">{header_html}</tr></thead>
    <tbody>{''.join(body_rows)}</tbody>
    </table>
    </div>
    """
    calc_h = min(620, 80 + len(df_data) * 36)
    components.html(table_html, height=calc_h, scrolling=True)

df, file_count, cost_master_dict, diff_cnt = load_and_prep_data(force_reload=False)

member_p_path, member_p_mtime = get_member_price_file_info()
member_cust_sku_map, member_default_min_map = load_member_price_master_v2(member_p_path, member_p_mtime)

jan_path, jan_mtime = get_jan_file_info()
jan_master_dict, jan_title_dict, jan_std_price_dict = load_jan_master_v2(jan_path, jan_mtime)

cust_path, cust_mtime = get_customer_file_info()
customer_master_df = load_customer_master(cust_path, cust_mtime)

hanko_bytes, hanko_file_path = get_hanko_bytes()

col_status, col_refresh = st.columns([4, 1.2])
with col_refresh:
    if st.button("🔄 データを最新に再読み込み"):
        st.cache_data.clear()
        if os.path.exists(CACHE_FILE): os.remove(CACHE_FILE)
        with st.spinner("最新データを高速読み込み中..."):
            df, file_count, cost_master_dict, diff_cnt = load_and_prep_data(force_reload=True)
            member_cust_sku_map, member_default_min_map = load_member_price_master_v2(member_p_path, member_p_mtime)
            jan_master_dict, jan_title_dict, jan_std_price_dict = load_jan_master_v2(jan_path, jan_mtime)
            customer_master_df = load_customer_master(cust_path, cust_mtime)
            hanko_bytes, hanko_file_path = get_hanko_bytes()
        st.success("最新データに更新しました！")
        time.sleep(0.5)
        st.rerun()

if df.empty:
    st.error(f"売上CSVデータが見つかりませんでした。`{BASE_SHARE}` フォルダを確認してください。")
    st.stop()

valid_ym = sorted([x for x in df["年月"].unique() if x and str(x).strip() != ""])
latest_ym = valid_ym[-1] if valid_ym else "2026年01月"

valid_dts = df["日付_dt"].dropna()
min_date = valid_dts.min().date() if not valid_dts.empty else datetime(2022, 1, 1).date()
max_date = valid_dts.max().date() if not valid_dts.empty else datetime.now().date()
desc_valid_ym = sorted(valid_ym, reverse=True)

col_top1, col_top2 = st.columns([1.6, 2])
with col_top1:
    view_type = st.selectbox("表示モード", [
        '🧮 簡易見積＆利益シミュレーター',
        '🏢 得意先別ランキング',
        '📦 商品別売上ランキング',
        '📊 営業会議用：上位社 売上比較シート',
        '📑 営業会議用：上位客先別売上実績',
        '👤 担当者別 売上＆目標達成分析',
        '💤 滞留・未販売商品分析（売れてない商品）',
        '✨ 新商品売上分析（直近1年）',
        '📈 商品別 前年実績比較',
        '📊 得意先別 前年実績比較',
        '🧩 得意先×商品 マトリクス',
        '📋 明細一覧'
    ])
with col_top2:
    if view_type == '🧮 簡易見積＆利益シミュレーター':
        st.write("")
        jan_status = f"JAN: **{len(jan_master_dict):,} 件**" if jan_master_dict else "⚠️ JAN未検出"
        hanko_status = "角印: **読込完了**" if hanko_bytes else "⚠️ 印鑑未検出"
        member_status = f"加盟店価格: **{len(member_default_min_map):,} 件**" if member_default_min_map else "⚠️ 加盟店価格未検出"
        st.info(f"📁 原価: **{len(cost_master_dict):,} 件** | {member_status} | {jan_status} | {hanko_status}")
    else:
        quick_choice = st.radio("期間ショートカット", ['当月', '前月', '当期累計', '前期累計', '全データ累計', '手動指定'], horizontal=True, key="period_shortcut_radio_top")

# ========================================================
# ★ 🧮 簡易見積＆利益シミュレーター
# ========================================================
if view_type == '🧮 簡易見積＆利益シミュレーター':
    code_to_name_map = df.dropna(subset=["商品コード", "商品名"]).drop_duplicates(subset=["商品コード"]).set_index("商品コード")["商品名"].to_dict()
    for k_t, v_t in jan_title_dict.items():
        if k_t not in code_to_name_map or not code_to_name_map[k_t]:
            code_to_name_map[k_t] = v_t

    raw_prod_title_map = {}
    if "元商品名" in df.columns:
        raw_prod_title_map = df.dropna(subset=["商品コード", "元商品名"]).drop_duplicates(subset=["商品コード"]).set_index("商品コード")["元商品名"].to_dict()
    for k_t, v_t in jan_title_dict.items():
        if k_t not in raw_prod_title_map: raw_prod_title_map[k_t] = v_t

    all_calc_codes = sorted(list(set(list(code_to_name_map.keys()) + list(cost_master_dict.keys()) + list(jan_master_dict.keys()) + list(member_default_min_map.keys()))))

    STANDARD_RANKS = [
        "ロット単価",
        "1.05倍", "1.07倍", "1.1倍", "1.15倍", "1.2倍", "1.25倍",
        "1.3倍", "1.35倍", "1.4倍", "1.45倍", "1.5倍", "0.9倍"
    ]

    RANK_MULTIPLIERS = {
        "1.05倍": 1.05, "1.07倍": 1.07, "1.1倍": 1.10, "1.15倍": 1.15,
        "1.2倍": 1.20, "1.25倍": 1.25, "1.3倍": 1.30, "1.35倍": 1.35,
        "1.4倍": 1.40, "1.45倍": 1.45, "1.5倍": 1.50, "0.9倍": 0.90
    }

    def get_sku_details(sku_code, rank_selection="ロット単価", cust_comp_name=""):
        p_name = code_to_name_map.get(sku_code, "")
        if not p_name: p_name = jan_title_dict.get(sku_code, sku_code)
        p_jan = jan_master_dict.get(sku_code, "")
        p_cost = int(cost_master_dict.get(sku_code, 0.0))
        
        norm_cust = re.sub(r'[\s ]+', '', str(cust_comp_name)) if cust_comp_name else ""
        min_p = member_default_min_map.get(sku_code, 0)

        comp_lot_p = 0
        if cust_comp_name and (cust_comp_name, sku_code) in member_cust_sku_map:
            comp_lot_p = member_cust_sku_map[(cust_comp_name, sku_code)]
        elif norm_cust and (norm_cust, sku_code) in member_cust_sku_map:
            comp_lot_p = member_cust_sku_map[(norm_cust, sku_code)]

        clean_r = str(rank_selection).strip() if rank_selection else "ロット単価"
        p_price = 0
        actual_rank_applied = clean_r

        if clean_r in RANK_MULTIPLIERS:
            if min_p > 0:
                p_price = int(round(min_p * RANK_MULTIPLIERS[clean_r]))
            else:
                p_price = 99999
        elif clean_r == "ロット単価" or clean_r == "":
            if comp_lot_p > 0:
                p_price = comp_lot_p
            else:
                p_price = 99999

        if p_price == 0: p_price = 99999
        return p_name, p_jan, p_price, p_cost, min_p, actual_rank_applied

    if "load_target_estimate" in st.session_state and st.session_state["load_target_estimate"]:
        loaded_est = st.session_state.pop("load_target_estimate")
        st.session_state["in_cust_comp"] = loaded_est.get("cust_comp", "")
        st.session_state["in_cust_person"] = loaded_est.get("cust_person", "")
        st.session_state["in_cust_post"] = loaded_est.get("cust_post", "")
        st.session_state["in_cust_addr"] = loaded_est.get("cust_addr", "")
        st.session_state["in_ship_comp"] = loaded_est.get("ship_comp", "")
        st.session_state["in_ship_person"] = loaded_est.get("ship_person", "")
        st.session_state["in_ship_post"] = loaded_est.get("ship_post", "")
        st.session_state["in_ship_addr"] = loaded_est.get("ship_addr", "")
        st.session_state["in_ship_tel"] = loaded_est.get("ship_tel", "")
        st.session_state["sel_cust_choice"] = loaded_est.get("cust_comp", "（手動入力）")
        st.session_state.quotation_items = loaded_est.get("items", [])
        st.session_state.editor_version += 1
        st.session_state["est_active_tab"] = "📝 新規見積作成・シミュレーター"

    if "quotation_items" not in st.session_state: st.session_state.quotation_items = []
    if "saved_estimates" not in st.session_state: st.session_state.saved_estimates = load_saved_quotes()
    if "editor_version" not in st.session_state: st.session_state.editor_version = 0

    tab_choice = st.radio(
        "見積機能メニュー",
        ["📝 新規見積作成・シミュレーター", f"📋 保存済み見積一覧・スマレジ注文CSV出力（{len(st.session_state.saved_estimates)}件）"],
        horizontal=True,
        key="est_active_tab"
    )

    if tab_choice.startswith("📝"):
        st.markdown("##### 🏢 提出先（得意先）の選択")
        cust_options = ["（手動入力）"]
        cust_lookup = {}
        if not customer_master_df.empty:
            c_name_col = "加盟店名" if "加盟店名" in customer_master_df.columns else customer_master_df.columns[0]
            for _, c_row in customer_master_df.iterrows():
                c_name = str(c_row[c_name_col]).strip()
                if c_name and c_name != "nan":
                    cust_options.append(c_name)
                    cust_lookup[c_name] = c_row

        for k in ["in_cust_comp", "in_cust_person", "in_cust_post", "in_cust_addr",
                  "in_ship_comp", "in_ship_person", "in_ship_post", "in_ship_addr", "in_ship_tel"]:
            if k not in st.session_state: st.session_state[k] = ""

        def on_customer_change():
            selected = st.session_state.get("sel_cust_choice", "（手動入力）")
            if selected == "（手動入力）" or selected not in cust_lookup:
                for k in ["in_cust_comp", "in_cust_person", "in_cust_post", "in_cust_addr", "in_ship_comp", "in_ship_person", "in_ship_post", "in_ship_addr", "in_ship_tel"]:
                    st.session_state[k] = ""
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

                st.session_state["in_cust_comp"] = selected
                st.session_state["in_cust_person"] = p_name
                st.session_state["in_cust_post"] = p_code
                st.session_state["in_cust_addr"] = full_addr
                st.session_state["in_ship_comp"] = selected
                st.session_state["in_ship_person"] = p_name
                st.session_state["in_ship_post"] = p_code
                st.session_state["in_ship_addr"] = full_addr
                st.session_state["in_ship_tel"] = tel

        st.selectbox("🔍 加盟店・取引先を選択（検索可能）", cust_options, key="sel_cust_choice", on_change=on_customer_change)

        col_addr1, col_addr2 = st.columns(2)
        with col_addr1:
            st.markdown("###### 📄 宛先情報")
            q_cust_comp = st.text_input("宛先 会社名", key="in_cust_comp")
            q_cust_person = st.text_input("宛先 担当者名", key="in_cust_person")
            q_cust_post = st.text_input("宛先 郵便番号", key="in_cust_post")
            q_cust_addr = st.text_input("宛先 住所", key="in_cust_addr")
        with col_addr2:
            st.markdown("###### 🚚 お届け先情報")
            ship_cust_comp = st.text_input("お届け先 会社名", key="in_ship_comp")
            ship_cust_person = st.text_input("お届け先 担当者名", key="in_ship_person")
            ship_cust_post = st.text_input("お届け先 郵便番号", key="in_ship_post")
            ship_cust_addr = st.text_input("お届け先 住所", key="in_ship_addr")
            ship_cust_tel = st.text_input("お届け先 TEL", key="in_ship_tel")

        hq1, hq2, hq3, hq4 = st.columns(4)
        with hq1: q_no = st.text_input("見積番号", value=f"{datetime.now().strftime('%Y%m%d')}-01")
        with hq2: q_date = st.date_input("お見積日", value=datetime.now().date())
        with hq3: q_exp_date = st.date_input("見積有効期限", value=(datetime.now() + timedelta(days=30)).date())
        with hq4: q_payment = st.selectbox("お支払い方法", ["銀行振込", "代金引換", "現行通り", "掛け払い（月末締め翌月末払い）"], index=0)

        hq5, hq6, hq7 = st.columns([2, 2, 1])
        with hq5: q_memo = st.text_input("伝達事項 / 発注NO等", placeholder="例: 発注NO.2352640")
        with hq6: q_staff_memo = st.text_input("営業担当・社内備考", placeholder="例: 営業担当：水原")
        with hq7: q_shipping_fee = st.number_input("送料(税抜)", min_value=0, value=0, step=100)

        st.markdown("---")
        st.markdown("##### 📑 複数品目のまとめ見積リスト")
        cols_order = ["選択", "商品コード", "商品名", "JANコード", "最低下代", "適用ランク", "数量", "見積単価", "合計金額", "単価原価", "粗利額", "粗利率"]

        if st.session_state.quotation_items:
            raw_items_df = pd.DataFrame(st.session_state.quotation_items)
            for c in cols_order:
                if c not in raw_items_df.columns:
                    raw_items_df[c] = False if c == "選択" else ("ロット単価" if c == "適用ランク" else ("-" if c == "粗利率" else 0))
            raw_items_df["数量"] = raw_items_df["数量"].apply(lambda x: max(1, int(clean_num(x))))
            raw_items_df["合計金額"] = raw_items_df["数量"] * raw_items_df["見積単価"]
            raw_items_df = raw_items_df[cols_order]
        else:
            raw_items_df = pd.DataFrame([{
                "選択": False, "商品コード": "", "商品名": "", "JANコード": "", "最低下代": 0,
                "適用ランク": "ロット単価", "数量": 1, "見積単価": 0, "合計金額": 0, "単価原価": 0, "粗利額": 0, "粗利率": "-"
            }], columns=cols_order)

        edited_df = st.data_editor(
            raw_items_df,
            column_config={
                "選択": st.column_config.CheckboxColumn("選択", default=False, width="small"),
                "商品コード": st.column_config.SelectboxColumn("商品コード", options=all_calc_codes, required=True, width="medium"),
                "商品名": st.column_config.TextColumn("商品名", width="large"),
                "適用ランク": st.column_config.SelectboxColumn("適用ランク", options=STANDARD_RANKS, width="medium"),
                "数量": st.column_config.NumberColumn("数量", min_value=1, step=1),
                "見積単価": st.column_config.NumberColumn("見積単価", min_value=0, step=10, format="¥%d"),
            },
            num_rows="dynamic",
            use_container_width=True,
            hide_index=True,
            key=f"quote_editor_{st.session_state.editor_version}"
        )

        updated_records = []
        for idx_e, row_e in edited_df.iterrows():
            c_code = str(row_e["商品コード"]).strip() if pd.notna(row_e["商品コード"]) else ""
            if not c_code or c_code == "None": continue
            cur_qty = max(1, int(clean_num(row_e.get("数量", 1))))
            cur_rank = str(row_e.get("適用ランク", "ロット単価")).strip()
            _, auto_jan, auto_price, auto_cost, auto_min, applied_lbl = get_sku_details(c_code, cur_rank, q_cust_comp)
            cur_price = int(clean_num(row_e.get("見積単価", auto_price)))
            new_total = cur_qty * cur_price
            new_profit = new_total - (cur_qty * auto_cost)
            new_margin_str = f"{(new_profit / new_total * 100):.1f}%" if new_total > 0 else "-"

            updated_records.append({
                "選択": False, "商品コード": c_code, "商品名": code_to_name_map.get(c_code, c_code),
                "JANコード": str(auto_jan), "最低下代": auto_min, "適用ランク": applied_lbl,
                "数量": cur_qty, "見積単価": cur_price, "合計金額": new_total,
                "単価原価": auto_cost, "粗利額": new_profit, "粗利率": new_margin_str
            })
        st.session_state.quotation_items = updated_records

        if updated_records:
            calc_df = pd.DataFrame(updated_records)
            q_tot_qty = int(calc_df["数量"].sum())
            q_tot_sales = int(calc_df["合計金額"].sum())
            q_tot_profit = int(calc_df["粗利額"].sum())
            q_tax = int(round((q_tot_sales + q_shipping_fee) * 0.10))
            q_grand_total = q_tot_sales + q_shipping_fee + q_tax
            current_estimate_obj = {
                "q_no": q_no, "created_at": datetime.now().strftime("%Y/%m/%d %H:%M"),
                "q_date": q_date.strftime("%Y/%m/%d"), "q_exp_date": q_exp_date.strftime("%Y/%m/%d"),
                "payment": q_payment, "cust_comp": q_cust_comp, "grand_total": q_grand_total, "items": updated_records
            }
            if st.button("💾 この見積を保存して一覧へ登録"):
                st.session_state.saved_estimates.insert(0, current_estimate_obj)
                save_quotes_list(st.session_state.saved_estimates)
                st.success("✅ 登録しました！")

    elif tab_choice.startswith("📋"):
        st.markdown("##### 📋 保存済み見積一覧")
        for idx_q, est in enumerate(st.session_state.saved_estimates):
            st.write(f"見積番号: {est.get('q_no')} ｜ 会社名: {est.get('cust_comp')} ｜ 総額: ¥{est.get('grand_total'):,}")
    st.stop()

# ========================================================
# ★ 各種分析ビューの設定＆分岐
# ========================================================
idx = valid_ym.index(latest_ym) if latest_ym in valid_ym else len(valid_ym) - 1
curr_y = int(latest_ym[:4]); curr_m = int(latest_ym[5:7])
fiscal_year = curr_y if curr_m >= 7 else curr_y - 1

quick_choice = st.radio("期間ショートカット", ['当月', '前月', '当期累計', '前期累計', '全データ累計', '手動指定'], horizontal=True, key="period_shortcut_radio_analysis")

if quick_choice == '当月': s_date_calc, e_date_calc = datetime(curr_y, curr_m, 1).date(), max_date
elif quick_choice == '前月': 
    prev = valid_ym[max(0, idx - 1)]
    py, pm = int(prev[:4]), int(prev[5:7])
    s_date_calc, e_date_calc = datetime(py, pm, 1).date(), (pd.Timestamp(year=py, month=pm, day=1) + pd.offsets.MonthEnd(1)).date()
elif quick_choice == '当期累計': s_date_calc, e_date_calc = datetime(fiscal_year, 7, 1).date(), max_date
elif quick_choice == '前期累計': s_date_calc, e_date_calc = datetime(fiscal_year - 1, 7, 1).date(), datetime(fiscal_year, 6, 30).date()
else: s_date_calc, e_date_calc = min_date, max_date

col_mode, col_p1, col_p2, col_staff = st.columns([1, 1.2, 1.2, 1.6])
with col_mode: date_mode = st.radio("指定単位", ["年月で指定", "日で指定"], horizontal=True, key="date_mode_radio_analysis")

if date_mode == "年月で指定":
    with col_p1: ym_start = st.selectbox("開始年月", desc_valid_ym, index=0)
    with col_p2: ym_end = st.selectbox("終了年月", desc_valid_ym, index=0)
else:
    with col_p1: d_start = st.date_input("開始日", value=s_date_calc, min_value=min_date, max_value=max_date)
    with col_p2: d_end = st.date_input("終了日", value=e_date_calc, min_value=min_date, max_value=max_date)

with col_staff:
    staff_options = ["（すべて）"] + sorted([x for x in df["担当者"].unique() if x and x != "nan"])
    sel_staff = st.selectbox("担当者指定", staff_options)

all_unique_custs = sorted([x for x in df["得意先名"].unique() if x])

if view_type not in ['📑 営業会議用：上位客先別売上実績', '📊 営業会議用：上位社 売上比較シート']:
    c1, c2, c3, _ = st.columns(3)
    with c1: sel_cust = st.selectbox("得意先名", ["（すべて）"] + all_unique_custs)
    with c2: search_cust = st.text_input("得意先検索", placeholder="会社名の一部...")
    with c3: limit_val = st.selectbox("表示件数", [50, 100, 200, "全件"], index=0)

one_year_ago_date = max_date - pd.DateOffset(years=1)

if view_type == '💤 滞留・未販売商品分析（売れてない商品）':
    dormant_mode = st.selectbox("滞留基準", ["半年以上", "1年以上"], index=0)
    filtered = df.copy(); period_text = f"基準日: {max_date.strftime('%Y/%m/%d')} （{dormant_mode}）"
elif view_type == '✨ 新商品売上分析（直近1年）':
    filtered = df[df["日付_dt"] >= one_year_ago_date].copy(); period_text = "直近1年間"
elif view_type == '📊 営業会議用：上位社 売上比較シート':
    col_m1, col_m2 = st.columns(2)
    with col_m1: target_meeting_ym = st.selectbox("対象月度", desc_valid_ym, index=0, key="tgt_ym_1")
    with col_m2: compare_cust_count = st.selectbox("比較上位社数", [10, 20, 30, 50], index=0)
    
    tgt_y = int(target_meeting_ym[:4]); tgt_m = int(target_meeting_ym[5:7]); ly_ym = f"{tgt_y - 1}年{tgt_m:02d}月"
    df_scope = df if sel_staff == "（すべて）" else df[df["担当者"] == sel_staff]
    top_custs = df_scope[df_scope["年月"] == target_meeting_ym].groupby("得意先名")["金額"].sum().nlargest(compare_cust_count).index.tolist()
    st.markdown(f"#### 📊 {tgt_m}月度 上位{len(top_custs)}社 売上比較シート（担当: {sel_staff}）")
    st.stop()
elif view_type == '📑 営業会議用：上位客先別売上実績':
    col_m1, col_m2 = st.columns(2)
    with col_m1: target_meeting_ym = st.selectbox("対象月度", desc_valid_ym, index=0, key="tgt_ym_2")
    with col_m2: meeting_cust_count = st.selectbox("表示上位数", [20, 50, 100, 200], index=0)

    tgt_y = int(target_meeting_ym[:4]); tgt_m = int(target_meeting_ym[5:7]); ly_ym = f"{tgt_y - 1}年{tgt_m:02d}月"
    df_scope = df if sel_staff == "（すべて）" else df[df["担当者"] == sel_staff]
    cur_cust_sales = df_scope[df_scope["年月"] == target_meeting_ym].groupby("得意先名")["金額"].sum()
    ly_cust_sales = df_scope[df_scope["年月"] == ly_ym].groupby("得意先名")["金額"].sum()
    target_cust_list = cur_cust_sales.nlargest(meeting_cust_count).index.tolist()
    m_rows = []
    for rank, cust in enumerate(target_cust_list, 1):
        s_cur = int(cur_cust_sales.get(cust, 0)); s_ly = int(ly_cust_sales.get(cust, 0))
        m_rows.append({"順位": str(rank), "客先名": cust, "売上実績": s_cur, "前年実績": s_ly, "前年比": f"{(s_cur / s_ly * 100):.0f}%" if s_ly > 0 else "-", "増減額": s_cur - s_ly})
    render_universal_table(pd.DataFrame(m_rows))
    st.stop()
else:
    if date_mode == "年月で指定":
        if ym_start > ym_end: ym_start, ym_end = ym_end, ym_start
        filtered = df[(df["年月"] >= ym_start) & (df["年月"] <= ym_end)]
        period_text = f"{ym_start} 〜 {ym_end}"
    else:
        dt_s, dt_e = pd.to_datetime(d_start), pd.to_datetime(d_end)
        filtered = df[(df["日付_dt"] >= dt_s) & (df["日付_dt"] <= dt_e)]
        period_text = f"{d_start} 〜 {d_end}"

if view_type not in ['📑 営業会議用：上位客先別売上実績', '📊 営業会議用：上位社 売上比較シート']:
    if sel_cust != "（すべて）": filtered = filtered[filtered["得意先名"] == sel_cust]
    if search_cust.strip():
        kw_clean = re.sub(r'[\s ]+', '', search_cust.strip())
        filtered = filtered[filtered["得意先名"].map(lambda x: kw_clean in re.sub(r'[\s ]+', '', str(x)))]
    if sel_staff != "（すべて）": filtered = filtered[filtered["担当者"] == sel_staff]

    disp_df = filtered.drop(columns=["年月", "日付_dt", "元商品名"], errors="ignore").sort_values(by="金額", ascending=False)
    disp_df.insert(0, "番号", range(1, len(disp_df) + 1))
    if limit_val != "全件": disp_df = disp_df.head(int(limit_val))
    render_universal_table(disp_df)
