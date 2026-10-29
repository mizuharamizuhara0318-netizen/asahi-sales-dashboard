import os
import io
import time
import json
import glob
import re
import base64
import warnings
from datetime import datetime, timedelta
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
import plotly.express as px
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
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
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-header">売上管理＆営業分析表 (Supabase連携版)</div>', unsafe_allow_html=True)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OTHER_DATA_DIR = os.path.join(BASE_DIR, "その他データ")
SAVED_QUOTES_FILE = os.path.join(BASE_DIR, "saved_estimates.json")

STAFF_MAP = {
    "1": "小泉", "2": "河端", "3": "安部", "4": "原",
    "5": "西山", "6": "萩原", "7": "田中", "8": "水原",
    "9": "川村", "10": "穴山", "11": "鈴木", "12": "松平"
}

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
    if val is None: return 0.0
    try:
        if pd.isna(val): return 0.0
    except:
        pass
    s = re.sub(r'[^\d.-]', '', str(val))
    try: return float(s) if s else 0.0
    except: return 0.0

def convert_staff_name(val):
    if val is None: return ""
    v = str(val).strip().replace(".0", "")
    return STAFF_MAP.get(v, v)

def clean_product_name(val):
    if not isinstance(val, str) or not val.strip(): return ""
    if "】" in val: val = val.split("】", 1)[-1]
    return re.sub(r'[\(（]\s*[\d０-９]{7,16}\s*[\)）]', '', val).strip()

def find_file_multidir(filename_pattern):
    search_dirs = [BASE_DIR, OTHER_DATA_DIR]
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
                    if sk and sk != "ABA-40": cost_dict[sk] = co
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
                if not sku or sku == "nan" or sku == "ABA-40": continue

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
                if not sk or sk == "nan" or sk == "ABA-40": continue
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
    search_dirs = [BASE_DIR, OTHER_DATA_DIR]
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

@st.cache_data(show_spinner=False)
def load_data_from_supabase():
    try:
        from supabase import create_client, Client
        url = st.secrets.get("SUPABASE_URL", os.environ.get("SUPABASE_URL", ""))
        key = st.secrets.get("SUPABASE_KEY", os.environ.get("SUPABASE_KEY", ""))
        if not url or not key:
            return pd.DataFrame()
        
        supabase: Client = create_client(url, key)
        response = supabase.table("sales_data").select("*").execute()
        data = response.data
        if not data:
            return pd.DataFrame()
        
        df = pd.DataFrame(data)
        return df
    except Exception as e:
        st.error(f"Supabaseからのデータ取得に失敗しました: {e}")
        return pd.DataFrame()

def prep_supabase_dataframe(df, cost_master):
    if df is None or (hasattr(df, "empty") and df.empty):
        return pd.DataFrame(columns=["商品コード", "商品名", "得意先名", "担当者", "数量", "単価", "金額", "日付_dt", "日付", "年月", "単価原価", "原価合計", "粗利額"])

    if not isinstance(df, pd.DataFrame):
        df = pd.DataFrame(df)

    cols = [str(c).strip() for c in df.columns]

    def find_col(target_keywords):
        for c in cols:
            c_low = c.lower()
            for kw in target_keywords:
                if kw in c_low:
                    return c
        return None

    c_code = find_col(["商品コード", "商品cd", "sku", "品番", "code"])
    c_name = find_col(["商品名", "品名", "name"])
    c_cust = find_col(["得意先名", "顧客", "会社", "customer"])
    c_staff = find_col(["担当", "staff"])
    c_qty = find_col(["数量", "個数", "qty", "quantity"])
    c_price = find_col(["単価", "price", "unit_price", "unit"])
    c_amt = find_col(["金額", "売上", "amount", "total"])
    c_date = find_col(["日付", "date", "day"])

    new_df = pd.DataFrame()
    new_df["商品コード"] = df[c_code].astype(str).str.strip() if c_code in df.columns else ""
    new_df["商品名"] = df[c_name].astype(str).map(clean_product_name) if c_name in df.columns else ""
    new_df["得意先名"] = df[c_cust].astype(str).map(clean_cust_name) if c_cust in df.columns else ""
    new_df["担当者"] = df[c_staff].map(convert_staff_name).fillna("").astype(str) if c_staff in df.columns else ""
    
    raw_qty = df[c_qty] if c_qty in df.columns else 0
    new_df["数量"] = pd.to_numeric(pd.Series(raw_qty).apply(clean_num), errors="coerce").fillna(0).astype(int)

    raw_price = df[c_price] if c_price in df.columns else 0
    new_df["単価"] = pd.to_numeric(pd.Series(raw_price).apply(clean_num), errors="coerce").fillna(0).astype(int)

    raw_amt_val = df[c_amt] if c_amt in df.columns else 0
    amt_series = pd.to_numeric(pd.Series(raw_amt_val).apply(clean_num), errors="coerce").fillna(0)
    new_df["金額"] = [int(a) if a > 0 else int(q * p) for a, q, p in zip(amt_series, new_df["数量"], new_df["単価"])]

    raw_date = df[c_date].astype(str).str.replace(r'[^0-9]', '', regex=True) if c_date in df.columns else pd.Series([""] * len(df))
    dt = pd.to_datetime(raw_date, format="%Y%m%d", errors="coerce")
    new_df["日付_dt"] = dt
    new_df["日付"] = dt.dt.strftime('%Y/%m/%d').fillna("")
    new_df["年月"] = dt.dt.strftime('%Y年%m月').fillna("2026年01月")

    new_df = new_df[new_df["商品コード"] != "ABA-40"].copy()

    new_df["単価原価"] = new_df["商品コード"].map(lambda c: cost_master.get(str(c).strip(), 0.0))
    new_df["原価合計"] = (new_df["数量"] * new_df["単価原価"]).round().astype(int)
    new_df["粗利額"] = new_df["金額"] - new_df["原価合計"]

    return new_df

def render_universal_table(df_data):
    if df_data.empty:
        st.info("該当するデータがありません。")
        return
    center_cols = {"番号", "順位", "日付", "販売期間", "最終売上日", "初回売上日"}
    left_cols = {"得意先名", "商品コード", "商品名", "商品番号", "ランク名", "担当者", "客先名", "販売価格帯"}

    header_html = "".join([f"<th style='border:1px solid #cbd5e1; padding:8px 10px; background:#f1f5f9; text-align:{('center' if c in center_cols else ('left' if c in left_cols else 'right'))}; white-space:nowrap;'>{c}</th>" for c in df_data.columns])
    body_rows = []
    for r_i, row in df_data.iterrows():
        is_tot_row = str(row.get("客先名", "")) == "合計" or str(row.get("得意先名", "")) == "合計"
        bg = "#f1f5f9" if is_tot_row else ("#ffffff" if r_i % 2 == 0 else "#f8fafc")
        font_w = "font-weight: bold;" if is_tot_row else ""
        cells = []
        for c in df_data.columns:
            val = row[c]
            val_str = "" if pd.isna(val) else str(val)
            align = "center" if c in center_cols else ("left" if c in left_cols else "right")
            if isinstance(val, (int, float)) and c not in center_cols:
                val_str = f"{int(val):+,}" if c in ["増減", "増減額", "目標差額", "前年差額"] else f"{int(val):,}"

            is_neg = False
            if isinstance(val, (int, float)) and val < 0: is_neg = True
            elif isinstance(val_str, str) and (val_str.startswith("-") or val_str.startswith("▲-")): is_neg = True

            style_extra = "color: #dc2626; font-weight: bold;" if is_neg else f"color: #0f172a; {font_w}"
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

def generate_excel_bytes(data_df, title_str):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        data_df.to_excel(writer, index=False, sheet_name=title_str[:30])
    return output.getvalue()

def build_meeting_comparison_card_data(cust, cid, cur_m, ly_m, next_m, prev_fy, tgt_y, tgt_m, next_y, next_m_val, f_year):
    cur_sales = int(cur_m[cur_m["得意先名"] == cust]["金額"].sum()) if not cur_m.empty else 0
    ly_sales = int(ly_m[ly_m["得意先名"] == cust]["金額"].sum()) if not ly_m.empty else 0
    next_sales = int(next_m[next_m["得意先名"] == cust]["金額"].sum()) if not next_m.empty else 0
    prev_fy_sales = int(prev_fy[prev_fy["得意先名"] == cust]["金額"].sum()) if not prev_fy.empty else 0
    return {
        "cust_name": cust,
        "cust_id": cid,
        "cur_sales": cur_sales,
        "ly_sales": ly_sales,
        "diff": cur_sales - ly_sales,
        "next_sales": next_sales,
        "prev_fy_sales": prev_fy_sales,
        "next_diff": next_sales - prev_fy_sales,
        "tgt_m": tgt_m,
        "next_m": next_m_val
    }

def render_meeting_card_html(card):
    diff_color = "#16a34a" if card["diff"] >= 0 else "#dc2626"
    next_diff_color = "#16a34a" if card["next_diff"] >= 0 else "#dc2626"
    return f"""
    <div style="border:1px solid #cbd5e1; border-radius:6px; padding:12px; background:#fff; margin-bottom:12px;">
        <div style="font-size:15px; font-weight:bold; color:#1e3a8a; border-bottom:1px solid #e2e8f0; padding-bottom:6px; margin-bottom:8px;">
            🏢 {card['cust_name']} <span style="font-size:12px; color:#64748b;">(ID: {card['cust_id']})</span>
        </div>
        <div style="display:flex; justify-content:space-between; font-size:13px; margin-bottom:4px;">
            <span><b>{card['tgt_m']}月実績:</b> ¥{card['cur_sales']:,}</span>
            <span>(前年: ¥{card['ly_sales']:,})</span>
            <span style="color:{diff_color}; font-weight:bold;">差額: ¥{card['diff']:,}</span>
        </div>
        <div style="display:flex; justify-content:space-between; font-size:13px;">
            <span><b>来月({card['next_m']}月)予測:</b> ¥{card['next_sales']:,}</span>
            <span>(前年同月: ¥{card['prev_fy_sales']:,})</span>
            <span style="color:{next_diff_color}; font-weight:bold;">差額: ¥{card['next_diff']:,}</span>
        </div>
    </div>
    """

def generate_meeting_comparison_excel(cards_data, tgt_ym):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "上位社売上比較"
    ws.cell(row=1, column=1, value=f"{tgt_ym} 営業会議 上位社売上比較シート").font = Font(size=14, bold=True)
    row_idx = 3
    for c in cards_data:
        ws.cell(row=row_idx, column=1, value=f"得意先名: {c['cust_name']}")
        ws.cell(row=row_idx+1, column=1, value=f"当月売上: {c['cur_sales']:,} 円")
        ws.cell(row=row_idx+2, column=1, value=f"来月予測: {c['next_sales']:,} 円")
        row_idx += 5
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()

def generate_meeting_top_cust_excel(res_df, title_str):
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine='openpyxl') as writer:
        res_df.to_excel(writer, index=False, sheet_name="上位客先実績")
    return out.getvalue()

cost_path, cost_mtime = get_cost_master_file_info()
cost_master = load_cost_master(cost_path, cost_mtime)

raw_sup_df = load_data_from_supabase()
df = prep_supabase_dataframe(raw_sup_df, cost_master)

member_p_path, member_p_mtime = get_member_price_file_info()
member_cust_sku_map, member_default_min_map = load_member_price_master_v2(member_p_path, member_p_mtime)

jan_path, jan_mtime = get_jan_file_info()
jan_master_dict, jan_title_dict, jan_std_price_dict = load_jan_master_v2(jan_path, jan_mtime)

cust_path, cust_mtime = get_customer_file_info()
customer_master_df = load_customer_master(cust_path, cust_mtime)

hanko_bytes, hanko_file_path = get_hanko_bytes()

col_status, col_refresh = st.columns([4, 1.2])
with col_refresh:
    if st.button("🔄 Supabaseから最新データを再読込", key="btn_reload_supabase_main"):
        st.cache_data.clear()
        st.success("最新データを取得しました！")
        time.sleep(0.3)
        st.rerun()

if df.empty:
    st.error("⚠️ Supabaseの `sales_data` テーブルからデータを取得できませんでした。Supabaseの接続情報（Secrets）やテーブル名をご確認ください。")
    st.stop()

valid_ym = sorted([x for x in df["年月"].unique() if x and str(x).strip() != ""])
latest_ym = valid_ym[-1] if valid_ym else "2026年01月"
desc_valid_ym = sorted(valid_ym, reverse=True)

valid_dts = df["日付_dt"].dropna()
min_date = valid_dts.min().date() if not valid_dts.empty else datetime(2022, 1, 1).date()
max_date = valid_dts.max().date() if not valid_dts.empty else datetime.now().date()

curr_y = int(latest_ym[:4])
curr_m = int(latest_ym[5:7])
fiscal_year = curr_y if curr_m >= 7 else curr_y - 1

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
        st.info(f"📁 原価: **{len(cost_master):,} 件** | {member_status} | {jan_status} | {hanko_status}")
    else:
        quick_choice = st.radio("期間ショートカット", ['当月', '前月', '当期累計', '前期累計', '全データ累計', '手動指定'], horizontal=True, key="period_shortcut_radio_top_v3")

# グローバルな担当者選択（分析モード用）
staff_options_global = ["（すべて）"] + sorted([x for x in df["担当者"].unique() if x and x != "nan"])

# ========================================================
# ★ 🧮 簡易見積＆利益シミュレーター（営業担当別切り替え対応）
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

    all_calc_codes = sorted(list(set(list(code_to_name_map.keys()) + list(cost_master.keys()) + list(jan_master_dict.keys()) + list(member_default_min_map.keys()))))
    if "ABA-40" in all_calc_codes: all_calc_codes.remove("ABA-40")

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
        p_cost = int(cost_master.get(sku_code, 0.0))
        
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
            elif min_p > 0:
                p_price = min_p
            else:
                p_price = 99999

        if p_price == 0: p_price = 99999
        return p_name, p_jan, p_price, p_cost, min_p, actual_rank_applied

    if "saved_estimates" not in st.session_state:
        st.session_state.saved_estimates = load_saved_quotes()

    est_staff_list = list(STAFF_MAP.values()) + ["（すべて表示）"]
    col_menu_left, col_menu_right = st.columns([1.6, 2.0])
    with col_menu_right:
        est_staff = st.selectbox(
            "👤 営業担当者を選択（作業スペース切替 ＆ 保存済み見積の絞り込み）",
            est_staff_list,
            index=7 if "水原" in est_staff_list else 0,
            key="est_staff_selector_box_v3"
        )

    staff_Filter_list = [
        e for e in st.session_state.saved_estimates
        if est_staff == "（すべて表示）" or e.get("staff", "") == est_staff or (est_staff in str(e.get("hitokoto", "")))
    ]

    with col_menu_left:
        tab_choice = st.radio(
            "見積機能メニュー",
            ["📝 新規見積作成・シミュレーター", f"📋 保存済み見積一覧・スマレジ注文CSV出力（{len(staff_Filter_list)}件）"],
            horizontal=True,
            key="est_active_tab_v3"
        )

    s_key = est_staff
    k_items = f"quotation_items_{s_key}"
    k_ver = f"editor_version_{s_key}"
    k_comp = f"in_cust_comp_{s_key}"
    k_person = f"in_cust_person_{s_key}"
    k_post = f"in_cust_post_{s_key}"
    k_addr = f"in_cust_addr_{s_key}"
    k_s_comp = f"in_ship_comp_{s_key}"
    k_s_person = f"in_ship_person_{s_key}"
    k_s_post = f"in_ship_post_{s_key}"
    k_s_addr = f"in_ship_addr_{s_key}"
    k_s_tel = f"in_ship_tel_{s_key}"
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
        st.session_state[k_sel_c] = loaded_est.get("cust_comp", "（手動入力）")
        st.session_state[k_items] = loaded_est.get("items", [])
        st.session_state[k_ver] += 1

    if tab_choice.startswith("📝"):
        st.markdown(f"##### 🏢 提出先（得意先）の選択 &nbsp;<span style='font-size:13px; color:#2563eb;'>［作業中の担当: {est_staff}］</span>", unsafe_allow_html=True)
        cust_options = ["（手動入力）"]
        cust_lookup = {}
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
                for k_f in [k_comp, k_person, k_post, k_addr, k_s_comp, k_s_person, k_s_post, k_s_addr, k_s_tel]:
                    st.session_state[k_f] = ""
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

                st.session_state[k_comp] = selected
                st.session_state[k_person] = p_name
                st.session_state[k_post] = p_code
                st.session_state[k_addr] = full_addr
                st.session_state[k_s_comp] = selected
                st.session_state[k_s_person] = p_name
                st.session_state[k_s_post] = p_code
                st.session_state[k_s_addr] = full_addr
                st.session_state[k_s_tel] = tel

                if st.session_state[k_items]:
                    for itm in st.session_state[k_items]:
                        if itm.get("適用ランク") == "ロット単価":
                            _, _, new_p, _, new_min, _ = get_sku_details(itm["商品コード"], "ロット単価", selected)
                            itm["見積単価"] = new_p
                            itm["最低下代"] = new_min
                    st.session_state[k_ver] += 1

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
                    item["適用ランク"] = applied_lbl
                    item["見積単価"] = int(new_p)
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
                    if c == "選択": raw_items_df[c] = False
                    elif c == "適用ランク": raw_items_df[c] = "ロット単価"
                    elif c == "粗利率": raw_items_df[c] = "-"
                    else: raw_items_df[c] = 0
            raw_items_df["数量"] = raw_items_df["数量"].apply(lambda x: max(1, int(clean_num(x))))
            raw_items_df["最低下代"] = raw_items_df["最低下代"].apply(lambda x: int(clean_num(x)))
            raw_items_df["見積単価"] = raw_items_df["見積単価"].apply(lambda x: int(clean_num(x)))
            raw_items_df["単価原価"] = raw_items_df["単価原価"].apply(lambda x: int(clean_num(x)))
            raw_items_df["合計金額"] = raw_items_df["数量"] * raw_items_df["見積単価"]
            raw_items_df["粗利額"] = raw_items_df["合計金額"] - (raw_items_df["数量"] * raw_items_df["単価原価"])
            raw_items_df["粗利率"] = raw_items_df.apply(lambda r: f"{(r['粗利額'] / r['合計金額'] * 100):.1f}%" if r['合計金額'] > 0 else "-", axis=1)
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
            num_rows="dynamic",
            use_container_width=True,
            hide_index=True,
            key=f"quote_editor_{s_key}_{st.session_state[k_ver]}"
        )

        updated_records = []
        needs_rerun = False

        for idx_e, row_e in edited_df.iterrows():
            is_selected = bool(row_e.get("選択", False))
            if trigger_delete_selected and is_selected:
                needs_rerun = True
                continue

            c_code = str(row_e["商品コード"]).strip() if pd.notna(row_e["商品コード"]) else ""
            if not c_code or c_code == "None" or c_code == "ABA-40": continue

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

            if c_code != prev_code or not cur_name or cur_name == "None" or cur_price == 0:
                auto_name, auto_jan, auto_price, auto_cost, auto_min, applied_lbl = get_sku_details(c_code, cur_rank if cur_rank != "None" else "ロット単価", q_cust_comp)
                cur_name = auto_name
                cur_jan = str(auto_jan)
                cur_price = int(auto_price)
                cur_cost = int(auto_cost)
                cur_min = int(auto_min)
                cur_rank = applied_lbl
                needs_rerun = True
            elif cur_rank != prev_rank:
                _, _, rank_p, _, _, applied_lbl = get_sku_details(c_code, cur_rank, q_cust_comp)
                cur_price = int(rank_p)
                cur_rank = applied_lbl
                needs_rerun = True
            elif cur_qty != prev_qty or cur_price != prev_price:
                needs_rerun = True

            new_total = cur_qty * cur_price
            new_profit = new_total - (cur_qty * cur_cost)
            new_margin_str = f"{(new_profit / new_total * 100):.1f}%" if new_total > 0 else "-"

            updated_records.append({
                "選択": False if trigger_delete_selected else is_selected,
                "商品コード": c_code,
                "商品名": cur_name,
                "JANコード": cur_jan,
                "最低下代": cur_min,
                "適用ランク": cur_rank,
                "数量": cur_qty,
                "見積単価": cur_price,
                "合計金額": new_total,
                "単価原価": cur_cost,
                "粗利額": new_profit,
                "粗利率": new_margin_str
            })

        if trigger_delete_selected or needs_rerun or len(updated_records) != len(st.session_state[k_items]):
            st.session_state[k_items] = updated_records
            st.session_state[k_ver] += 1
            st.rerun()
        else:
            st.session_state[k_items] = updated_records

        if updated_records:
            calc_df = pd.DataFrame(updated_records)
            q_tot_qty = int(calc_df["数量"].sum())
            q_tot_unit_price = int(calc_df["見積単価"].sum())
            q_tot_sales = int(calc_df["合計金額"].sum())
            q_tot_profit = int(calc_df["粗利額"].sum())
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

            current_estimate_obj = {
                "id": len(st.session_state.saved_estimates) + 840,
                "q_no": q_no,
                "staff": est_staff,
                "created_at": datetime.now().strftime("%Y/%m/%d %H:%M"),
                "q_date": q_date.strftime("%Y/%m/%d"),
                "q_exp_date": q_exp_date.strftime("%Y/%m/%d"),
                "payment": q_payment,
                "cust_comp": q_cust_comp,
                "cust_person": q_cust_person,
                "cust_post": q_cust_post,
                "cust_addr": q_cust_addr,
                "ship_comp": ship_cust_comp,
                "ship_person": ship_cust_person,
                "ship_post": ship_cust_post,
                "ship_addr": ship_cust_addr,
                "ship_tel": ship_cust_tel,
                "shipping_fee": q_shipping_fee,
                "sub_total": q_tot_sales,
                "tax": q_tax,
                "grand_total": q_grand_total,
                "memo": q_memo,
                "hitokoto": q_staff_memo,
                "items": updated_records
            }

            hanko_base64 = base64.b64encode(hanko_bytes).decode("utf-8") if hanko_bytes else ""
            pdf_rows_html = ""
            for idx, r in calc_df.iterrows():
                p_display = f"{r['商品名']} ({r['商品コード']})" if r['商品名'] else str(r['商品コード'])
                pdf_rows_html += f"""
                <tr>
                    <td style="text-align:center;">{idx + 1}</td>
                    <td style="text-align:left;">{p_display}</td>
                    <td style="text-align:center;">{r['JANコード']}</td>
                    <td style="text-align:right;">{r['数量']:,}</td>
                    <td style="text-align:right;">¥{r['見積単価']:,}</td>
                    <td style="text-align:right;">¥{r['合計金額']:,}</td>
                </tr>
                """

            hanko_img_tag = f'<img src="data:image/png;base64,{hanko_base64}" style="position:absolute; right:15px; top:5px; width:75px; height:75px; opacity:0.92;" />' if hanko_base64 else ''

            pdf_template_html = f"""
            <!DOCTYPE html>
            <html>
            <head>
                <meta charset="utf-8">
                <title>御見積書_{q_cust_comp}</title>
                <style>
                    @media print {{
                        @page {{ size: A4 portrait; margin: 12mm; }}
                        body {{ -webkit-print-color-adjust: exact; }}
                        .no-print {{ display: none !important; }}
                    }}
                    body {{
                        font-family: -apple-system, BlinkMacSystemFont, "Yu Gothic", "Meiryo", sans-serif;
                        color: #111827; padding: 10px; max-width: 800px; margin: 0 auto;
                    }}
                    .title {{ text-align: center; font-size: 26px; font-weight: bold; letter-spacing: 12px; padding-bottom: 6px; border-bottom: 2px solid #111; margin-bottom: 20px; }}
                    .header-flex {{ display: flex; justify-content: space-between; margin-bottom: 15px; }}
                    .cust-area {{ font-size: 14px; line-height: 1.6; width: 55%; }}
                    .cust-comp {{ font-size: 17px; font-weight: bold; margin-top: 4px; }}
                    .cust-person {{ font-size: 17px; font-weight: bold; margin-top: 2px; }}
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
                </style>
            </head>
            <body>
                <div class="no-print" style="text-align:right; margin-bottom:12px;">
                    <button class="btn-print" onclick="window.print();">🖨️ PDFとして保存 / 印刷する</button>
                </div>
                <div class="title">見積書</div>
                <div class="header-flex">
                    <div class="cust-area">
                        <div>{'〒 ' + q_cust_post if q_cust_post else ''}</div>
                        <div>{q_cust_addr}</div>
                        <div class="cust-comp">{q_cust_comp}</div>
                        <div class="cust-person">{q_cust_person}</div>
                    </div>
                    <div class="company-area">
                        {hanko_img_tag}
                        <div class="company-name">スマイルキッズ株式会社</div>
                        <div>住所: 〒537-0003 大阪市東成区神路4丁目3番18号</div>
                        <div>TEL : 06-6976-1371 &nbsp; FAX : 06-6976-8940</div>
                    </div>
                </div>
                <div class="section-header">見積内容</div>
                <div class="content-box">
                    <div>■見積金額 &nbsp;&nbsp;&nbsp;&nbsp;: &nbsp;<b>{q_grand_total:,} 円</b> (税込)</div>
                    <div>■お支払い方法 : &nbsp;{q_payment}</div>
                </div>
                <table class="quote-table">
                    <thead>
                        <tr>
                            <th style="width: 35px;">No.</th>
                            <th>品名・規格</th>
                            <th style="width: 140px;">JANコード</th>
                            <th style="width: 75px;">数量</th>
                            <th style="width: 90px;">単価</th>
                            <th style="width: 105px;">金額</th>
                        </tr>
                    </thead>
                    <tbody>
                        {pdf_rows_html}
                        <tr style="background:#f9fafb; font-weight:bold; border-top:2px solid #9ca3af;">
                            <td colspan="3" style="text-align:center;">小計 / 合計</td>
                            <td style="text-align:right;">{q_tot_qty:,}</td>
                            <td style="text-align:right;">¥{q_tot_unit_price:,}</td>
                            <td style="text-align:right;">¥{q_tot_sales:,}</td>
                        </tr>
                    </tbody>
                </table>
                <table class="summary-table">
                    <tr>
                        <td style="background:#f3f4f6; text-align:center; font-weight:bold;">税抜金額合計</td>
                        <td style="text-align:right; font-weight:bold;">¥{q_tot_sales:,}</td>
                    </tr>
                    <tr>
                        <td style="background:#f3f4f6; text-align:center; font-weight:bold;">送料 (税抜)</td>
                        <td style="text-align:right; font-weight:bold;">¥{q_shipping_fee:,}</td>
                    </tr>
                    <tr>
                        <td style="background:#f3f4f6; text-align:center; font-weight:bold;">消費税 (10%)</td>
                        <td style="text-align:right; font-weight:bold;">¥{q_tax:,}</td>
                    </tr>
                    <tr style="background:#f9fafb;">
                        <td style="background:#e5e7eb; text-align:center; font-weight:bold; font-size:14px;">税込合計金額</td>
                        <td style="text-align:right; font-weight:bold; font-size:15px; color:#d93025;">¥{q_grand_total:,}</td>
                    </tr>
                </table>
            </body>
            </html>
            """

            btn_c1, _ = st.columns([1.4, 2.0])
            with btn_c1:
                if st.button("💾 この見積を保存して一覧へ登録", use_container_width=True, key=f"save_btn_{s_key}"):
                    st.session_state.saved_estimates.insert(0, current_estimate_obj)
                    save_quotes_list(st.session_state.saved_estimates)
                    st.success("✅ 保存済み見積一覧に登録しました！")

            st.write("")
            st.markdown("##### 👁️ PDFプレビュー表示＆ワンクリックPDF保存")
            components.html(pdf_template_html, height=720, scrolling=True)

    elif tab_choice.startswith("📋"):
        st.markdown(f"##### 📋 保存済み見積一覧（担当: {est_staff}）")
        if not staff_Filter_list:
            st.info(f"担当「{est_staff}」の保存された見積もりはありません。")
        else:
            for idx_q, est in enumerate(st.session_state.saved_estimates):
                if est_staff != "（すべて表示）" and est.get("staff", "") != est_staff:
                    continue
                st.write(f"見積番号: {est.get('q_no')} | 宛先: {est.get('cust_comp')} | 合計: ¥{est.get('grand_total'):,}")

    st.stop()

# ========================================================
# ★ 分析用モード（会議用資料・各種ランキング）
# ========================================================
sel_staff = "（すべて）" # デフォルト定義

if view_type == '📊 営業会議用：上位社 売上比較シート':
    col_m1, col_m2, col_m3 = st.columns([1.5, 1.5, 1.5])
    with col_m1: target_meeting_ym = st.selectbox("対象月度", desc_valid_ym, index=0, key="tgt_ym_meeting_v3")
    with col_m2: compare_cust_count = st.selectbox("比較上位社数", [4, 6, 8, 10, 20, 30], index=3, key="compare_cust_count_meeting_v3")
    with col_m3:
        staff_options_meeting = ["（すべて）"] + sorted([x for x in df["担当者"].unique() if x and x != "nan"])
        sel_staff = st.selectbox("担当者指定", staff_options_meeting, key="meeting_staff_sel_v3")

    tgt_y = int(target_meeting_ym[:4]); tgt_m = int(target_meeting_ym[5:7])
    ly_ym = f"{tgt_y - 1}年{tgt_m:02d}月"
    next_y = tgt_y + 1 if tgt_m == 12 else tgt_y
    next_m = 1 if tgt_m == 12 else tgt_m + 1
    next_ym = f"{next_y}年{next_m:02d}月"
    f_year = tgt_y if tgt_m >= 7 else tgt_y - 1
    prev_fy_ym = f"{next_y - 1}年{next_m:02d}月"

    df_scope = df if sel_staff == "（すべて）" else df[df["担当者"] == sel_staff]
    cur_m_df = df_scope[df_scope["年月"] == target_meeting_ym] if not df_scope.empty else pd.DataFrame()
    ly_m_df = df_scope[df_scope["年月"] == ly_ym] if not ly_m_df.empty else pd.DataFrame()
    next_m_df = df_scope[df_scope["年月"] == next_ym] if not next_m_df.empty else pd.DataFrame()
    prev_fy_df = df_scope[df_scope["年月"] == prev_fy_ym] if not prev_fy_df.empty else pd.DataFrame()

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

    cards_data = []
    for cust in top_custs:
        cid_s = cust_id_lookup.get(cust, "")
        cd = build_meeting_comparison_card_data(
            cust, cid_s, cur_m_df, ly_m_df, next_m_df, prev_fy_df,
            tgt_y, tgt_m, next_y, next_m, f_year
        )
        cards_data.append(cd)

    st.download_button(
        label="📥 上位社 売上比較シートをExcel保存",
        data=generate_meeting_comparison_excel(cards_data, target_meeting_ym),
        file_name=f"営業会議_上位社売上比較シート_{target_meeting_ym}_{datetime.now().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="dl_btn_meeting_comp_v3"
    )

    st.markdown(f"#### 📊 {tgt_m}月度 上位{len(top_custs)}社 売上比較シート（担当: {sel_staff}）")
    for idx_c in range(0, len(cards_data), 2):
        cols_card = st.columns(2)
        for sub_i in range(2):
            if idx_c + sub_i < len(cards_data):
                with cols_card[sub_i]:
                    st.markdown(render_meeting_card_html(cards_data[idx_c + sub_i]), unsafe_allow_html=True)

elif view_type == '📑 営業会議用：上位客先別売上実績':
    col_m1, col_m2, col_m3 = st.columns([1.5, 1.5, 1.5])
    with col_m1: target_meeting_ym = st.selectbox("対象月度", desc_valid_ym, index=0, key="tgt_ym_topcust_v3")
    with col_m2: meeting_cust_count = st.selectbox("表示上位数", [10, 20, 30, 50, 100], index=0, key="meeting_cust_count_v3")
    with col_m3:
        staff_options_topcust = ["（すべて）"] + sorted([x for x in df["担当者"].unique() if x and x != "nan"])
        sel_staff = st.selectbox("担当者指定", staff_options_topcust, key="topcust_staff_sel_v3")

    tgt_y = int(target_meeting_ym[:4]); tgt_m = int(target_meeting_ym[5:7])
    ly_ym = f"{tgt_y - 1}年{tgt_m:02d}月"
    f_year = tgt_y if tgt_m >= 7 else tgt_y - 1
    term_start_ym = f"{f_year}年07月"
    ly_term_start_ym = f"{f_year - 1}年07月"
    ly_term_end_ym = f"{tgt_y - 1}年{tgt_m:02d}月"

    df_scope = df if sel_staff == "（すべて）" else df[df["担当者"] == sel_staff]
    cur_m_df = df_scope[df_scope["年月"] == target_meeting_ym] if not df_scope.empty else pd.DataFrame()
    ly_m_df = df_scope[df_scope["年月"] == ly_ym] if not ly_m_df.empty else pd.DataFrame()
    cur_ytd_df = df_scope[(df_scope["年月"] >= term_start_ym) & (df_scope["年月"] <= target_meeting_ym)] if not df_scope.empty else pd.DataFrame()
    ly_ytd_df = df_scope[(df_scope["年月"] >= ly_term_start_ym) & (df_scope["年月"] <= ly_term_end_ym)] if not ly_ytd_df.empty else pd.DataFrame()

    cur_cust_sales = cur_m_df.groupby("得意先名")["金額"].sum() if not cur_m_df.empty else pd.Series()
    ly_cust_sales = ly_m_df.groupby("得意先名")["金額"].sum() if not ly_m_df.empty else pd.Series()
    cur_ytd_sales = cur_ytd_df.groupby("得意先名")["金額"].sum() if not cur_ytd_df.empty else pd.Series()
    ly_ytd_sales = ly_ytd_df.groupby("得意先名")["金額"].sum() if not ly_ytd_df.empty else pd.Series()

    target_cust_list = cur_cust_sales.nlargest(meeting_cust_count).index.tolist() if not cur_cust_sales.empty else []
    m_rows = []
    tot_cur_m, tot_ly_m, tot_cur_ytd, tot_ly_ytd = 0, 0, 0, 0

    for rank, cust in enumerate(target_cust_list, 1):
        s_cur = int(cur_cust_sales.get(cust, 0))
        s_ly = int(ly_cust_sales.get(cust, 0))
        ytd_cur = int(cur_ytd_sales.get(cust, 0))
        ytd_ly = int(ly_ytd_sales.get(cust, 0))
        diff = s_cur - s_ly

        tot_cur_m += s_cur
        tot_ly_m += s_ly
        tot_cur_ytd += ytd_cur
        tot_ly_ytd += ytd_ly

        m_rows.append({
            "順位": str(rank),
            "客先名": cust,
            "売上実績": s_cur,
            "前年実績": s_ly,
            "前年比": f"{(s_cur / s_ly * 100):.0f}%" if s_ly != 0 else "-",
            "増減": diff,
            "今期累計実績": ytd_cur,
            "前期累計実績": ytd_ly,
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

    st.download_button(
        label="📥 上位客先別売上実績をExcel保存",
        data=generate_meeting_top_cust_excel(res_m_df_tot, title_sheet),
        file_name=f"営業会議_{title_sheet}_{target_meeting_ym}_{datetime.now().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="dl_btn_top_cust_v3"
    )

    st.markdown(f"### {title_sheet}（担当: {sel_staff}）")
    render_universal_table(res_m_df_tot)

else:
    # 通常の分析・ランキング表示モード
    col_p1, col_p2, col_staff_sub = st.columns([1.2, 1.2, 1.6])
    with col_p1: ym_start = st.selectbox("開始年月", desc_valid_ym, index=0, key="analysis_ym_start_v3")
    with col_p2: ym_end = st.selectbox("終了年月", desc_valid_ym, index=0, key="analysis_ym_end_v3")
    with col_staff_sub:
        staff_options_sub = ["（すべて）"] + sorted([x for x in df["担当者"].unique() if x and x != "nan"])
        sel_staff_sub = st.selectbox("担当者指定", staff_options_sub, key="analysis_staff_sel_v3")

    s_ym, e_ym = (ym_end, ym_start) if ym_start > ym_end else (ym_start, ym_end)
    filtered = df[(df["年月"] >= s_ym) & (df["年月"] <= e_ym)] if not df.empty else pd.DataFrame()
    period_text = f"{s_ym} 〜 {e_ym}"

    if sel_staff_sub != "（すべて）":
        filtered = filtered[filtered["担当者"] == sel_staff_sub]

    total_sales = filtered["金額"].sum() if not filtered.empty else 0

    if view_type == '🏢 得意先別ランキング':
        agg = filtered.groupby("得意先名").agg(
            担当者=("担当者", "last"), 販売金額=("金額", "sum"), 数量=("数量", "sum"), 原価合計=("原価合計", "sum"), 粗利額=("粗利額", "sum"), 取扱品目数=("商品コード", "nunique")
        ).reset_index() if not filtered.empty else pd.DataFrame()
        
        if not agg.empty:
            agg["粗利率"] = agg.apply(lambda r: f"{(r['粗利額'] / r['販売金額'] * 100):.1f}%" if r['販売金額'] > 0 else "-", axis=1)
            agg["構成比"] = (agg["販売金額"] / total_sales * 100).round(2).map(lambda x: f"{x:.2f}%") if total_sales > 0 else "0.00%"
            agg = agg.sort_values(by="販売金額", ascending=False)
            agg.insert(0, "順位", range(1, len(agg) + 1))
        show_df = agg[["順位", "得意先名", "担当者", "販売金額", "粗利額", "粗利率", "原価合計", "数量", "取扱品目数", "構成比"]].copy() if not agg.empty else pd.DataFrame()

        st.markdown(f"<div class='total-card'><b>抽出期間:</b> {period_text} &nbsp;|&nbsp; <b>売上合計:</b> ¥{int(total_sales):,}</div>", unsafe_allow_html=True)
        render_universal_table(show_df)

    elif view_type == '📦 商品別売上ランキング':
        group_cols = [c for c in ["商品コード", "商品名"] if c in filtered.columns]
        agg = filtered.groupby(group_cols).agg(
            販売金額=("金額", "sum"), 数量=("数量", "sum"), 単価原価=("単価原価", "first"), 原価合計=("原価合計", "sum"), 粗利額=("粗利額", "sum")
        ).reset_index() if not filtered.empty else pd.DataFrame()
        
        if not agg.empty:
            agg["粗利率"] = agg.apply(lambda r: f"{(r['粗利額'] / r['販売金額'] * 100):.1f}%" if r['販売金額'] > 0 else "-", axis=1)
            agg["構成比"] = (agg["販売金額"] / total_sales * 100).round(2).map(lambda x: f"{x:.2f}%") if total_sales > 0 else "0.00%"
            agg = agg.sort_values(by="販売金額", ascending=False)
            agg.insert(0, "番号", range(1, len(agg) + 1))
        show_df = agg[["番号", "商品コード", "商品名", "数量", "単価原価", "販売金額", "原価合計", "粗利額", "粗利率", "構成比"]].copy() if not agg.empty else pd.DataFrame()

        st.markdown(f"<div class='total-card'><b>抽出期間:</b> {period_text} &nbsp;|&nbsp; <b>売上合計:</b> ¥{int(total_sales):,}</div>", unsafe_allow_html=True)
        render_universal_table(show_df)

    else:
        disp_df = filtered.drop(columns=["年月", "日付_dt"], errors="ignore").sort_values(by=["日付", "金額"], ascending=[False, False]) if not filtered.empty else pd.DataFrame()
        if not disp_df.empty:
            disp_df.insert(0, "番号", range(1, len(disp_df) + 1))
        render_universal_table(disp_df)
