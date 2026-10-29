import os
import glob
import pandas as pd
from supabase import create_client, Client

# --- 設定 ---
SUPABASE_URL = "https://zyeemvbskaefbtfxurdw.supabase.co"
SUPABASE_KEY = "sb_publishable_S19HrcRBy63BGNbsSq69Xg_pZ05d3Z3"

# Gドライブのパス
BASE_SHARE = r"G:\マイドライブ\売上ダッシュボード"

def main():
    print("Supabaseに接続しています...")
    supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)

    print("CSVファイルを検索中...")
    csv_files = glob.glob(os.path.join(BASE_SHARE, "**", "*.csv"), recursive=True)
    print(f"合計 {len(csv_files)} 件のCSVファイルが見つかりました。")

    if len(csv_files) == 0:
        print("CSVファイルが見つかりませんでした。パスを確認してください。")
        return

    all_dfs = []
    for f in csv_files:
        try:
            # ヘッダーなしで読み込み
            df_temp = pd.read_csv(f, encoding="cp932", header=None, low_memory=False)
            
            # Supabaseの sales_data テーブルにある列名リスト
            expected_cols = [
                "商品コード", "商品名", "得意先名", "担当者", 
                "数量", "単価", "金額", "日付", "年月", 
                "日付_dt", "単価原価", "原価合計", "粗利額"
            ]
            
            # 列数が一致するものだけ切り取る
            num_cols = min(len(df_temp.columns), len(expected_cols))
            df_temp = df_temp.iloc[:, :num_cols]
            df_temp.columns = expected_cols[:num_cols]
            
            # --- 数値列の型変換・クリーニング ---
            numeric_cols = ["数量", "単価", "金額", "単価原価", "原価合計", "粗利額"]
            for col in numeric_cols:
                if col in df_temp.columns:
                    df_temp[col] = pd.to_numeric(
                        df_temp[col].astype(str).str.replace(',', '', regex=True), 
                        errors='coerce'
                    )

            # --- 日付・日時列の型変換・クリーニング（変な文字列が混ざっていたらNaT/Noneにする） ---
            date_cols = ["日付_dt"]
            for col in date_cols:
                if col in df_temp.columns:
                    df_temp[col] = pd.to_datetime(df_temp[col], errors='coerce')

            all_dfs.append(df_temp)
        except Exception as e:
            print(f"読込スキップ ({os.path.basename(f)}): {e}")

    if not all_dfs:
        print("有効なデータがありませんでした。")
        return

    df_combined = pd.concat(all_dfs, ignore_index=True)
    print(f"結合後の総データ行数: {len(df_combined)} 行")

    # --- 空欄（NaNやNaT）をNoneに置き換え ---
    df_combined = df_combined.astype(object).where(pd.notnull(df_combined), None)

    # データを辞書型リストに変換して一括アップロード
    print("Supabaseへのデータ送信を開始します...")
    records = df_combined.to_dict(orient="records")
    
    batch_size = 1000
    for i in range(0, len(records), batch_size):
        batch = records[i:i + batch_size]
        response = supabase.table("sales_data").insert(batch).execute()
        print(f"{i + len(batch)} / {len(records)} 行を送信しました。")

    print("すべてのデータのアップロードが完了しました！")

if __name__ == "__main__":
    main()