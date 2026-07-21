import streamlit as st
import yfinance as yf
import pandas as pd
import requests
import os
import time
import io
from datetime import datetime

# --- 介面基本設定 (針對手機優化) ---
st.set_page_config(page_title="籌碼追蹤 APP", page_icon="📈", layout="centered")
CACHE_DIR = "cache_data"
os.makedirs(CACHE_DIR, exist_ok=True)

# --- 核心處理函數 ---
@st.cache_data(ttl=3600)
def get_trading_days(days=20):
    """取得最近 N 個台股交易日"""
    twii = yf.Ticker("^TWII")
    df = twii.history(period="2mo") 
    dates = df.index[-days:].strftime('%Y%m%d').tolist()
    return dates

def clean_twse_number(val):
    """清理證交所數字格式 (去除逗號)"""
    if pd.isna(val):
        return 0
    try:
        return float(str(val).replace(',', '').strip())
    except:
        return 0

def fetch_twse_data(date_str, type_):
    """抓取單日證交所 CSV 資料並解析"""
    if type_ == 'foreign':
        url = f"https://www.twse.com.tw/fund/TWT38U?date={date_str}&response=csv"
    else:
        url = f"https://www.twse.com.tw/fund/TWT44U?date={date_str}&response=csv"
        
    try:
        res = requests.get(url, timeout=10)
        res.raise_for_status()
        
        lines = [line for line in res.text.split('\n') if len(line.split('",')) >= 6]
        text_data = '\n'.join(lines)
        
        df = pd.read_csv(io.StringIO(text_data), header=None, on_bad_lines='skip')
        
        if df.shape[1] > 5:
            df = df[[1, 2, 5]]
            df.columns = ['Code', 'Name', 'Net']
            df['Code'] = df['Code'].astype(str).str.replace('=', '').str.replace('"', '').str.strip()
            df['Net'] = df['Net'].apply(clean_twse_number)
            return df
    except Exception as e:
        pass
        
    return pd.DataFrame(columns=['Code', 'Name', 'Net'])

def get_data_for_date(date_str, force_update=False):
    """取得資料：優先從本地 cache 讀取，沒有或強制更新才下載"""
    f_path = os.path.join(CACHE_DIR, f"{date_str}_foreign.csv")
    t_path = os.path.join(CACHE_DIR, f"{date_str}_trust.csv")

    if force_update or not (os.path.exists(f_path) and os.path.exists(t_path)):
        f_df = fetch_twse_data(date_str, 'foreign')
        t_df = fetch_twse_data(date_str, 'trust')
        
        if not f_df.empty:
            f_df.to_csv(f_path, index=False, encoding='utf-8-sig')
        if not t_df.empty:
            t_df.to_csv(t_path, index=False, encoding='utf-8-sig')
            
        time.sleep(2.5) 
        return f_df, t_df
    else:
        f_df = pd.read_csv(f_path, dtype={'Code': str})
        t_df = pd.read_csv(t_path, dtype={'Code': str})
        return f_df, t_df

# --- 主程式 UI 與邏輯 ---
st.title("📊 法人籌碼近20日追蹤")

# 1. 讀取 TW150.xlsx
try:
    # 讀取 Excel，A欄=0, B欄=1, C欄=2
    tw150_df = pd.read_excel('TW150.xlsx', header=None, dtype=str)
    target_stocks = {}
    for _, row in tw150_df.iterrows():
        code = str(row[0]).strip()
        name = str(row[1]).strip()
        # 處理 C 欄可能為空的情況，確保格式乾淨一致
        industry = str(row[2]).strip() if len(row) > 2 and pd.notna(row[2]) else "無分類"
        target_stocks[code] = {'Name': name, 'Industry': industry}
except Exception as e:
    st.error("找不到 TW150.xlsx，或檔案格式有誤。請確認 A 欄為代號，B 欄為名稱，C 欄為產業類別。")
    st.stop()

# 2. 獲取交易日並切分近 5 日
trading_days = get_trading_days(20)
recent_5_days = trading_days[-5:] # 取最後 5 天

# 3. 側邊欄設定
with st.sidebar:
    st.header("⚙️ 狀態")
    st.write(f"📅 20日範圍:\n {trading_days[0]} ~ {trading_days[-1]}")
    st.write(f"🔥 近5日範圍:\n {recent_5_days[0]} ~ {recent_5_days[-1]}")
    st.write(f"🎯 觀察清單: {len(target_stocks)} 檔")
    force_update = st.button("🔄 強制更新資料")

# 4. 初始化統計字典 (加入近5日欄位與產業)
stats = {
    code: {
        'Industry': info['Industry'], 'Name': info['Name'], 
        'f_buy_20': 0, 't_buy_20': 0, 'f_sell_20': 0, 't_sell_20': 0,
        'f_buy_5': 0, 't_buy_5': 0, 'f_sell_5': 0, 't_sell_5': 0
    } 
    for code, info in target_stocks.items()
}

progress_bar = st.progress(0, text="讀取/下載資料中，請稍候...")

# 5. 資料運算
for idx, date in enumerate(trading_days):
    f_df, t_df = get_data_for_date(date, force_update)
    is_recent_5 = date in recent_5_days # 判斷是否為近 5 日
    
    # 計算外資
    if not f_df.empty:
        f_filtered = f_df[f_df['Code'].isin(target_stocks.keys())]
        for _, row in f_filtered.iterrows():
            code = row['Code']
            if row['Net'] > 0:
                stats[code]['f_buy_20'] += 1
                if is_recent_5: stats[code]['f_buy_5'] += 1
            elif row['Net'] < 0:
                stats[code]['f_sell_20'] += 1
                if is_recent_5: stats[code]['f_sell_5'] += 1
                
    # 計算投信
    if not t_df.empty:
        t_filtered = t_df[t_df['Code'].isin(target_stocks.keys())]
        for _, row in t_filtered.iterrows():
            code = row['Code']
            if row['Net'] > 0:
                stats[code]['t_buy_20'] += 1
                if is_recent_5: stats[code]['t_buy_5'] += 1
            elif row['Net'] < 0:
                stats[code]['t_sell_20'] += 1
                if is_recent_5: stats[code]['t_sell_5'] += 1
                
    progress_bar.progress((idx + 1) / 20, text=f"正在處理: {date} ({idx+1}/20)")

progress_bar.empty()

# 6. 資料整理與分頁顯示
result_df = pd.DataFrame.from_dict(stats, orient='index')

# 優化 1：縮短 Tab 名稱，讓手機版四個 Tab 可以擠在同一行，不用下拉選單
tab1, tab2, tab3, tab4 = st.tabs(["🔥外買", "📈投買", "🩸外賣", "📉投賣"])

def format_mobile_df(df, sort_cols, main_col, sub_col, is_buy=True):
    """手機版表格格式化工具：合併天數並縮減欄位"""
    # 先做數值排序 (保留原本的 DataFrame 進行排序)
    df_sorted = df.sort_values(by=sort_cols, ascending=[False, False, False]).head(50)
    
    # 建立手機版顯示用的 DataFrame
    df_mobile = pd.DataFrame()
    df_mobile['產業'] = df_sorted['Industry'].str[:4] # 產業名稱若太長，最多截斷顯示前4個字
    df_mobile['名稱'] = df_sorted['Name']
    
    # 組合數據字串 "5日 / 20日"
    if is_buy:
        df_mobile['外資(5/20)'] = df_sorted['f_buy_5'].astype(str) + " / " + df_sorted['f_buy_20'].astype(str)
        df_mobile['投信(5/20)'] = df_sorted['t_buy_5'].astype(str) + " / " + df_sorted['t_buy_20'].astype(str)
    else:
        df_mobile['外資(5/20)'] = df_sorted['f_sell_5'].astype(str) + " / " + df_sorted['f_sell_20'].astype(str)
        df_mobile['投信(5/20)'] = df_sorted['t_sell_5'].astype(str) + " / " + df_sorted['t_sell_20'].astype(str)
        
    return df_mobile

with tab1:
    st.caption("外資買超前 50 名 (單位: 天數)")
    df_1 = format_mobile_df(
        result_df, 
        sort_cols=['f_buy_20', 'f_buy_5', 't_buy_20'], 
        main_col='外資', sub_col='投信', is_buy=True
    )
    st.dataframe(df_1, use_container_width=True, hide_index=True)

with tab2:
    st.caption("投信買超前 50 名 (單位: 天數)")
    df_2 = format_mobile_df(
        result_df, 
        sort_cols=['t_buy_20', 't_buy_5', 'f_buy_20'], 
        main_col='投信', sub_col='外資', is_buy=True
    )
    # 配合投信為主，把投信欄位移到前面
    df_2 = df_2[['產業', '名稱', '投信(5/20)', '外資(5/20)']]
    st.dataframe(df_2, use_container_width=True, hide_index=True)

with tab3:
    st.caption("外資賣超前 50 名 (單位: 天數)")
    df_3 = format_mobile_df(
        result_df, 
        sort_cols=['f_sell_20', 'f_sell_5', 't_sell_20'], 
        main_col='外資', sub_col='投信', is_buy=False
    )
    st.dataframe(df_3, use_container_width=True, hide_index=True)

with tab4:
    st.caption("投信賣超前 50 名 (單位: 天數)")
    df_4 = format_mobile_df(
        result_df, 
        sort_cols=['t_sell_20', 't_sell_5', 'f_sell_20'], 
        main_col='投信', sub_col='外資', is_buy=False
    )
    # 配合投信為主，把投信欄位移到前面
    df_4 = df_4[['產業', '名稱', '投信(5/20)', '外資(5/20)']]
    st.dataframe(df_4, use_container_width=True, hide_index=True)
