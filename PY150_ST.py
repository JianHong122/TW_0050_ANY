import streamlit as st
import pandas as pd
import yfinance as yf
import re
import io
import os 
import altair as alt
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.chart import BarChart, Reference

# ==========================================
# 網頁設定與 CSS 優化
# ==========================================
st.set_page_config(page_title="台灣50中100分價量試分析", layout="wide", page_icon="📈")
hide_st_style = """
            <style>
            #MainMenu {visibility: hidden;}
            footer {visibility: hidden;}
            header {visibility: hidden;}
            </style>
            """
st.markdown(hide_st_style, unsafe_allow_html=True)

# Session State 初始化
if 'selected_stock' not in st.session_state: st.session_state.selected_stock = None
if 'analysis_results' not in st.session_state: st.session_state.analysis_results = None

# ==========================================
# 核心演算法函式
# ==========================================
def get_tick_size(price):
    if price < 10: return 0.01
    elif price < 50: return 0.05
    elif price < 100: return 0.1
    elif price < 500: return 0.5
    elif price < 1000: return 1.0
    else: return 5.0

def generate_ticks(low, high):
    low, high = round(low, 2), round(high, 2)
    if low >= high: return [low]
    ticks, curr = [], low
    while curr <= high:
        ticks.append(curr)
        curr = round(curr + get_tick_size(curr), 2)
    if ticks[-1] < high: ticks.append(high)
    return list(dict.fromkeys(ticks))

def run_analysis(df_excel):
    # 自動偵測欄位，解決 KeyError
    col_ticker = df_excel.columns[0]
    col_name = df_excel.columns[1]
    
    results, serial_num = [], 1
    total = len(df_excel)
    progress_bar = st.progress(0)
    
    for row_idx, row in df_excel.iterrows():
        progress_bar.progress((row_idx + 1) / total)
        raw_ticker = str(row[col_ticker]).strip()
        stock_name = str(row[col_name]).strip()
        if raw_ticker.endswith('.0'): raw_ticker = raw_ticker[:-2]
        
        try:
            hist = yf.Ticker(f"{raw_ticker}.TW").history(period="6mo")
            if hist.empty: hist = yf.Ticker(f"{raw_ticker}.TWO").history(period="6mo")
            if hist.empty: continue
                
            hist_64 = hist.tail(64).copy()
            curr_p = round(hist_64['Close'].iloc[-1], 2)
            max_p, min_p = hist_64['High'].max(), hist_64['Low'].min()
            if max_p == min_p: max_p, min_p = min_p * 1.05, min_p * 0.95
            
            bin_size = (max_p - min_p) / 20
            curr_idx = min(19, max(0, int((curr_p - min_p) / bin_size)))
            
            bins = [{'idx': i, 'start': min_p + i*bin_size, 'end': min_p + (i+1)*bin_size, 'vol': 0} for i in range(20)]
            
            price_vol = {}
            for _, d in hist_64.iterrows():
                vol_c, vol_o, vol_r = d['Volume']*0.3, d['Volume']*0.05, d['Volume']*0.65
                price_vol[round(d['Close'],2)] = price_vol.get(round(d['Close'],2), 0) + vol_c
                price_vol[round(d['Open'],2)] = price_vol.get(round(d['Open'],2), 0) + vol_o
                ticks = generate_ticks(d['Low'], d['High'])
                for t in ticks: price_vol[t] = price_vol.get(t, 0) + (vol_r / len(ticks))
            
            for p, v in price_vol.items():
                idx = min(19, max(0, int((p - min_p) / bin_size)))
                bins[idx]['vol'] += v
            
            top3 = sorted([b for b in bins if b['vol'] > 0], key=lambda x: x['vol'], reverse=True)[:3]
            for rank, t_bin in enumerate(top3, 1):
                if t_bin['idx'] - 3 <= curr_idx <= t_bin['idx'] + 3:
                    disp = [{'區間標籤': f"{'🎯 ' if b['idx']==curr_idx else ''}{b['start']:.2f} ~ {b['end']:.2f}", '累積成交量': int(b['vol'])} for b in bins]
                    disp.reverse()
                    results.append({'代碼': raw_ticker, '個股名稱': stock_name, '當日現價': curr_p, '落點分價': f"第{rank}大量", '分價範圍': f"{t_bin['start']:.2f} ~ {t_bin['end']:.2f}", 'bins_data': disp})
                    break
        except: continue
    progress_bar.empty()
    return results

# ==========================================
# 主介面邏輯
# ==========================================
st.title("📈 台灣50中100分價量試分析")
file_path = 'TW50100.xlsx'

if os.path.exists(file_path):
    # 讀取時自動偵測標題列，若無標題則 header=None
    df = pd.read_excel(file_path, engine='openpyxl', dtype=str)
    
    if st.session_state.analysis_results is None:
        if st.button("🚀 開始分析", type="primary"):
            st.session_state.analysis_results = run_analysis(df)
            st.rerun()
    
    if st.session_state.analysis_results:
        if st.session_state.selected_stock is None:
            st.subheader("📋 符合條件個股總覽")
            if st.button("🧹 清除並重新分析"): 
                st.session_state.analysis_results = None
                st.rerun()
            for r in st.session_state.analysis_results:
                with st.container(border=True):
                    c1, c2 = st.columns([3, 1])
                    c1.markdown(f"### {r['個股名稱']} ({r['代碼']})")
                    if c2.button("🔍 查看", key=f"btn_{r['代碼']}"):
                        st.session_state.selected_stock = r['代碼']
                        st.rerun()
                    st.caption(f"現價: {r['當日現價']} | 命中: {r['落點分價']}")
                    st.write(f"大量區: `{r['分價範圍']}`")
        else:
            target = next((r for r in st.session_state.analysis_results if r['代碼'] == st.session_state.selected_stock), None)
            if target:
                if st.button("🔙 返回總覽", type="primary"):
                    st.session_state.selected_stock = None
                    st.rerun()
                st.subheader(f"📌 {target['代碼']} {target['個股名稱']}")
                df_bins = pd.DataFrame(target['bins_data'])
                c1, c2 = st.columns([1, 2])
                with c1: st.dataframe(df_bins.set_index('區間標籤'), use_container_width=True)
                with c2:
                    chart = alt.Chart(df_bins).mark_bar().encode(
                        x='累積成交量:Q', y=alt.Y('區間標籤:N', sort=df_bins['區間標籤'].tolist()), tooltip=['區間標籤', '累積成交量']
                    ).properties(height=550)
                    st.altair_chart(chart, use_container_width=True)
else:
    st.error(f"❌ 找不到 {file_path}，請確認檔案已上傳至 GitHub 儲存庫。")
