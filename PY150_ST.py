# 6. 資料整理與分頁顯示
result_df = pd.DataFrame.from_dict(stats, orient='index')

# Tab 名稱維持縮減，適合手機操作
tab1, tab2, tab3, tab4 = st.tabs(["🔥外買", "📈投買", "🩸外賣", "📉投賣"])

def format_styled_mobile_df(df, sort_cols, main_col, sub_col, is_buy=True):
    """手機版表格格式化與上色工具"""
    df_sorted = df.sort_values(by=sort_cols, ascending=[False, False, False]).head(50)
    
    df_mobile = pd.DataFrame()
    df_mobile['產業'] = df_sorted['Industry'].str[:4]
    df_mobile['名稱'] = df_sorted['Name']
    
    if is_buy:
        df_mobile['外資(5/20)'] = df_sorted['f_buy_5'].astype(str) + " / " + df_sorted['f_buy_20'].astype(str)
        df_mobile['投信(5/20)'] = df_sorted['t_buy_5'].astype(str) + " / " + df_sorted['t_buy_20'].astype(str)
        # 加入隱藏數值欄位作為判斷依據
        df_mobile['_f_20'] = df_sorted['f_buy_20'].values
        df_mobile['_t_20'] = df_sorted['t_buy_20'].values
    else:
        df_mobile['外資(5/20)'] = df_sorted['f_sell_5'].astype(str) + " / " + df_sorted['f_sell_20'].astype(str)
        df_mobile['投信(5/20)'] = df_sorted['t_sell_5'].astype(str) + " / " + df_sorted['t_sell_20'].astype(str)
        # 加入隱藏數值欄位作為判斷依據
        df_mobile['_f_20'] = df_sorted['f_sell_20'].values
        df_mobile['_t_20'] = df_sorted['t_sell_20'].values
        
    # 調整投信為主的欄位順序
    if main_col == '投信':
        df_mobile = df_mobile[['產業', '名稱', '投信(5/20)', '外資(5/20)', '_f_20', '_t_20']]

    # 定義上色邏輯
    def highlight_row(row):
        # 初始化整列的樣式皆為空字串
        styles = [''] * len(row)
        # 找到「名稱」欄位在資料中的位置
        name_idx = row.index.get_loc('名稱')
        
        # 判斷：外資與投信的 20 日天數皆大於 9
        if row['_f_20'] > 9 and row['_t_20'] > 9:
            if is_buy:
                # 買超：底色黃色，字體強制黑色 (避免深色模式看不清)
                styles[name_idx] = 'background-color: #FFD700; color: #000000;'
            else:
                # 賣超：底色淺綠色，字體強制黑色
                styles[name_idx] = 'background-color: #90EE90; color: #000000;'
        return styles

    # 應用 Pandas Styler
    styled_df = df_mobile.style.apply(highlight_row, axis=1)
    return styled_df

# Streamlit 設定：將用來輔助計算的隱藏欄位設為 None (不在畫面上顯示)
hidden_cols_config = {
    "_f_20": None,
    "_t_20": None
}

with tab1:
    st.caption("外資買超前 50 名 (單位: 天數)")
    styled_1 = format_styled_mobile_df(
        result_df, sort_cols=['f_buy_20', 'f_buy_5', 't_buy_20'], 
        main_col='外資', sub_col='投信', is_buy=True
    )
    st.dataframe(styled_1, use_container_width=True, hide_index=True, column_config=hidden_cols_config)

with tab2:
    st.caption("投信買超前 50 名 (單位: 天數)")
    styled_2 = format_styled_mobile_df(
        result_df, sort_cols=['t_buy_20', 't_buy_5', 'f_buy_20'], 
        main_col='投信', sub_col='外資', is_buy=True
    )
    st.dataframe(styled_2, use_container_width=True, hide_index=True, column_config=hidden_cols_config)

with tab3:
    st.caption("外資賣超前 50 名 (單位: 天數)")
    styled_3 = format_styled_mobile_df(
        result_df, sort_cols=['f_sell_20', 'f_sell_5', 't_sell_20'], 
        main_col='外資', sub_col='投信', is_buy=False
    )
    st.dataframe(styled_3, use_container_width=True, hide_index=True, column_config=hidden_cols_config)

with tab4:
    st.caption("投信賣超前 50 名 (單位: 天數)")
    styled_4 = format_styled_mobile_df(
        result_df, sort_cols=['t_sell_20', 't_sell_5', 'f_sell_20'], 
        main_col='投信', sub_col='外資', is_buy=False
    )
    st.dataframe(styled_4, use_container_width=True, hide_index=True, column_config=hidden_cols_config)
