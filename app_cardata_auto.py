#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Mon Jul 20 16:13:50 2026

@author: jehsu
"""

import streamlit as st
import pandas as pd
import numpy as np
import io

st.set_page_config(page_title="GA4 數據自動 VLOOKUP 系統", layout="wide")

# ==============================================================================
# 核心輔助函數
# ==============================================================================
def clean_key_string(val):
    if pd.isna(val):
        return ""
    s = str(val).strip().lower().replace(" ", "").replace("_", "").replace("-", "")
    if s in ["nan", "notset", "(notset)", "null", "none"]:
        return ""
    return s

def seconds_to_mm_ss(val):
    if pd.isna(val):
        return "00:00"
    try:
        total_seconds = int(float(val))
        minutes = total_seconds // 60
        seconds = total_seconds % 60
        return f"{minutes:02d}:{seconds:02d}"
    except (ValueError, TypeError):
        return str(val)

def find_fuzzy_column(df, keyword_list):
    for col in df.columns:
        col_lower = str(col).lower().replace(" ", "").replace("_", "").strip()
        if all(kw.lower() in col_lower for kw in keyword_list):
            return col
    return None

# ==============================================================================
# Streamlit 網頁介面
# ==============================================================================
st.title("📊 GA4 跨表數據自動對齊與更新工具")
st.write("上傳最新的 GA4 總表與需要更新的車款範本，系統將自動進行比對、加總與公式計算。")

# 1. 檔案上傳區
col1, col2 = st.columns(2)

with col1:
    st.subheader("1. 上傳最新數據源（總大表）")
    new_data_file = st.file_uploader("請上傳 GA4_Merged_Report.xlsx", type=["xlsx"])

with col2:
    st.subheader("2. 上傳欲更新的範本（可多選）")
    template_files = st.file_uploader("請上傳車款範本 Excel (可一次拉入多個檔案)", type=["xlsx"], accept_multiple_files=True)

# 2. 欄位與 Key 設定（提供網頁端微調，不用改 Code）
with st.sidebar:
    st.header("⚙️ 系統欄位對應設定")
    
    st.subheader("VLOOKUP 比對 Key")
    new_keys_input = st.text_input("大表比對欄位 (逗號隔開)", "媒體,媒介,活動代號,內容")
    temp_keys_input = st.text_input("範本比對欄位 (逗號隔開)", "媒體(utm_source),媒介(utm_medium),活動代號(utm_campaign),廣告內容(utm_content)")
    
    NEW_DATA_KEYS = [k.strip() for k in new_keys_input.split(",")]
    TEMPLATE_KEYS = [k.strip() for k in temp_keys_input.split(",")]
    
    st.subheader("特定欄位強制指定")
    forced_dealer_col = st.text_input("Find Dealer 大表實際欄位名稱", "find_dealer")

# 3. 核心處理邏輯
if new_data_file and template_files:
    if st.button("🚀 開始執行數據對齊", type="primary"):
        try:
            with st.spinner("正在讀取大表並進行群組計算中..."):
                # 讀取大表
                df_new = pd.read_excel(new_data_file)
                df_new.columns = df_new.columns.astype(str).str.strip()
                
                df_new_clean = df_new.copy()
                for col in NEW_DATA_KEYS:
                    if col in df_new_clean.columns:
                        df_new_clean[col + "_clean"] = df_new_clean[col].apply(clean_key_string)
                    else:
                        st.error(f"在大表中找不到比對欄位 '{col}'，請檢查設定。")
                        st.stop()
                
                # 欄位模糊解析
                DATA_MAP_KEYWORDS = {
                    "Page Views": ["瀏覽"],  
                    "Sessions": ["工作階段"],  
                    "Active Users": ["活躍使用者"],  
                    "Bounce Rate": ["跳出率"],
                    "Time Spent per Session": ["時間"],
                    "Brochure PDF (15)": ["brochure"],  
                    "LEADS": ["名單"],                  
                }
                
                resolved_data_map = {}
                st.info("### 🔍 欄位解析狀態報告：")
                
                for temp_col, keywords in DATA_MAP_KEYWORDS.items():
                    found_col = find_fuzzy_column(df_new_clean, keywords)
                    if found_col:
                        resolved_data_map[temp_col] = found_col
                        st.success(f"✔️ 範本 **'{temp_col}'** ➡️ 成功綁定大表 『{found_col}』")
                    else:
                        st.warning(f"⚠️ 大表中找不到對應 '{temp_col}' 的欄位 (關鍵字: {keywords})")
                
                # 強制指定 find_dealer
                actual_find_dealer_col = None
                for col in df_new_clean.columns:
                    if col.strip().lower() == forced_dealer_col.strip().lower():
                        actual_find_dealer_col = col
                        break
                
                if actual_find_dealer_col:
                    resolved_data_map["Find Dealer (1)"] = actual_find_dealer_col
                    st.success(f"🎯 [強制鎖定] 範本 **'Find Dealer (1)'** ➡️ 完美綁定大表 『{actual_find_dealer_col}』")
                else:
                    st.error(f"❌ 在大表找不到你指定的 '{forced_dealer_col}' 欄位，請檢查大表欄位名稱！")
                    st.stop()
                
                # 時間與參與度
                time_col_in_new = find_fuzzy_column(df_new_clean, ["time", "spent"])
                if not time_col_in_new:
                    time_col_in_new = find_fuzzy_column(df_new_clean, ["時間"])
                
                engagement_col = find_fuzzy_column(df_new_clean, ["參與"])
                
                # 建立大表群組化計算法則
                agg_rules = {}
                for temp_col, actual_col in resolved_data_map.items():
                    agg_rules[actual_col] = "mean" if "bounce" in temp_col.lower() else "sum"
                
                if time_col_in_new and time_col_in_new not in agg_rules:
                    agg_rules[time_col_in_new] = "mean"
                if engagement_col and engagement_col not in agg_rules:
                    agg_rules[engagement_col] = "mean"
                
                # 群組化
                clean_keys_list = [col + "_clean" for col in NEW_DATA_KEYS]
                df_new_grouped = df_new_clean.groupby(clean_keys_list).agg(agg_rules).reset_index()
                
                # 建立大表聯結 Key
                df_new_grouped["_join_key_"] = (
                    df_new_grouped[NEW_DATA_KEYS[0] + "_clean"] + "_" +
                    df_new_grouped[NEW_DATA_KEYS[1] + "_clean"] + "_" +
                    df_new_grouped[NEW_DATA_KEYS[2] + "_clean"] + "_" +
                    df_new_grouped[NEW_DATA_KEYS[3] + "_clean"]
                )
                
            # 開始依序處理範本
            st.write("---")
            st.subheader("📦 處理並產生下載檔案")
            
            for t_file in template_files:
                with st.spinner(f"處理檔案中: {t_file.name}"):
                    df_temp = pd.read_excel(t_file)
                    original_cols = list(df_temp.columns)
                    df_match = df_temp.copy()
                    
                    for col in TEMPLATE_KEYS:
                        df_match[col + "_clean"] = df_match[col].apply(clean_key_string)
                    
                    df_match["_join_key_"] = (
                        df_match[TEMPLATE_KEYS[0] + "_clean"] + "_" +
                        df_match[TEMPLATE_KEYS[1] + "_clean"] + "_" +
                        df_match[TEMPLATE_KEYS[2] + "_clean"] + "_" +
                        df_match[TEMPLATE_KEYS[3] + "_clean"]
                    )
                    
                    needed_cols = ["_join_key_"] + list(agg_rules.keys())
                    df_new_subset = df_new_grouped[needed_cols]
                    
                    # Merge
                    df_merged = pd.merge(df_match, df_new_subset, on="_join_key_", how="left")
                    
                    # 覆蓋數據
                    for temp_col, actual_col in resolved_data_map.items():
                        if actual_col in df_merged.columns:
                            df_merged[temp_col] = df_merged[actual_col].fillna(0)
                    
                    # 特殊處理：秒數與分秒轉換
                    if time_col_in_new and time_col_in_new in df_merged.columns:
                        df_merged[time_col_in_new] = df_merged[time_col_in_new].fillna(0)
                        if "Time Spent per Session" in df_merged.columns:
                            df_merged["Time Spent per Session"] = df_merged[time_col_in_new]
                        if "Time Spent per Session (M)" in df_merged.columns:
                            df_merged["Time Spent per Session (M)"] = df_merged[time_col_in_new].apply(seconds_to_mm_ss)
                    
                    # 特殊處理：參與度
                    if "Non-bounce Rate" in df_merged.columns:
                        if engagement_col and engagement_col in df_merged.columns:
                            df_merged["Non-bounce Rate"] = df_merged[engagement_col].fillna(0)
                        else:
                            df_merged["Non-bounce Rate"] = 0.0
                            
                    # 特殊處理：KBA/Visit 公式
                    if "KBA/Visit" in df_merged.columns:
                        brochure_val = df_merged["Brochure PDF (15)"].fillna(0)
                        dealer_val = df_merged["Find Dealer (1)"].fillna(0)
                        leads_val = df_merged["LEADS"].fillna(0)
                        sessions_col_name = resolved_data_map.get("Sessions", "Sessions")
                        sessions_val = df_merged[sessions_col_name].fillna(0) if sessions_col_name in df_merged.columns else pd.Series(0, index=df_merged.index)
                        
                        df_merged["KBA/Visit"] = np.where(
                            sessions_val > 0,
                            (brochure_val + dealer_val + leads_val) / sessions_val,
                            0.0
                        )
                    
                    df_final = df_merged[original_cols]
                    
                    # 將結果轉成記憶體內 Excel 檔供使用者下載
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df_final.to_excel(writer, index=False)
                    excel_data = output.getvalue()
                    
                    # 顯示下載按鈕
                    st.download_button(
                        label=f"📥 下載已更新的 {t_file.name}",
                        data=excel_data,
                        file_name=f"Updated_{t_file.name}",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            
            st.balloons()
            st.success("🎉 所有檔案處理完成！請點選上方按鈕下載。")
            
        except Exception as e:
            st.error(f"執行失敗，原因：{e}")
else:
    st.info("💡 請在上方上傳最新的『數據源總表』與『車款範本』開始作業。")