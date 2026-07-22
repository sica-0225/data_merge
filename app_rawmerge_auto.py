#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Wed Jul 22 15:58:09 2026

@author: jehsu
"""

import io
import pandas as pd
import streamlit as st

# 設定網頁標題與圖示
st.set_page_config(page_title="GA4 報表合併工具", page_icon="📊", layout="centered")

st.title("📊 GA4 多來源報表合併工具")
st.write("請依序上傳對應的 Excel 檔案，系統將自動進行欄位轉換、群組加總並合併。")

# 定義 GA4 的基準比對欄位
ga4_keys = ["媒體", "媒介", "活動代號", "內容"]

# --- 1. 檔案上傳介面 ---
st.subheader("1. 上傳檔案")

col1, col2 = st.columns(2)
with col1:
    file_ga4 = st.file_uploader(
        "上傳 GA4 主要檔案 (GA4_xxxx.xlsx)", type=["xlsx"]
    )
    file_leads = st.file_uploader(
        "上傳 Leads 檔案 (leads_xxxx.xlsx)", type=["xlsx"]
    )

with col2:
    file_brochure = st.file_uploader(
        "上傳 Brochure 檔案 (brochure_xxxx.xlsx)", type=["xlsx"]
    )
    file_dealer = st.file_uploader(
        "上傳 Find Dealer 檔案 (find_dealer_xxxx.xlsx)", type=["xlsx"]
    )

# --- 2. 處理與合併邏輯 ---
if st.button("開始合併報表", type="primary"):
    # 檢查是否所有檔案都已上傳
    if not (file_ga4 and file_leads and file_brochure and file_dealer):
        st.error("請確認四個檔案皆已上傳！")
    else:
        try:
            with st.spinner("正在處理與合併資料，請稍候..."):

                # ---- 讀取與處理 GA4 ----
                df_ga4 = pd.read_excel(file_ga4)
                # 檢查基準欄位是否存在
                missing_ga4_cols = [
                    col for col in ga4_keys if col not in df_ga4.columns
                ]
                if missing_ga4_cols:
                    st.error(
                        f"GA4 主要檔案缺少以下必要欄位: {missing_ga4_cols}"
                    )
                    st.stop()

                for col in ga4_keys:
                    df_ga4[col] = df_ga4[col].astype(str).str.strip()

                # ---- 處理 leads 檔案 ----
                df_leads_raw = pd.read_excel(file_leads)
                leads_rename_dict = {
                    "utm_source": "媒體",
                    "utm_medium": "媒介",
                    "utm_campaign": "活動代號",
                    "utm_content": "內容",
                }
                # 檢查 leads 必要欄位
                missing_leads_cols = [
                    col
                    for col in leads_rename_dict.keys()
                    if col not in df_leads_raw.columns
                ]
                if missing_leads_cols:
                    st.error(f"Leads 檔案缺少以下必要欄位: {missing_leads_cols}")
                    st.stop()

                df_leads_raw = df_leads_raw.rename(columns=leads_rename_dict)

                for col in ga4_keys:
                    df_leads_raw[col] = (
                        df_leads_raw[col].astype(str).str.strip()
                    )

                df_leads_raw["lead_count"] = pd.to_numeric(
                    df_leads_raw["lead_count"], errors="coerce"
                ).fillna(0)
                df_leads_grouped = (
                    df_leads_raw.groupby(ga4_keys)["lead_count"]
                    .sum()
                    .reset_index()
                )
                df_leads_grouped = df_leads_grouped.rename(
                    columns={"lead_count": "名單"}
                )

                # ---- 處理 brochure 檔案 ----
                df_brochure_raw = pd.read_excel(file_brochure)
                brochure_rename_dict = {
                    "工作階段手動來源": "媒體",
                    "工作階段手動媒介": "媒介",
                    "工作階段手動廣告活動名稱": "活動代號",
                    "工作階段手動廣告素材": "內容",
                }
                # 檢查 brochure 必要欄位
                missing_brochure_cols = [
                    col
                    for col in brochure_rename_dict.keys()
                    if col not in df_brochure_raw.columns
                ]
                if missing_brochure_cols:
                    st.error(
                        f"Brochure 檔案缺少以下必要欄位: {missing_brochure_cols}"
                    )
                    st.stop()

                df_brochure_raw = df_brochure_raw.rename(
                    columns=brochure_rename_dict
                )

                for col in ga4_keys:
                    df_brochure_raw[col] = (
                        df_brochure_raw[col].astype(str).str.strip()
                    )

                if "brochure" in df_brochure_raw.columns:
                    df_brochure_raw["brochure"] = pd.to_numeric(
                        df_brochure_raw["brochure"], errors="coerce"
                    ).fillna(0)
                    df_brochure_grouped = (
                        df_brochure_raw.groupby(ga4_keys)["brochure"]
                        .sum()
                        .reset_index()
                    )
                else:
                    df_brochure_grouped = (
                        df_brochure_raw.groupby(ga4_keys)
                        .size()
                        .reset_index(name="brochure")
                    )

                # ---- 處理 find_dealer 檔案 ----
                df_dealer_raw = pd.read_excel(file_dealer)
                dealer_rename_dict = {
                    "工作階段手動來源": "媒體",
                    "工作階段手動媒介": "媒介",
                    "工作階段手動廣告活動名稱": "活動代號",
                    "工作階段手動廣告素材": "內容",
                }
                # 檢查 dealer 必要欄位
                missing_dealer_cols = [
                    col
                    for col in dealer_rename_dict.keys()
                    if col not in df_dealer_raw.columns
                ]
                if missing_dealer_cols:
                    st.error(
                        f"Find Dealer 檔案缺少以下必要欄位: {missing_dealer_cols}"
                    )
                    st.stop()

                df_dealer_raw = df_dealer_raw.rename(columns=dealer_rename_dict)

                for col in ga4_keys:
                    df_dealer_raw[col] = (
                        df_dealer_raw[col].astype(str).str.strip()
                    )

                if "find_dealer" in df_dealer_raw.columns:
                    df_dealer_raw["find_dealer"] = pd.to_numeric(
                        df_dealer_raw["find_dealer"], errors="coerce"
                    ).fillna(0)
                    df_dealer_grouped = (
                        df_dealer_raw.groupby(ga4_keys)["find_dealer"]
                        .sum()
                        .reset_index()
                    )
                else:
                    df_dealer_grouped = (
                        df_dealer_raw.groupby(ga4_keys)
                        .size()
                        .reset_index(name="find_dealer")
                    )

                # ---- 開始合併 (Left Join) ----
                merged_df = pd.merge(
                    df_ga4, df_leads_grouped, on=ga4_keys, how="left"
                )
                merged_df = pd.merge(
                    merged_df, df_brochure_grouped, on=ga4_keys, how="left"
                )
                merged_df = pd.merge(
                    merged_df, df_dealer_grouped, on=ga4_keys, how="left"
                )

                # 補零與型態轉換
                fill_cols = ["名單", "brochure", "find_dealer"]
                merged_df[fill_cols] = merged_df[fill_cols].fillna(0)
                for col in fill_cols:
                    merged_df[col] = merged_df[col].astype(int)

                # ---- 將結果轉為 Excel 記憶體二進制流 ----
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                    merged_df.to_excel(writer, index=False)
                buffer.seek(0)

            st.success("🎉 報表合併完成！")

            # 提供下載按鈕
            st.download_button(
                label="📥 下載合併後的報表",
                data=buffer,
                file_name="GA4_Merged_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

            # 預覽前 10 筆資料
            st.subheader("🔍 合併結果預覽 (前 10 筆)")
            st.dataframe(merged_df.head(10))

        except Exception as e:
            st.error(f"處理檔案時發生錯誤：{str(e)}")