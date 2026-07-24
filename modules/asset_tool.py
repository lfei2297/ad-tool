import streamlit as st
import pandas as pd
import io
from utils import read_uploaded_excel

def find_col(df, *candidates):
    """
    智能列名容错查找函数：
    忽略大小写、空格及下划线，尝试在 df.columns 中匹配 candidates 中任意一个备选列名。
    找到则返回 df 中实际的列名，未找到返回 None。
    """
    cols_normalized = {
        c.lower().replace(" ", "").replace("_", ""): c 
        for c in df.columns
    }
    for cand in candidates:
        key = cand.lower().replace(" ", "").replace("_", "")
        if key in cols_normalized:
            return cols_normalized[key]
    return None

def run():
    st.subheader("🔗 账号-像素自动匹配工具")
    
    up_file = st.file_uploader("📂 上传资产分配表 (.xlsx)", type=["xlsx"], key="asset_up")
    
    if up_file and st.button("🚀 开始匹配资产", key="asset_btn"):
        file_bytes = up_file.getvalue()
        
        try:
            df_acc = read_uploaded_excel(file_bytes, sheet_name="账号表")
            df_pix = read_uploaded_excel(file_bytes, sheet_name="像素表")
        except Exception as e:
            st.error("⚠️ 读取失败：请确保 Excel 文件中包含名为【账号表】和【像素表】的 Sheet 页！")
            return

        if df_acc.empty or df_pix.empty:
            st.warning("⚠️ 检测到账号表或像素表为空，请检查上传文件！")
            return

        # 智能匹配两表的核心字段名
        acc_asset_col = find_col(df_acc, "资产", "资产名称", "Asset")
        pix_asset_col = find_col(df_pix, "资产", "资产名称", "Asset")
        
        acc_id_col = find_col(df_acc, "账号ID", "账号 ID", "账号id", "AccountID", "Account_ID")
        pix_id_col = find_col(df_pix, "像素ID", "像素 ID", "像素id", "PixelID", "Pixel_ID")

        if not acc_asset_col or not pix_asset_col:
            st.error("⚠️ 两表中均需包含名为【资产】的关联列，请检查表头！")
            return

        if not acc_id_col or not pix_id_col:
            st.error("⚠️ 未能在账号表中找到【账号ID】或像素表中找到【像素ID】列，请检查表头！")
            return

        # 记录原始相对位置以保持自然排序
        df_acc['_acc_order'] = range(len(df_acc))
        df_pix['_pix_order'] = range(len(df_pix))

        # 统一重命名关联键字段为“资产”，并给其他字段加上前缀区分
        df_acc_renamed = df_acc.rename(columns={
            acc_asset_col: "资产",
            acc_id_col: "账号表-账号ID"
        })
        df_pix_renamed = df_pix.rename(columns={
            pix_asset_col: "资产",
            pix_id_col: "像素表-像素ID"
        })

        # 重命名其余重名业务列
        df_acc_renamed = df_acc_renamed.rename(columns={
            c: f"账号表-{c}" if c not in ["资产", "账号表-账号ID"] and not c.startswith("_") else c 
            for c in df_acc_renamed.columns
        })
        df_pix_renamed = df_pix_renamed.rename(columns={
            c: f"像素表-{c}" if c not in ["资产", "像素表-像素ID"] and not c.startswith("_") else c 
            for c in df_pix_renamed.columns
        })

        # 关联 Merge
        combined_df = pd.merge(
            df_acc_renamed, 
            df_pix_renamed, 
            on="资产", 
            how="inner"
        )

        if combined_df.empty:
            st.warning("⚠️ 匹配结果为空，未在两表中找到名称完全相同的【资产】！")
            return

        # 按原始顺位稳定排序
        combined_df = combined_df.sort_values(by=['_acc_order', '_pix_order']).reset_index(drop=True)
        combined_df = combined_df.drop(columns=['_acc_order', '_pix_order'], errors='ignore')

        # 提取业务需要的对齐列
        sheet1_df = combined_df[['账号表-账号ID', '像素表-像素ID']].rename(columns={
            '账号表-账号ID': '账号ID',
            '像素表-像素ID': '像素ID'
        })

        out_b = io.BytesIO()
        with pd.ExcelWriter(out_b, engine="xlsxwriter") as writer:
            sheet1_df.to_excel(writer, index=False, sheet_name="账号像素对")
            combined_df.to_excel(writer, index=False, sheet_name="完整匹配详情")

        st.success("🎉 匹配成功！已按“资产”归属完美对齐。")
        st.download_button(
            "💾 下载：账号像素匹配结果 (.xlsx)", 
            data=out_b.getvalue(), 
            file_name="账号像素匹配结果.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )