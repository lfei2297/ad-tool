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
    asset_sig = up_file.name if up_file else None
    if st.session_state.get("asset_sig") != asset_sig:
        st.session_state["asset_sig"] = asset_sig
        st.session_state.pop("asset_xlsx", None)
        st.session_state.pop("asset_error", None)

    start_btn = st.button("🚀 开始匹配资产", key="asset_btn", disabled=not bool(up_file))
    if start_btn and up_file:
        file_bytes = up_file.getvalue()
        try:
            df_acc = read_uploaded_excel(file_bytes, sheet_name="账号表")
            df_pix = read_uploaded_excel(file_bytes, sheet_name="像素表")
        except Exception:
            st.session_state.pop("asset_xlsx", None)
            st.session_state["asset_error"] = "⚠️ 读取失败：请确保 Excel 文件中包含名为【账号表】和【像素表】的 Sheet 页！"
            df_acc = df_pix = None

        if df_acc is not None:
            if df_acc.empty or df_pix.empty:
                st.session_state.pop("asset_xlsx", None)
                st.session_state["asset_error"] = "⚠️ 检测到账号表或像素表为空，请检查上传文件！"
            else:
                acc_asset_col = find_col(df_acc, "资产", "资产名称", "Asset")
                pix_asset_col = find_col(df_pix, "资产", "资产名称", "Asset")
                acc_id_col = find_col(df_acc, "账号ID", "账号 ID", "账号id", "AccountID", "Account_ID")
                pix_id_col = find_col(df_pix, "像素ID", "像素 ID", "像素id", "PixelID", "Pixel_ID")

                if not acc_asset_col or not pix_asset_col:
                    st.session_state.pop("asset_xlsx", None)
                    st.session_state["asset_error"] = "⚠️ 两表中均需包含名为【资产】的关联列，请检查表头！"
                elif not acc_id_col or not pix_id_col:
                    st.session_state.pop("asset_xlsx", None)
                    st.session_state["asset_error"] = "⚠️ 未能在账号表中找到【账号ID】或像素表中找到【像素ID】列，请检查表头！"
                else:
                    df_acc["_acc_order"] = range(len(df_acc))
                    df_pix["_pix_order"] = range(len(df_pix))
                    df_acc_renamed = df_acc.rename(columns={
                        acc_asset_col: "资产",
                        acc_id_col: "账号表-账号ID",
                    })
                    df_pix_renamed = df_pix.rename(columns={
                        pix_asset_col: "资产",
                        pix_id_col: "像素表-像素ID",
                    })
                    df_acc_renamed = df_acc_renamed.rename(columns={
                        c: f"账号表-{c}" if c not in ["资产", "账号表-账号ID"] and not c.startswith("_") else c
                        for c in df_acc_renamed.columns
                    })
                    df_pix_renamed = df_pix_renamed.rename(columns={
                        c: f"像素表-{c}" if c not in ["资产", "像素表-像素ID"] and not c.startswith("_") else c
                        for c in df_pix_renamed.columns
                    })
                    combined_df = pd.merge(df_acc_renamed, df_pix_renamed, on="资产", how="inner")
                    if combined_df.empty:
                        st.session_state.pop("asset_xlsx", None)
                        st.session_state["asset_error"] = "⚠️ 匹配结果为空，未在两表中找到名称完全相同的【资产】！"
                    else:
                        combined_df = combined_df.sort_values(by=["_acc_order", "_pix_order"]).reset_index(drop=True)
                        combined_df = combined_df.drop(columns=["_acc_order", "_pix_order"], errors="ignore")
                        sheet1_df = combined_df[["账号表-账号ID", "像素表-像素ID"]].rename(columns={
                            "账号表-账号ID": "账号ID",
                            "像素表-像素ID": "像素ID",
                        })
                        out_b = io.BytesIO()
                        with pd.ExcelWriter(out_b, engine="xlsxwriter") as writer:
                            sheet1_df.to_excel(writer, index=False, sheet_name="账号像素对")
                            combined_df.to_excel(writer, index=False, sheet_name="完整匹配详情")
                        st.session_state["asset_xlsx"] = out_b.getvalue()
                        st.session_state.pop("asset_error", None)

    err = st.session_state.get("asset_error")
    if err:
        if "读取失败" in err or "关联列" in err or "未能在账号表" in err:
            st.error(err)
        else:
            st.warning(err)

    if "asset_xlsx" in st.session_state:
        st.success("🎉 匹配成功！已按“资产”归属完美对齐。")
        st.download_button(
            "💾 下载：账号像素匹配结果 (.xlsx)",
            data=st.session_state["asset_xlsx"],
            file_name="账号像素匹配结果.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="dl_asset_xlsx",
        )