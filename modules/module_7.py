import streamlit as st
import pandas as pd
import gc
from collections import defaultdict
from utils import write_excel_final, read_uploaded_excel, safe_int

def run(params):
    st.subheader("📦 模块七：账号品类匹配与智能SKU组合")

    with st.expander("💡 点击查看：运行逻辑说明", expanded=False):
        st.markdown("""
        - **模式一 (不补齐)**：仅抽取独立 SKU，匹配多少行生成多少行，【备注】列标注实际匹配素材数。
        - **模式二 (变体最大化补齐)**：独立 SKU 不足时，自动组合该 SKU 的其他素材/着陆页版本凑数；若所有可用组合消耗完毕仍不满足，则导出全部有效组合并提示缺口。
        """)

    # 1. 设置与上传区域（左右分列）
    st.markdown("##### ⚙️ 1. 匹配模式与文件上传")
    c_left, c_right = st.columns([1, 1], gap="large")

    with c_left:
        run_mode = st.radio(
            "选择补齐与填充策略", 
            ["模式一：有多少是多少 (不补齐)", "模式二：不足补齐 (优先配不同素材)"],
            help="模式一仅使用独立的SKU；模式二会在独立SKU不够时，尝试抽取相同SKU的其他素材版本/着陆页"
        )

    with c_right:
        up = st.file_uploader("上传模块七专用模板 (.xlsx)", type=["xlsx"], key="m7_up")

    # 2. 核心操作区域
    if up:
        st.write("")
        st.markdown("##### 🚀 2. 执行匹配与结果导出")
        btn_col1, btn_col2 = st.columns([1, 1], gap="large")
        
        with btn_col1:
            start_btn = st.button("🚀 开始匹配并生成", key="m7_btn", use_container_width=True)

        if start_btn:
            file_bytes = up.getvalue()

            try:
                df_acc_raw = read_uploaded_excel(file_bytes, sheet_name="账号表")
                df_sku_raw = read_uploaded_excel(file_bytes, sheet_name="SKU表")
            except Exception:
                st.error("⚠️ 读取失败：请确保 Excel 中包含【账号表】和【SKU表】！")
                return

            acc_rows = [r for r in df_acc_raw.to_dict('records') if any(str(v).strip() for v in r.values())]
            sku_rows = [r for r in df_sku_raw.to_dict('records') if any(str(v).strip() for v in r.values())]

            if not acc_rows or not sku_rows:
                st.warning("⚠️ 检测到账号表或SKU表中无有效数据！")
                return

            # 提取原 SKU 表的全部列名顺序，方便最终对齐列头
            original_sku_columns = list(df_sku_raw.columns)

            category_sku_map = defaultdict(list)
            for r in sku_rows:
                cat = str(r.get("商品分类", "")).strip()
                if cat: category_sku_map[cat].append(r)

            final_rows = []
            warning_logs = []
            gap_summary_list = []

            for idx, acc_dict in enumerate(acc_rows):
                excel_line = idx + 2
                acc_id = str(acc_dict.get("账号ID", "")).strip()
                category = str(acc_dict.get("品类", "")).strip()
                target_sku_count = safe_int(acc_dict.get("需要组合的SKU数量"), default=1)
                zhuye = str(acc_dict.get("主页ID", "")).strip()
                xiangsu = str(acc_dict.get("像素ID", "")).strip()

                if not category:
                    warning_logs.append(f"⚠️ 行 {excel_line} [账号:{acc_id}] 未填写品类，跳过。")
                    continue

                raw_pool = category_sku_map.get(category, [])
                if not raw_pool:
                    warning_logs.append(f"❓ 行 {excel_line} [账号:{acc_id}] 品类【{category}】未找到对应SKU。")
                    gap_summary_list.append({"广告账号ID": acc_id, "品类": category, "目标数量": target_sku_count, "实际获取": 0, "缺口数量": target_sku_count})
                    continue

                virt_skus_dict, real_skus_dict = defaultdict(list), defaultdict(list)
                virt_unique_codes, real_unique_codes = [], []

                for item in raw_pool:
                    v_code = str(item.get("虚拟SKU", "")).strip()
                    r_code = str(item.get("真实SKU", "")).strip()
                    if v_code and v_code.lower() != "nan":
                        if v_code not in virt_skus_dict: virt_unique_codes.append(v_code)
                        virt_skus_dict[v_code].append(item)
                    elif r_code and r_code.lower() != "nan":
                        if r_code not in real_skus_dict: real_unique_codes.append(r_code)
                        real_skus_dict[r_code].append(item)

                ordered_unique_codes = virt_unique_codes + real_unique_codes
                total_unique_available = len(ordered_unique_codes)

                if total_unique_available == 0:
                    warning_logs.append(f"❓ 行 {excel_line} [账号:{acc_id}] 品类【{category}】无有效SKU。")
                    continue

                selected_items = []

                if "模式一" in run_mode:
                    take_codes = ordered_unique_codes[:target_sku_count]
                    for code in take_codes:
                        if code in virt_skus_dict:
                            item = virt_skus_dict[code][0].copy()
                            item["虚拟SKU"], item["真实SKU"] = code, ""
                        else:
                            item = real_skus_dict[code][0].copy()
                            item["真实SKU"], item["虚拟SKU"] = code, ""
                        selected_items.append(item)
                else:
                    for code in ordered_unique_codes:
                        if len(selected_items) >= target_sku_count: break
                        if code in virt_skus_dict:
                            item = virt_skus_dict[code][0].copy()
                            item["虚拟SKU"], item["真实SKU"] = code, ""
                        else:
                            item = real_skus_dict[code][0].copy()
                            item["真实SKU"], item["虚拟SKU"] = code, ""
                        selected_items.append(item)

                    if len(selected_items) < target_sku_count:
                        extra_pool = []
                        for code in ordered_unique_codes:
                            items = virt_skus_dict[code] if code in virt_skus_dict else real_skus_dict[code]
                            if len(items) > 1:
                                for extra_item in items[1:]:
                                    copy_item = extra_item.copy()
                                    if code in virt_skus_dict:
                                        copy_item["虚拟SKU"], copy_item["真实SKU"] = code, ""
                                    else:
                                        copy_item["真实SKU"], copy_item["虚拟SKU"] = code, ""
                                    extra_pool.append(copy_item)

                        for extra_item in extra_pool:
                            if len(selected_items) >= target_sku_count: break
                            selected_items.append(extra_item)

                actual_count = len(selected_items)
                gap = target_sku_count - actual_count

                if gap == 0:
                    note_text = "完全匹配"
                else:
                    note_text = f"⚠️缺{gap}个 ({actual_count}/{target_sku_count})"
                    gap_summary_list.append({"广告账号ID": acc_id, "品类": category, "目标数量": target_sku_count, "实际获取": actual_count, "缺口数量": gap})

                # 保留完整的字典列（包含出价/竞价等扩展列），只更新账号属性和备注
                for item in selected_items:
                    item["广告账号ID"] = acc_id
                    item["主页ID"] = zhuye
                    item["像素ID"] = xiangsu
                    item["备注"] = note_text
                    final_rows.append(item)

            # 导出按钮
            with btn_col2:
                if final_rows:
                    final_df = pd.DataFrame(final_rows)

                    # 调整列显示顺序：优先按照 SKU 表原列名排布，确保出价/竞价等字段紧跟在后面
                    preferred_order = ["广告账号ID", "主页ID", "像素ID"]
                    remaining_cols = [c for c in original_sku_columns if c not in preferred_order and c in final_df.columns]
                    
                    # 重新组装完整的列顺序
                    all_cols = preferred_order + remaining_cols
                    if "备注" not in all_cols and "备注" in final_df.columns:
                        all_cols.append("备注")
                    
                    # 补充其他可能存在的列
                    other_cols = [c for c in final_df.columns if c not in all_cols]
                    final_cols = all_cols + other_cols

                    final_df = final_df.reindex(columns=final_cols)

                    file_prefix = params.get("prefix", "项目_")
                    xlsx_data = write_excel_final(final_df, "品类匹配结果", params, color_by="广告账号ID")

                    st.download_button(
                        f"💾 下载：{file_prefix}结果 ({len(final_rows)}行)", 
                        data=xlsx_data, 
                        file_name=f"{file_prefix}模块七组合结果.xlsx",
                        use_container_width=True
                    )

            # 3. 缺口与成功提示顺延平铺在下方
            st.write("")
            if gap_summary_list:
                st.warning(f"⚠️ 检测到 **{len(gap_summary_list)}** 个账号存在素材缺口，详情请在导出表格的【备注】列核对！")
                with st.expander("🔍 点击展开查看缺口明细概要"):
                    for gap_item in gap_summary_list:
                        st.caption(f"• **{gap_item['广告账号ID']}** [{gap_item['品类']}]：目标 {gap_item['目标数量']} ➔ 实际 {gap_item['实际获取']} (缺 {gap_item['缺口数量']})")
            else:
                st.success("🎉 所有账号素材均 100% 完全匹配！")

            possible_vars = ['df_acc_raw', 'df_sku_raw', 'acc_rows', 'sku_rows', 'final_rows', 'final_df', 'gap_summary_list']
            for var in possible_vars:
                if var in locals(): del locals()[var]
            gc.collect()