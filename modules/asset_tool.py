import streamlit as st
import pandas as pd
import io

def run():
    st.subheader("🎯 账号-像素自动匹配工具")
    
    st.info(
        "🔗 **账号-像素自动匹配**：上传包含账号和像素的 Excel，工具会自动按照资产进行匹配，并输出一个新的 Excel，包含账号ID与像素ID的绑定关系，以及完整的对照表。"
    )
    
    with st.expander("💡 运行逻辑说明"):
        st.write("1. **资产隔离**：账号绑定的像素必须是同资产的，防止交叉绑定。")
        st.write("2. **动态复制**：根据像素个数自动复制账号行数。")
        st.write("3. **顺序锁定**：严格按照在原始 Excel 中手动调好的账号和像素顺序排列。")
        st.write("4. **格式优化**：自动适配列宽，且 ID 类长数字不会变成科学计数法。")

    uploaded_file = st.file_uploader("📤 上传原始 Excel 数据 (需包含账号表和像素表)", type=["xlsx"])

    if uploaded_file:
        try:
            # 1. 读取数据并记录原始顺序
            # 强制将ID列设为字符串类型，防止科学计数法
            df_acc = pd.read_excel(uploaded_file, sheet_name=0, dtype={'账号ID': str})
            df_pix = pd.read_excel(uploaded_file, sheet_name=1, dtype={'像素ID': str})

            # 校验必要字段
            required_acc = ['资产', '账号ID', '账号名称']
            required_pix = ['资产', '像素ID', '像素名称']
            
            if not all(col in df_acc.columns for col in required_acc) or \
               not all(col in df_pix.columns for col in required_pix):
                st.error("❌ 上传的表格字段不全，请检查是否包含：资产、账号ID、账号名称、像素ID、像素名称")
                return

            # 记录原始行索引
            df_acc['_acc_order'] = range(len(df_acc))
            df_pix['_pix_order'] = range(len(df_pix))

            # 准备合并
            df_acc_r = df_acc.rename(columns={'资产': '账号表-资产'})
            df_pix_r = df_pix.rename(columns={'资产': '像素表-资产'})

            # 2. 合并与恢复原始顺序
            combined_df = pd.merge(
                df_acc_r, 
                df_pix_r, 
                left_on='账号表-资产', 
                right_on='像素表-资产', 
                how='inner'
            )
            
            # 排序：账号原始位置 -> 像素原始位置
            combined_df = combined_df.sort_values(by=['_acc_order', '_pix_order'])

            # 3. 准备子表
            sheet1 = combined_df[['账号ID', '像素ID']]
            sheet2 = combined_df[[
                '账号表-资产', '账号名称', '账号ID', 
                '像素表-资产', '像素名称', '像素ID'
            ]]

            # 4. 写入内存 Excel 并设置样式
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                sheet1.to_excel(writer, sheet_name='ID绑定关系', index=False)
                sheet2.to_excel(writer, sheet_name='完整明细对照', index=False)
                
                # 设置列宽自适应和文本格式
                for sheetname in writer.sheets:
                    ws = writer.sheets[sheetname]
                    for col in ws.columns:
                        max_length = 0
                        column_letter = col[0].column_letter
                        for cell in col:
                            if cell.value:
                                val_str = str(cell.value)
                                length = sum(2 if ord(char) > 127 else 1 for char in val_str)
                                if length > max_length:
                                    max_length = length
                        ws.column_dimensions[column_letter].width = max_length + 3
                        # 再次强制文本格式
                        for cell in col:
                            cell.number_format = '@'

            st.success("✅ 处理完成！")
            st.download_button(
                label="💾 点击下载匹配结果",
                data=output.getvalue(),
                file_name="账号像素匹配结果.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"⚠️ 处理出错: {e}")