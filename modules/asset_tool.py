import streamlit as st
import pandas as pd
# 导入之前写好的处理逻辑函数

st.title("广告账号-像素自动绑定工具")

uploaded_file = st.file_uploader("请上传原始数据 Excel (包含账号表和像素表)", type=["xlsx"])

def generate_pixel_mapping_final(input_excel, output_excel):
    try:
        # 1. 读取数据并记录原始顺序（强制 ID 为字符串）
        df_acc = pd.read_excel(input_excel, sheet_name=0, dtype={'账号ID': str})
        df_pix = pd.read_excel(input_excel, sheet_name=1, dtype={'像素ID': str})

        # 记录原始行索引，用于最后恢复顺序
        df_acc['_acc_order'] = range(len(df_acc))
        df_pix['_pix_order'] = range(len(df_pix))

        # 准备合并，重命名资产字段
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
        # 严格按照原表位置排序
        combined_df = combined_df.sort_values(by=['_acc_order', '_pix_order'])

        # 3. 准备子表字段
        sheet1 = combined_df[['账号ID', '像素ID']]
        sheet2 = combined_df[[
            '账号表-资产', '账号名称', '账号ID', 
            '像素表-资产', '像素名称', '像素ID'
        ]]

        # 4. 写入Excel并设置列宽自适应
        with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
            sheet1.to_excel(writer, sheet_name='ID绑定关系', index=False)
            sheet2.to_excel(writer, sheet_name='完整明细对照', index=False)
            
            for sheetname in writer.sheets:
                ws = writer.sheets[sheetname]
                
                # 遍历每一列设置自适应宽度
                for col in ws.columns:
                    max_length = 0
                    column_letter = col[0].column_letter
                    
                    for cell in col:
                        if cell.value:
                            # 宽度计算：中文字符按2，英文字符按1计算
                            val_str = str(cell.value)
                            length = sum(2 if ord(char) > 127 else 1 for char in val_str)
                            if length > max_length:
                                max_length = length
                    
                    # 设置列宽：最大长度 + 3个单位的缓冲余量
                    ws.column_dimensions[column_letter].width = max_length + 3
                    
                    # 再次强化设置：所有单元格设为文本格式，解决科学计数法
                    for cell in col:
                        cell.number_format = '@'

        print(f"处理完成！文件已保存且列宽已自动调整：{output_excel}")

    except Exception as e:
        print(f"处理过程中发生错误: {e}")

if uploaded_file:
    if st.button("开始处理"):
        # 调用你之前的逻辑
        output_bytes = generate_pixel_mapping_final(uploaded_file) 
        
        st.success("处理成功！")
        st.download_button(
            label="点击下载处理后的结果",
            data=output_bytes,
            file_name="像素绑定结果.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )