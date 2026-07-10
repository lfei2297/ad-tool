import pandas as pd
import re
import io

def natural_sort_key(s):
    return [int(t) if t.isdigit() else t.lower() for t in re.split('([0-9]+)', str(s))]

def expand_material_versions(row, lp_v=""):
    """
    精准解析带有多连字符的素材名称，并根据广告素材数量进行尾部数字严格递增。
    例如：'优化组版本-OPDY-S-260630-1-10', 数量 2
    生成：['优化组版本-OPDY-S-260630-1-10', '优化组版本-OPDY-S-260630-1-11']
    """
    import re
    
    # 1. 获取原始素材名称与数量
    base_name = str(row.get("广告素材版本名称", "素材")).strip()
    try:
        # 优先读取表格中的素材数量，取不到或不合法则兜底为 1
        provided_count = int(float(str(row.get("广告素材数量", 1)).strip()))
    except:
        provided_count = 1
        
    if provided_count <= 1:
        return [base_name]

    # 2. 精准匹配最后一个连字符及后面的数字
    # text: "优化组版本-OPDY-S-260630-1-10"
    # match.group(1) -> "优化组版本-OPDY-S-260630-1"
    # match.group(2) -> "10"
    match = re.match(r"^(.*)-(\d+)$", base_name)
    
    if match:
        clean_prefix = match.group(1)   # 不带末尾连字符的前缀
        start_num_str = match.group(2)  # 尾部的初始数字字符串
        start_num = int(start_num_str)  # 转为整数用于递增
        padding_len = len(start_num_str) # 保持原有的数字位数（如 01, 02 保持两位）
        
        versions = []
        for i in range(provided_count):
            current_num = start_num + i
            # 使用 zfill 保持原本的数字前导零（例如 10 变成 11，如果是 09 则变成 10）
            formatted_num = str(current_num).zfill(padding_len)
            versions.append(f"{clean_prefix}-{formatted_num}")
        return versions
    else:
        # 如果根本没有连字符加数字结尾（如直接叫 "Material"），则走常规编号形式
        return [f"{base_name}-{i}" for i in range(1, provided_count + 1)]

def write_excel_final(df, sheet_name, params, is_m3=False, color_by=None):
    """统一清洗、格式化并导出 Excel"""
    # 1. 基础清理
    cols_to_del = ["广告素材数量", "素材选取 (X-Y)", "素材选取"]
    df_out = df.drop(columns=[c for c in cols_to_del if c in df.columns]).copy()
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df_out.to_excel(writer, index=False, sheet_name=sheet_name)
        
        if not params['fast_mode']:
            workbook, worksheet = writer.book, writer.sheets[sheet_name]
            
            # 设置列宽
            for i, col in enumerate(df_out.columns):
                worksheet.set_column(i, i, max(len(str(col)), 15))
            
            if params['enable_color']:
                colors = ["#FFF2CC", "#E2EFDA", "#DDEBF7", "#F8CBAD", "#E4DFEC", "#D9D9D9", "#EBF1DE"]
                
                # --- ✨ 决定着色基准列 ---
                if color_by and color_by in df_out.columns:
                    target_col = color_by  # 模块五会走这里：按账号
                elif is_m3 and "分组" in df_out.columns:
                    target_col = "分组"    # 模块三会走这里：按分组
                else:
                    # 默认逻辑：按 SKU 组合
                    df_out["_color_key"] = df_out["真实SKU"].astype(str) + df_out["虚拟SKU"].astype(str)
                    target_col = "_color_key"

                # --- 执行着色 ---
                current_color_idx = 0
                last_val = None
                
                for row_idx in range(len(df_out)):
                    current_val = df_out.iloc[row_idx][target_col]
                    # 如果这一行的值和上一行不一样，就换颜色
                    if last_val is not None and current_val != last_val:
                        current_color_idx = (current_color_idx + 1) % len(colors)
                    
                    fmt = workbook.add_format({'bg_color': colors[current_color_idx]})
                    worksheet.set_row(row_idx + 1, None, fmt)
                    last_val = current_val

    return output.getvalue()