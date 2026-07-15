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
    import io
    
    cols_to_del = ["广告素材数量", "素材选取 (X-Y)", "素材选取"]
    df_out = df.drop(columns=[c for c in cols_to_del if c in df.columns]).copy()
    
    target_order = [
        "广告账号ID", "主页ID", "像素ID", "真实SKU", "虚拟SKU", 
        "国家", "着陆页版本名称", "广告素材版本名称", "出价/竞价"
    ]
    ordered_cols = [c for c in target_order if c in df_out.columns]
    other_cols = [c for c in df_out.columns if c not in target_order]
    df_out = df_out[ordered_cols + other_cols]

    # ==========================================
    # ✨ 核心极简改法：在最终出口统一拼装说明行
    # ==========================================
    hints = {
        "主页ID": "可不填，不填则使用资产管理中的默认主页",
        "像素ID": "可不填，不填则使用资产管理中的默认像素",
        "真实SKU": "填了真实SKU就不能填虚拟SKU",
        "虚拟SKU": "填了虚拟SKU就不能填真实SKU",
        "国家": "美国/英国/德国/法国/西班牙",
        "着陆页版本名称": "着陆页库中的具体版本名称",
        "广告素材版本名称": "广告素材库中的具体版本名称",
        "出价/竞价": "如需指定“真实/虚拟SKU”与“出价/竞价”的关系，请填写，最多2位小数，可不填，不填则全不填，填了则全填",
        "注意事项": "此行为说明，勿删除，请从第三行开始填写"
    }
    # 动态匹配当前表格的列，生成一行字典
    hint_row = pd.DataFrame([{c: hints.get(c, "") for c in df_out.columns}])
    # 把说明行顶在最上面
    df_out = pd.concat([hint_row, df_out], ignore_index=True)
    # ==========================================

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df_out.to_excel(writer, index=False, sheet_name=sheet_name)
        
        if not params['fast_mode']:
            workbook, worksheet = writer.book, writer.sheets[sheet_name]
            
            for i, col in enumerate(df_out.columns):
                worksheet.set_column(i, i, max(len(str(col)), 15))
            
            if params['enable_color']:
                colors = ["#FFF2CC", "#E2EFDA", "#DDEBF7", "#F8CBAD", "#E4DFEC", "#D9D9D9", "#EBF1DE"]
                
                if color_by and color_by in df_out.columns:
                    target_col = color_by
                elif is_m3 and "分组" in df_out.columns:
                    target_col = "分组"
                else:
                    df_out["_color_key"] = df_out["真实SKU"].astype(str) + df_out["虚拟SKU"].astype(str)
                    target_col = "_color_key"

                current_color_idx = 0
                last_val = None
                
                for row_idx in range(len(df_out)):
                    # ✨ 给说明行单独上个“醒目”的颜色（灰底红字）
                    if row_idx == 0:
                        hint_fmt = workbook.add_format({'bg_color': '#FFFFCC', 'font_color': 'red', 'bold': True})
                        worksheet.set_row(row_idx + 1, 25, hint_fmt) # ✨ 25 是行高，显得更宽敞
                        continue
                        
                    current_val = df_out.iloc[row_idx].get(target_col, "")
                    if last_val is not None and current_val != last_val:
                        current_color_idx = (current_color_idx + 1) % len(colors)
                    
                    fmt = workbook.add_format({'bg_color': colors[current_color_idx]})
                    worksheet.set_row(row_idx + 1, None, fmt)
                    last_val = current_val

    return output.getvalue()