import pandas as pd
import re
import io

def natural_sort_key(s):
    return [int(t) if t.isdigit() else t.lower() for t in re.split('([0-9]+)', str(s))]

def expand_material_versions(row, lp_version):
    """通用版本解析逻辑"""
    base_name = str(row.get("广告素材版本名称", "")).strip()
    try:
        material_count = int(row["广告素材数量"]) if row["广告素材数量"] else 1
    except:
        material_count = 1
    material_select = str(row.get("素材选取", "") or row.get("素材选取 (X-Y)", "")).strip()
    sel_match = re.search(r'(\d+)-(\d+)', material_select)
    s_m, e_m = (int(sel_match.group(1)), int(sel_match.group(2))) if sel_match else (0,0)

    m_two = re.search(r'-(\d{2})$', base_name)
    m_hyphen = re.search(r'-(\d+)-(\d+)$', base_name)

    if m_two:
        suffix = m_two.group(1)
        g_n, s_mat = suffix[0], int(suffix[1])
        new_prefix = f"{base_name[:m_two.start()]}-{g_n}-{lp_version}-" if lp_version else f"{base_name[:m_two.start()]}-{g_n}-"
        start, end = (s_m, e_m) if sel_match else (s_mat, s_mat + material_count - 1)
        return [f"{new_prefix}{i}" for i in range(start, end + 1)]
    elif m_hyphen:
        g_p, s_mat = m_hyphen.group(1), int(m_hyphen.group(2))
        new_prefix = f"{base_name[:m_hyphen.start()]}-{g_p}-"
        start, end = (s_m, e_m) if sel_match else (s_mat, s_mat + material_count - 1)
        return [f"{new_prefix}{i}" for i in range(start, end + 1)]
    else:
        m_s = re.search(r'-(\d+)$', base_name)
        start_mat = int(m_s.group(1)) if m_s else 1
        prefix = base_name[:m_s.start()] + "-" if m_s else base_name + "-"
        start, end = (s_m, e_m) if sel_match else (start_mat, start_mat + material_count - 1)
        return [f"{prefix}{i}" for i in range(start, end + 1)]

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
                    
                    fmt = workbook.add_format({'bg_color': colors[current_color_idx], 'border': 1})
                    worksheet.set_row(row_idx + 1, None, fmt)
                    last_val = current_val

    return output.getvalue()