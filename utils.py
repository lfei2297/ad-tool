import pandas as pd
import io
import zipfile
import concurrent.futures
import streamlit as st
from itertools import cycle, islice

# ==========================================
# 1. 基础数据读取与通用工具函数
# ==========================================
def safe_int(val, default=0):
    """全局通用整型安全转换函数"""
    try:
        s = str(val).strip()
        return int(float(s)) if s else default
    except:
        return default

@st.cache_data(ttl=300, max_entries=3)
def read_uploaded_excel(file_bytes, sheet_name=0):
    df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name, dtype=str)
    df = df.loc[:, ~df.columns.str.contains('^Unnamed', na=False)]
    for col in df.columns:
        df[col] = df[col].astype(str).str.strip()
    df = df.replace(["", "nan", "NaN", "None"], None)
    df = df.dropna(how='all', axis=0)
    return df.fillna("")

# ==========================================
# 2. 智能素材版本展开引擎
# ==========================================
def expand_material_versions(row_dict):
    base_name = str(row_dict.get("广告素材版本名称", "素材")).strip()
    
    selection = str(row_dict.get("素材选取 (X-Y)", "")).strip()
    if not selection or selection.lower() == "nan":
        selection = str(row_dict.get("素材选取", "")).strip()

    if '-' in base_name and base_name.rsplit('-', 1)[1].isdigit():
        clean_prefix, base_start_num_str = base_name.rsplit('-', 1)
        padding_len = len(base_start_num_str)
        base_start_num = int(base_start_num_str)
        has_suffix = True
    else:
        clean_prefix = base_name
        padding_len = 0
        base_start_num = 1
        has_suffix = False

    if '-' in selection and all(part.strip().isdigit() for part in selection.split('-', 1)):
        parts = selection.split('-', 1)
        start_num, end_num = int(parts[0]), int(parts[1])
        if start_num > end_num:
            start_num, end_num = end_num, start_num

        versions = []
        for i in range(start_num, end_num + 1):
            if has_suffix:
                formatted_num = str(i).zfill(padding_len)
                versions.append(f"{clean_prefix}-{formatted_num}")
            else:
                versions.append(f"{clean_prefix}-{i}")
        return versions

    if selection.isdigit():
        target_num = int(selection)
        if has_suffix:
            return [f"{clean_prefix}-{str(target_num).zfill(padding_len)}"]
        else:
            return [f"{clean_prefix}-{target_num}"]

    provided_count = safe_int(row_dict.get("广告素材数量", 1), default=1)
    if provided_count <= 1:
        return [base_name]

    versions = []
    if has_suffix:
        for i in range(provided_count):
            current_num = base_start_num + i
            formatted_num = str(current_num).zfill(padding_len)
            versions.append(f"{clean_prefix}-{formatted_num}")
        return versions
    else:
        return [f"{base_name}-{i}" for i in range(1, provided_count + 1)]

# ==========================================
# 3. 辅助循环工具函数
# ==========================================
def cycle_repeat(items, target_len):
    if not items or target_len <= 0:
        return []
    repeated = list(islice(cycle(items), target_len))
    if repeated and isinstance(repeated[0], dict):
        return [item.copy() for item in repeated]
    return repeated

# ==========================================
# 4. 终极 Excel 渲染引擎
# ==========================================
def write_excel_final(df, sheet_name, params, is_m3=False, color_by=None):
    """
    【生成结果文件专用】：彻底取消保护锁，确保所有处理后的结果文件 100% 完全可自由编辑！
    """
    df = df.rename(columns={"真实sku": "真实SKU", "虚拟sku": "虚拟SKU"})
    if "真实SKU" not in df.columns: df["真实SKU"] = ""
    if "虚拟SKU" not in df.columns: df["虚拟SKU"] = ""
        
    cols_to_del = ["广告素材数量", "素材选取 (X-Y)", "素材选取"]
    df_out = df.drop(columns=cols_to_del, errors='ignore').copy()
    
    target_order = [
        "广告账号ID", "主页ID", "像素ID", "真实SKU", "虚拟SKU", 
        "国家", "着陆页版本名称", "广告素材版本名称", "出价/竞价", "系列标注", "注意事项"
    ]
    ordered_cols = [c for c in target_order if c in df_out.columns]
    other_cols = [c for c in df_out.columns if c not in target_order]
    df_out = df_out[ordered_cols + other_cols]

    hint_df = params.get('dynamic_hint')
    if hint_df is not None and not hint_df.empty:
        hint_dict = {c: str(hint_df.iloc[0].get(c, "")) for c in df_out.columns}
        hint_row = pd.DataFrame([hint_dict])
    else:
        hints = {
            "主页ID": "可不填，不填则使用资产管理中的默认主页",
            "像素ID": "可不填，不填则使用资产管理中的默认像素",
            "真实SKU": "填了真实SKU就不能填虚拟SKU",
            "虚拟SKU": "填了虚拟SKU就不能填真实SKU",
            "国家": "美国/英国/德国/法国/西班牙",
            "着陆页版本名称": "着陆页库中的具体版本名称",
            "广告素材版本名称": "广告素材库中的具体版本名称",
            "出价/竞价": "如需指定“真实/虚拟SKU”与“出价/竞价”的关系，请填写，最多2位小数，可不填，不填则全不填，填了则全填",
            "系列标注": "可不填",
            "注意事项": "此行为说明，勿删除，请从第三行开始填写"
        }
        hint_row = pd.DataFrame([{c: hints.get(c, "") for c in df_out.columns}])
        
    df_out = pd.concat([hint_row, df_out], ignore_index=True)

    with io.BytesIO() as output:
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            df_out.to_excel(writer, index=False, sheet_name=sheet_name)
            workbook, worksheet = writer.book, writer.sheets[sheet_name]
            
            num_rows, num_cols = df_out.shape
            
            # 自动筛选，但不开启 protect 锁
            worksheet.autofilter(0, 0, 0, num_cols - 1)

            if not params.get('fast_mode', False):
                # 列宽高精度计算
                for i, col in enumerate(df_out.columns):
                    gbk_len = len(str(col).encode('gbk', errors='ignore'))
                    worksheet.set_column(i, i, max(gbk_len + 4, 15))
                
                if params.get('enable_color', True):
                    colors = ["#FFF2CC", "#E2EFDA", "#DDEBF7", "#F8CBAD", "#E4DFEC", "#D9D9D9", "#EBF1DE"]
                    color_fmts = [workbook.add_format({'bg_color': c, 'locked': False}) for c in colors]
                    hint_fmt = workbook.add_format({'bg_color': '#FFFFCC', 'font_color': 'red', 'bold': True, 'locked': False})
                    
                    if color_by and color_by in df_out.columns:
                        target_series = df_out[color_by]
                    elif is_m3 and "分组" in df_out.columns:
                        target_series = df_out["分组"]
                    else:
                        target_series = df_out["真实SKU"].astype(str).str.cat(df_out["虚拟SKU"].astype(str), sep="_")

                    color_keys = target_series.tolist()
                    worksheet.set_row(1, 25, hint_fmt)
                    
                    current_color_idx = 0
                    last_val = None
                    
                    for idx in range(1, len(color_keys)):
                        current_val = color_keys[idx]
                        if last_val is not None and current_val != last_val:
                            current_color_idx = (current_color_idx + 1) % len(colors)
                        worksheet.set_row(idx + 1, None, color_fmts[current_color_idx])
                        last_val = current_val

        excel_bytes = output.getvalue()
    del df_out
    return excel_bytes

# ==========================================
# 5. 通用多线程 ZIP 打包器
# ==========================================
def create_zip_package(tasks_dict, params, total_excel_bytes=None):
    def process_task(task_item):
        n, d = task_item
        if n == "总表" and total_excel_bytes is not None: 
            return n, total_excel_bytes
        return n, write_excel_final(d, "Data", params)

    zip_b = io.BytesIO()
    with zipfile.ZipFile(zip_b, "w", compression=zipfile.ZIP_STORED) as zf:
        with concurrent.futures.ThreadPoolExecutor() as executor:
            results = executor.map(process_task, tasks_dict.items())
            for n, file_bytes in results:
                zf.writestr(f"{params.get('prefix', '项目_')}{n}.xlsx", file_bytes)
    return zip_b.getvalue()

# ==========================================
# 6. 通用标准模板生成器
# ==========================================
def build_template(columns_dict, sheet_name):
    df = pd.DataFrame([columns_dict])
    out = io.BytesIO()
    
    with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        workbook = writer.book
        worksheet = writer.sheets[sheet_name]
        
        worksheet.autofilter(0, 0, 0, len(columns_dict) - 1)
        worksheet.protect('', {'autofilter': True, 'sort': True, 'format_columns': True})
        
        unlocked_fmt = workbook.add_format({'locked': False})
        header_fmt = workbook.add_format({'bold': True, 'bg_color': '#E0E0E0', 'border': 1, 'locked': True})
        warning_fmt = workbook.add_format({'bg_color': '#FFFFCC', 'font_color': 'red', 'bold': True, 'locked': True})
        
        worksheet.set_column(0, 100, 18, unlocked_fmt)
        worksheet.set_row(1, 18)
        
        for i, (col_name, hint_text) in enumerate(columns_dict.items()):
            header_width = len(str(col_name).encode('gbk', errors='ignore'))
            if col_name in ["真实SKU", "虚拟SKU", "国家", "出价/竞价", "系列标注"]:
                col_width = header_width + 2
            elif col_name in ["注意事项", "着陆页版本名称", "广告素材版本名称"]:
                col_width = max(header_width + 4, 25)
            else:
                col_width = max(header_width + 4, 15)
                
            worksheet.set_column(i, i, col_width, unlocked_fmt)
            worksheet.write(0, i, col_name, header_fmt)
            worksheet.write(1, i, str(hint_text), warning_fmt)
            
    return out.getvalue()

def get_template_standard():
    return build_template({
        "广告账号ID": "", 
        "主页ID": "可不填，不填则使用资产管理中的默认主页", 
        "像素ID": "可不填，不填则使用资产管理中的默认像素", 
        "真实SKU": "填了真实SKU就不能填虚拟SKU", 
        "虚拟SKU": "填了虚拟SKU就不能填真实SKU", 
        "国家": "美国/英国/德国/法国/西班牙", 
        "着陆页版本名称": "着陆页库中的具体版本名称", 
        "广告素材版本名称": "广告素材库中的具体版本名称", 
        "广告素材数量": "根据需要填写", 
        "素材选取 (X-Y)": "指定素材区间填写，无指定可不填", 
        "出价/竞价": "如需指定“真实/虚拟SKU”与“出价/竞价”的关系，请填写，最多2位小数，可不填，不填则全不填，填了则全填",
        "系列标注": "可不填",
        "注意事项": "此行为说明，勿删除，请从第三行开始填写"
    }, sheet_name="标准模板")

def get_template_m4():
    return build_template({
        "广告账号ID": "", 
        "主页ID": "可不填", 
        "像素ID": "可不填", 
        "真实SKU": "", 
        "虚拟SKU": "", 
        "国家": "", 
        "着陆页版本名称": "", 
        "广告素材版本名称": "", 
        "提供素材版本数量": "", 
        "广告组数量": "填实际组数 (默认1)",         # ✨ 补齐：广告组数量
        "导品系列数": "填系列数 (默认1)", 
        "补充默认版本数": "校验用，可填0", 
        "出价/竞价": "",
        "系列标注": "可不填",
        "注意事项": "此行为说明，勿删除，请从第三行开始填写"
    }, sheet_name="模块四模板")
    
def get_template_m7():
    """
    构建模块七专用模板（包含账号表与SKU表两个Sheet，取消第二行说明行，全表可自由编辑）
    """
    acc_cols = ["资产", "账号ID", "主页ID", "像素ID", "品类", "需要组合的SKU数量", "系列标注"]
    sku_cols = ["真实SKU", "虚拟SKU", "国家", "着陆页版本名称", "广告素材版本名称", "商品分类", "出价/竞价"]

    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
        # Sheet1: 账号表
        df_acc = pd.DataFrame(columns=acc_cols)
        df_acc.to_excel(writer, index=False, sheet_name="账号表")
        ws_acc = writer.sheets["账号表"]
        
        # Sheet2: SKU表
        df_sku = pd.DataFrame(columns=sku_cols)
        df_sku.to_excel(writer, index=False, sheet_name="SKU表")
        ws_sku = writer.sheets["SKU表"]

        workbook = writer.book
        header_fmt = workbook.add_format({
            'bold': True, 
            'bg_color': '#E0E0E0', 
            'border': 1
        })

        for ws, cols in [(ws_acc, acc_cols), (ws_sku, sku_cols)]:
            # 仅在第 1 行（表头）开启自动筛选，不加 protect() 锁，全表自由编辑
            ws.autofilter(0, 0, 0, len(cols) - 1)
            
            # 设置漂亮的默认列宽与表头样式
            for i, col_name in enumerate(cols):
                w = max(len(str(col_name).encode('gbk', errors='ignore')) + 6, 18)
                ws.set_column(i, i, w)
                ws.write(0, i, col_name, header_fmt)

    return out.getvalue()