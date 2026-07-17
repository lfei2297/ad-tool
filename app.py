import streamlit as st
import pandas as pd
import io
# 导入你原有的模块
from modules import module_1, module_2, module_3, module_4, module_6
# 假设你把新工具逻辑写在 modules/asset_tool.py 里
from modules import asset_tool 

st.set_page_config(page_title="广告自动化工具箱", layout="wide")

# ==========================================
# ✨ 终极视觉优化：注入极致紧凑前端样式
# ==========================================
st.markdown("""
    <style>
        /* 1. 彻底根绝顶部空白（收窄图4红框处的头部大空位） */
        .block-container {
            padding-top: 1rem !important;
            padding-bottom: 1rem !important;
            padding-left: 3rem !important;
            padding-right: 3rem !important;
        }
        
        /* 2. 收窄侧边栏顶部空白（配合图4的侧边栏头部） */
        [data-testid="stSidebarUserContent"] {
            padding-top: 1.5rem !important;
        }

        /* 3. 柔化并收窄分割线（解决图1、图2、图3红框处难看的粗横线与巨大间距） */
        hr {
            margin-top: 0.8rem !important;
            margin-bottom: 0.8rem !important;
            border-bottom: 1px solid rgba(49, 51, 63, 0.1) !important;
        }

        /* 4. 收窄常规组件（如输入框、单选框、文件上传器）之间的纵向间距 */
        [data-testid="stVerticalBlock"] > div {
            padding-bottom: 0.4rem !important;
            margin-bottom: 0px !important;
        }
        
        /* 5. 让单选框（子模块导航）的间距更加精致好点 */
        div[data-testid="stRadio"] > div {
            gap: 4px !important;
            padding: 0px !important;
        }
        
        /* 6. 稍微收窄侧边栏各个大模块之间的间距 */
        [data-testid="stSidebarV1"] .stMarkdown {
            margin-bottom: -5px !important;
        }
    </style>
""", unsafe_allow_html=True)

# 在侧边栏和正文上方都给一个清晰的标识
st.sidebar.title("🧰 工具箱控制台")
# st.title("🧰 广告自动化工具箱") # 这样页面顶部永远有一个大标题

# ==========================
# 🎯 侧边栏：核心导航升级
# ==========================
main_mode = st.sidebar.selectbox(
    "请选择业务类型", 
    ["🎨 素材批量生成", "🔧 投放资产配置"],
    index=0 # 默认选素材
)

st.sidebar.markdown("---")

def get_template_standard():
    data = {
        "广告账号ID": [""], "主页ID": ["可不填，不填则使用资产管理中的默认主页"], "像素ID": ["可不填，不填则使用资产管理中的默认像素"],
        "真实SKU": ["填了真实SKU就不能填虚拟SKU"], "虚拟SKU": ["填了虚拟SKU就不能填真实SKU"], "国家": ["美国/英国/德国/法国/西班牙"],
        "着陆页版本名称": ["着陆页库中的具体版本名称"], "广告素材版本名称": ["广告素材库中的具体版本名称"],
        "广告素材数量": ["根据需要填写"], "素材选取 (X-Y)": ["指定素材区间填写，无指定可不填"], "出价/竞价": ["如需指定“真实/虚拟SKU”与“出价/竞价”的关系，请填写，可不填，最多2位小数，不填则全不填，填了则全填"],
        "注意事项": ["此行为说明，勿删除，请从第三行开始填写"]
    }
    df = pd.DataFrame(data)
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="标准模板")
        
        workbook = writer.book
        worksheet = writer.sheets["标准模板"]
        
        worksheet.protect() 
        
        unlocked_fmt = workbook.add_format({'locked': False}) 
        header_fmt = workbook.add_format({'bold': True, 'bg_color': '#E0E0E0', 'border': 1, 'locked': True})
        
        # ✨ 关键修复1：在这里给模板的红字加上自动换行（text_wrap: True）和垂直居中（align: vcenter）
        warning_fmt = workbook.add_format({
            'bg_color': '#FFFFCC', 
            'font_color': 'red', 
            'bold': True, 
            'locked': True,
        })
        
        worksheet.set_column(0, 100, 18, unlocked_fmt)
        
        for i, col_name in enumerate(df.columns):
            # ✨ 关键修复2：把动态列宽算法同步到模板生成器里
            header_width = len(str(col_name).encode('gbk', errors='ignore'))
            
            # 靶向收窄这四列，贴合表头宽度
            if col_name in ["真实SKU", "虚拟SKU", "国家", "出价/竞价"]:
                col_width = header_width + 2
            elif col_name in ["注意事项", "着陆页版本名称", "广告素材版本名称"]:
                col_width = max(header_width + 4, 25)
            else:
                col_width = max(header_width + 4, 15)
                
            worksheet.set_column(i, i, col_width, unlocked_fmt)
            
            # 重新写入表头和说明行以应用新格式
            worksheet.write(0, i, col_name, header_fmt)
            worksheet.write(1, i, str(df.iloc[0, i]), warning_fmt)
            
        # ✨ 关键修复3：把第二行的行高撑大到 60，防止换行后的文字被遮挡
        worksheet.set_row(1, 18) 
            
    return out.getvalue()

def get_template_m4():
    columns = [
        "广告账号ID", "主页ID", "像素ID", "真实SKU", "虚拟SKU", 
        "国家", "着陆页版本名称", "广告素材版本名称", 
        "提供素材版本数量", "广告组数量", "导品系列数", "补充默认版本数", "出价/竞价",
        "注意事项"
    ]
    data = {
        "广告账号ID": [""], "主页ID": ["可不填，不填则使用资产管理中的默认主页"], "像素ID": ["可不填，不填则使用资产管理中的默认像素"],
        "真实SKU": ["填了真实SKU就不能填虚拟SKU"], "虚拟SKU": ["填了虚拟SKU就不能填真实SKU"], "国家": ["美国/英国/德国/法国/西班牙"], "着陆页版本名称": ["着陆页库中的具体版本名称"],
        "广告素材版本名称": ["广告素材库中的具体版本名称"], "提供素材版本数量": [3],
        "广告组数量": [1], "导品系列数": [1], "补充默认版本数": [1], "出价/竞价": ["如需指定“真实/虚拟SKU”与“出价/竞价”的关系，请填写，最多2位小数，可不填，不填则全不填，填了则全填"],
        "注意事项": ["此行为说明，勿删除，请从第三行开始填写"]
    }
    df = pd.DataFrame(data, columns=columns)
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="模块四模板")
        
        workbook = writer.book
        worksheet = writer.sheets["模块四模板"]
        
        worksheet.protect() 
        
        unlocked_fmt = workbook.add_format({'locked': False}) 
        header_fmt = workbook.add_format({'bold': True, 'bg_color': '#E0E0E0', 'border': 1, 'locked': True})
        warning_fmt = workbook.add_format({'bg_color': '#FFFFCC', 'font_color': 'red', 'bold': True, 'locked': True})
        
        # --- ✨ 核心改法 ---
        worksheet.set_column(0, 100, 18, unlocked_fmt)
        
        for i, col_name in enumerate(df.columns):
            col_width = max(len(str(col_name)) * 2, 18)
            worksheet.set_column(i, i, col_width, unlocked_fmt)
            
            worksheet.write(0, i, col_name, header_fmt)
            worksheet.write(1, i, str(df.iloc[0, i]), warning_fmt)
            
        worksheet.set_row(1, 18) 
        # 已删除 freeze_panes
            
    return out.getvalue()

if main_mode == "🎨 素材批量生成":
    st.title("🚀 广告素材批量生成工具")
    
    # --- 原有的下载区 ---
    st.markdown("### 📥 资源下载")
    col_t1, col_t2 = st.columns(2)
    # ... 此处保留你原有的 get_template_standard 和 get_template_m4 代码 ...
    with col_t1:
        st.download_button("⬇️ 下载：模块一/二/三/五/六 标准模板", data=get_template_standard(), file_name="标准素材模板.xlsx")
    with col_t2:
        st.download_button("⬇️ 下载：模块四 结构补齐模板", data=get_template_m4(), file_name="模块四专用模板.xlsx")
    st.markdown("---")

    # --- 原有的模块导航 ---
    st.sidebar.header("🎯 子模块导航")
    mode = st.sidebar.radio("请选择功能模块", ["模块一：基础独立拆分", "模块二：同SKU+国家聚合拆分", "模块三：智能分组 (SKU去重)", "模块四：补齐默认版本", "模块五：循环填充与分流", "模块六：多行组合循环填充"])

    st.sidebar.markdown("---")
    st.sidebar.header("⚙️ 全局设置")
    FILE_PREFIX = st.sidebar.text_input("✏️ 自定义文件前缀", value="项目A_")
    ENABLE_COLOR = st.sidebar.checkbox("开启颜色标记", value=True)
    FAST_MODE = st.sidebar.checkbox("开启极速模式(跳过样式渲染)", value=False)

    params = {"prefix": FILE_PREFIX, "enable_color": ENABLE_COLOR, "fast_mode": FAST_MODE}

    # 动态加载原有素材模块
    if mode == "模块一：基础独立拆分":
        module_1.run(params)
    elif mode == "模块二：同SKU+国家聚合拆分":
        module_2.run(params)
    elif mode == "模块三：智能分组 (SKU去重)":
        module_3.run(params)
    elif mode == "模块四：补齐默认版本":
        module_4.run(params)
    elif mode == "模块五：循环填充与分流":
        import modules.module_5 as module_5
        module_5.run(params)
    elif mode == "模块六：多行组合循环填充":
        module_6.run(params)

else:
    # ==========================
    # 🔧 业务模块：投放资产配置
    # ==========================
    st.title("🔗 广告账号-像素自动匹配工具")
    
    # --- 侧边栏：子模块导航 ---
    st.sidebar.header("🎯 子模块导航")
    # 采用单选框，为以后增加更多资产工具（如主页绑定等）预留空间
    asset_sub_mode = st.sidebar.radio(
        "请选择功能模块", 
        ["🔗 账号-像素自动匹配"]
    )
    
    st.sidebar.markdown("---")
    
    # --- 侧边栏：模块简介 ---
    st.sidebar.markdown("### 💡 模块定位")
    st.sidebar.info(
        "该模块为了快速匹配账号与像素，方便中台批量绑定。\n\n"
    )
    # --- 主界面：新工具的下载区 ---
    st.markdown("### 📥 资源下载")
    
    def get_asset_template():
        # 根据使用指南，确保模板包含必要列名 [cite: 62]
        acc_df = pd.DataFrame({"资产": ["资产A"], "账号名称": ["Acc-1"], "账号ID": ["ID001"]})
        pix_df = pd.DataFrame({"资产": ["资产A"], "像素名称": ["Pix-1"], "像素ID": ["P001"]})
        out = io.BytesIO()
        with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
            acc_df.to_excel(writer, index=False, sheet_name="账号表")
            pix_df.to_excel(writer, index=False, sheet_name="像素表")
        return out.getvalue()

    st.download_button("⬇️ 下载：账号像素匹配模板", data=get_asset_template(), file_name="账号像素匹配模板.xlsx")
    st.markdown("---")

    # 根据侧边栏的选择动态加载子模块
    if asset_sub_mode == "🔗 账号-像素自动匹配":
        # 调用封装好的资产处理逻辑 [cite: 61, 62]
        asset_tool.run()