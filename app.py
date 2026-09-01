import streamlit as st
import pandas as pd
import io

from utils import get_template_standard, get_template_m4, get_template_m7, get_template_m8
from modules import module_1, module_2, module_3, module_4, module_6, module_7, module_8
from modules import asset_tool 
import modules.module_5 as module_5

st.set_page_config(page_title="广告素材批量生成工具", layout="wide")

st.markdown("""
    <style>
        [data-testid="stSidebarHeader"] {
            position: absolute !important;
            top: 0.6rem !important;
            right: 0.5rem !important;
            z-index: 999999 !important;
            padding: 0 !important;
            margin: 0 !important;
            background: transparent !important;
        }
        [data-testid="stSidebarUserContent"] {
            padding-top: 0.8rem !important;
        }
        [data-testid="stSidebarUserContent"] h1 {
            padding-top: 0rem !important;
            margin-top: 0rem !important;
            font-size: 1.4rem !important;
        }
        .block-container {
            padding-top: 1.2rem !important;
            padding-bottom: 2rem !important;
            padding-left: 2.5rem !important;
            padding-right: 2.5rem !important;
        }
        h1 {
            margin-top: 0.2rem !important;
            margin-bottom: 0.6rem !important;
        }
        div[data-testid="stExpander"] {
            margin-bottom: 0.8rem !important;
        }
    </style>
""", unsafe_allow_html=True)

st.sidebar.title("🧰 工具箱控制台")

main_mode = st.sidebar.selectbox(
    "请选择业务类型", 
    ["🎨 素材批量生成", "🔧 投放资产配置"],
    index=0
)

if main_mode == "🎨 素材批量生成":
    st.sidebar.header("🎯 子模块导航")
    mode = st.sidebar.radio(
        "请选择功能模块", 
        [
            "模块一：基础独立拆分", 
            "模块二：同SKU+国家聚合拆分", 
            "模块三：智能分组 (SKU去重)", 
            "模块四：补齐默认版本", 
            "模块五：循环填充与分流", 
            "模块六：多行组合循环填充",
            "模块七：账号品类智能匹配",
            "模块八：着陆页导入与素材匹配"
        ]
    )

    st.sidebar.header("⚙️ 全局设置")
    FILE_PREFIX = st.sidebar.text_input("✏️ 自定义文件前缀", value="项目A_")
    ENABLE_COLOR = st.sidebar.checkbox("开启颜色标记", value=True)
    FAST_MODE = st.sidebar.checkbox("开启极速模式", value=False)

    params = {"prefix": FILE_PREFIX, "enable_color": ENABLE_COLOR, "fast_mode": FAST_MODE}

    st.title("🚀 广告素材批量生成工具")
    
    with st.expander("📥 点击展开/收起：快捷资源模板下载", expanded=False):
        c_dl1, c_dl2, c_dl3, c_dl4 = st.columns(4)
        with c_dl1:
            st.download_button("⬇️ 标准素材模板", data=get_template_standard(), file_name="标准素材模板.xlsx", use_container_width=True)
        with c_dl2:
            st.download_button("⬇️ 模块四专用模板", data=get_template_m4(), file_name="模块四专用模板.xlsx", use_container_width=True)
        with c_dl3:
            st.download_button("⬇️ 模块七专用模板", data=get_template_m7(), file_name="模块七专用模板.xlsx", use_container_width=True)
        with c_dl4:
            st.download_button("⬇️ 模块八着陆页模板", data=get_template_m8(), file_name="模块八着陆页导入模板.xlsx", use_container_width=True)

    try:
        if mode == "模块一：基础独立拆分":
            module_1.run(params)
        elif mode == "模块二：同SKU+国家聚合拆分":
            module_2.run(params)
        elif mode == "模块三：智能分组 (SKU去重)":
            module_3.run(params)
        elif mode == "模块四：补齐默认版本":
            module_4.run(params)
        elif mode == "模块五：循环填充与分流":
            module_5.run(params)
        elif mode == "模块六：多行组合循环填充":
            module_6.run(params)
        elif mode == "模块七：账号品类智能匹配":
            module_7.run(params)
        elif mode == "模块八：着陆页导入与素材匹配":
            module_8.run(params)
    finally:
        params.pop('dynamic_hint', None)
        params.pop('repeat_1', None)
        params.pop('repeat_2', None)

else:
    st.title("🔗 广告账号-像素自动匹配工具")
    asset_sub_mode = st.sidebar.radio("请选择功能模块", ["🔗 账号-像素自动匹配"])
    
    def get_asset_template():
        acc_df = pd.DataFrame({"资产": ["资产A"], "账号名称": ["Acc-1"], "账号ID": ["ID001"]})
        pix_df = pd.DataFrame({"资产": ["资产A"], "像素名称": ["Pix-1"], "像素ID": ["P001"]})
        out = io.BytesIO()
        with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
            acc_df.to_excel(writer, index=False, sheet_name="账号表")
            pix_df.to_excel(writer, index=False, sheet_name="像素表")
        return out.getvalue()

    st.download_button("⬇️ 下载：账号像素匹配模板", data=get_asset_template(), file_name="账号像素匹配模板.xlsx", use_container_width=True)

    if asset_sub_mode == "🔗 账号-像素自动匹配":
        asset_tool.run()