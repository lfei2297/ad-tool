import streamlit as st
import pandas as pd
import io
from modules import module_1, module_2, module_3, module_4

st.set_page_config(page_title="广告素材工具箱", layout="wide")
st.title("🚀 广告素材批量生成工具")

# ==========================
# 📥 资源下载区
# ==========================
def get_template_standard():
    data = {
        "广告账号ID": [""], "主页ID": [""], "像素ID": [""],
        "真实SKU": ["SKU01"], "虚拟SKU": [""], "国家": ["美国"],
        "着陆页版本名称": ["LP-1"], "广告素材版本名称": ["Material-1"],
        "广告素材数量": [5], "素材选取 (X-Y)": [""]
    }
    df = pd.DataFrame(data)
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="标准模板")
    return out.getvalue()

def get_template_m4():
    columns = [
        "广告账号ID", "主页ID", "像素ID", "真实SKU", "虚拟SKU", 
        "国家", "着陆页版本名称", "广告素材版本名称", 
        "提供素材版本数量", "广告组数量", "导品系列数", "补充默认版本数"
    ]
    data = {
        "广告账号ID": ["1018641536297906"], "主页ID": ["391776977343478"], "像素ID": ["806016751628770"],
        "真实SKU": ["L2295705"], "虚拟SKU": [""], "国家": ["德国"], "着陆页版本名称": ["优化组版本-OPDY-1"],
        "广告素材版本名称": ["优化组版本-OPDY-S-2560525-3-1"], "提供素材版本数量": [3],
        "广告组数量": [1], "导品系列数": [1], "补充默认版本数": [1]
    }
    df = pd.DataFrame(data, columns=columns)
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="模块四模板")
    return out.getvalue()

st.markdown("### 📥 资源下载")
col_t1, col_t2 = st.columns(2)
with col_t1:
    st.download_button("⬇️ 下载：模块一/二/三 标准模板", data=get_template_standard(), file_name="标准素材模板.xlsx")
with col_t2:
    st.download_button("⬇️ 下载：模块四 结构补齐模板", data=get_template_m4(), file_name="模块四专用模板.xlsx")
st.markdown("---")

# ==========================
# 🎯 侧边栏
# ==========================
st.sidebar.header("🎯 模块导航")
mode = st.sidebar.radio("请选择功能模块", ["模块一：基础独立拆分", "模块二：同SKU+国家聚合拆分", "模块三：智能分组 (SKU去重)", "模块四：补齐默认版本"])

st.sidebar.markdown("---")
st.sidebar.header("⚙️ 全局设置")
FILE_PREFIX = st.sidebar.text_input("✏️ 自定义文件前缀", value="项目A_")
ENABLE_COLOR = st.sidebar.checkbox("开启颜色标记", value=True)
FAST_MODE = st.sidebar.checkbox("开启极速模式(跳过样式渲染)", value=False)

# 基础参数字典（移除了 repeat）
params = {
    "prefix": FILE_PREFIX,
    "enable_color": ENABLE_COLOR,
    "fast_mode": FAST_MODE
}


# 第三步： 动态加载
if mode == "模块一：基础独立拆分":
    module_1.run(params)
elif mode == "模块二：同SKU+国家聚合拆分":
    module_2.run(params)
elif mode == "模块三：智能分组 (SKU去重)":
    module_3.run(params)
elif mode == "模块四：补齐默认版本":
    module_4.run(params)
