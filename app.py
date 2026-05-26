import streamlit as st
import pandas as pd
import io
import utils  # 先导入工具类
from modules import module_1, module_2, module_3  # 假设你已经拆分了模块

# 1. 页面配置
st.set_page_config(page_title="广告素材批量生成工具", layout="wide")
st.title("🚀 广告素材批量生成工具")

# ==========================
# 📥 第一步：资源下载 (放在这里)
# ==========================
def get_template_excel():
    """生成原始表模板"""
    template_data = {
        "广告账号ID": ["", "", ""],
        "主页ID": ["", "", ""],
        "像素ID": ["", "", ""],
        "真实SKU": ["SKU001", "SKU002", "SKU003"],
        "虚拟SKU": ["V-SKU001", "", ""],
        "国家": ["美国", "德国", "英国"],
        "着陆页版本名称": ["优化组版本-LP-1", "优化组版本-LP-2", "优化组版本-LP-1"],
        "广告素材版本名称": ["素材版本-A-1", "素材版本-B-5", "素材版本-C-1"],
        "广告素材数量": [5, 3, 10],
        "素材选取 (X-Y)": ["", "1-3", "5-14"]
    }
    df = pd.DataFrame(template_data)
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="模板")
    return out.getvalue()

# 在主界面顶部渲染下载区域
st.markdown("### 📥 资源下载")
st.download_button(
    label="⬇️ 下载：原始素材表 Excel 模板",
    data=get_template_excel(),
    file_name="广告素材原始表模板.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

st.markdown("---") # 分割线，区分下载区和下方的功能操作区

# ==========================
# 2. 侧边栏及功能模块加载
# ==========================

st.sidebar.header("🎯 模块导航")
mode = st.sidebar.radio("请选择功能模块", ["模块一", "模块二", "模块三", "模块四(待开发)"])

# 2. 公共参数透传
st.sidebar.header("⚙️ 全局参数设置")

# 1. 文件前缀
FILE_PREFIX = st.sidebar.text_input("✏️ 自定义结果文件前缀", value="项目A_")

# 2. 重复次数逻辑 (模块一、二、三通用的展开逻辑)
st.sidebar.subheader("🔄 素材展开设置")
REPEAT_FIRST = st.sidebar.number_input("第一次重复次数", min_value=1, value=1, help="单条素材基础展开次数")
REPEAT_SECOND = st.sidebar.number_input("第二次重复次数", min_value=1, value=1, help="总表整体复制次数")

# 3. 业务层级的数量设置 (用于文件名或特定逻辑)
st.sidebar.subheader("📊 业务规模设置")
AD_GROUP_COUNT = st.sidebar.number_input("广告组数", min_value=1, value=1)
SERIES_COUNT = st.sidebar.number_input("系列数", min_value=1, value=1)

# 4. 开关
ENABLE_COLOR = st.sidebar.checkbox("开启颜色标记", value=True)
FAST_MODE = st.sidebar.checkbox("开启极速模式 (跳过样式渲染)", value=False)

# 将所有参数打包传给模块
params = {
    "prefix": FILE_PREFIX,
    "repeat_1": REPEAT_FIRST,
    "repeat_2": REPEAT_SECOND,
    "ad_group": AD_GROUP_COUNT,
    "series": SERIES_COUNT,
    "enable_color": ENABLE_COLOR,
    "fast_mode": FAST_MODE
}

# 3. 动态加载
if mode == "模块一":
    module_1.run(params)
elif mode == "模块二":
    module_2.run(params)
elif mode == "模块三":
    module_3.run(params)