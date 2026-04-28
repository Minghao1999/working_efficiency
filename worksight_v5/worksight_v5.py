import streamlit as st

st.set_page_config(page_title="WorkSight Pro", layout="wide")

# ======================
# 顶部标题
# ======================
st.markdown(
    """
    <h1 style='font-size:42px;'>📊 WorkSight Pro</h1>
    <p style='font-size:18px; color:gray;'>
    仓库运营数据分析平台｜人效 · 单量 · 件量 · 趋势分析
    </p>
    """,
    unsafe_allow_html=True
)

st.markdown("---")

# ======================
# 功能卡片
# ======================
col1, col2 = st.columns(2)

with col1:
    st.markdown(
        """
        <div style="
            padding:20px;
            border-radius:12px;
            border:1px solid #eee;
            box-shadow: 0 2px 6px rgba(0,0,0,0.05);
        ">
            <h3>👷 人效看板</h3>
            <p style='color:gray;'>
            查看员工工作效率、甘特图、Idle / Work 分布
            </p>
        </div>
        """,
        unsafe_allow_html=True
    )

    if st.button("进入人效看板 →", use_container_width=True):
        st.switch_page("pages/work_efficiency_page.py")

with col2:
    st.markdown(
        """
        <div style="
            padding:20px;
            border-radius:12px;
            border:1px solid #eee;
            box-shadow: 0 2px 6px rgba(0,0,0,0.05);
        ">
            <h3>📦 单量 / 件量分析</h3>
            <p style='color:gray;'>
            按天/周统计订单量、件量趋势，支持导出
            </p>
        </div>
        """,
        unsafe_allow_html=True
    )

    if st.button("进入单件量分析 →", use_container_width=True):
        st.switch_page("pages/weekly_quantity_page.py")

# ======================
# 底部说明
# ======================
st.markdown("---")