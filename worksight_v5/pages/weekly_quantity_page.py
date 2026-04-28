import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="周单件量分析", layout="wide")

st.title("📦 周单件量分析")

# ======================
# 上传 + 来源说明
# ======================
st.markdown("### 上传件量数据（销售单综合）")
st.caption("📌 数据来源：iWMS销售单综合查询")

file = st.file_uploader("", type=["xlsx"], label_visibility="collapsed")

if file:
    try:
        df = pd.read_excel(file)
        df.columns = df.columns.str.strip()

        # ======================
        # ✅ 关键字段校验（核心）
        # ======================
        required_cols = ["打包完成时间", "件数", "京东订单号", "状态"]

        missing_cols = [col for col in required_cols if col not in df.columns]

        if missing_cols:
            st.warning(f"⚠️ 上传的文件格式不正确，缺少字段：{', '.join(missing_cols)}")
            st.info("请上传【iWMS销售单综合查询】导出的表")
            st.stop()

        # ======================
        # 1. 时间字段
        # ======================
        df["打包完成时间"] = pd.to_datetime(df["打包完成时间"], errors="coerce")
        df = df.dropna(subset=["打包完成时间"])

        if df.empty:
            st.warning("⚠️ 时间字段解析失败，数据为空")
            st.stop()

        # ======================
        # 2. 业务日期（5点切日）
        # ======================
        def get_business_date(dt):
            if pd.isna(dt):
                return pd.NaT
            if dt.hour < 5:
                return (dt - pd.Timedelta(days=1)).date()
            else:
                return dt.date()

        df["业务日期"] = df["打包完成时间"].apply(get_business_date)

        # ======================
        # 3. 数据清洗
        # ======================
        df["件数"] = pd.to_numeric(df["件数"], errors="coerce").fillna(0)

        # ======================
        # 4. 过滤
        # ======================
        df = df[df["状态"] == "交接完成"]

        if "是否取消" in df.columns:
            df = df[df["是否取消"] != "是"]

        if df.empty:
            st.warning("⚠️ 过滤后没有有效数据（可能全部被过滤掉）")
            st.stop()

        # ======================
        # 5. 每日统计
        # ======================
        daily = (
            df.groupby("业务日期")
            .agg(
                单量=("京东订单号", "nunique"),
                件量=("件数", "sum"),
            )
            .reset_index()
            .sort_values("业务日期")
        )

        # ======================
        # 6. KPI
        # ======================
        col1, col2 = st.columns(2)

        total_orders = int(daily["单量"].sum())
        total_units = int(daily["件量"].sum())

        col1.metric("总单量", total_orders)
        col2.metric("总件量", total_units)

        # ======================
        # 7. 每日趋势
        # ======================
        st.subheader("📈 每日单量 / 件量趋势")
        st.line_chart(daily.set_index("业务日期")[["单量", "件量"]])

        # ======================
        # 8. 下载 Excel
        # ======================
        def to_excel(df):
            output = BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                df.to_excel(writer, index=False, sheet_name="每日明细")
            return output.getvalue()

        excel_data = to_excel(daily)

        st.download_button(
            label="📥 下载每日明细 Excel",
            data=excel_data,
            file_name="每日单量件量.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        # ======================
        # 9. 表格
        # ======================
        st.subheader("📋 每日明细")
        st.dataframe(daily, width="stretch")

    except Exception as e:
        # ✅ 最终兜底（不会再爆红）
        st.error("❌ 文件解析失败，请确认上传的是正确格式的Excel")
        st.caption("（常见原因：不是销售单综合表 / 表头不对 / 文件损坏）")

else:
    st.info("请上传销售单综合数据")