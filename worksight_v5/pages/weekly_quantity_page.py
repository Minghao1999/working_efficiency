import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="周单件量分析", layout="wide")

st.title("📦 周单件量分析")

# ======================
# 上传区域
# ======================
col_upload1, col_upload2 = st.columns(2)

with col_upload1:
    st.markdown("### 📦 上传件量数据")
    st.caption("📌 来源：iWMS销售单综合查询")
    file = st.file_uploader("", type=["xlsx"], key="volume")

with col_upload2:
    st.markdown("### ⏱ 上传工时数据")
    st.caption("📌 来源：ISC出勤工时->明细&汇总(需要使用Global Export)")
    file_isc = st.file_uploader("", type=["xlsx"], key="isc")

# ======================
# 主逻辑
# ======================
if file:
    try:
        df = pd.read_excel(file)
        df.columns = df.columns.str.strip()

        # ======================
        # 字段校验
        # ======================
        required_cols = ["打包完成时间", "件数", "京东订单号", "状态"]
        missing_cols = [col for col in required_cols if col not in df.columns]

        if missing_cols:
            st.warning(f"⚠️ 缺少字段：{', '.join(missing_cols)}")
            st.stop()

        # ======================
        # 时间处理
        # ======================
        df["打包完成时间"] = pd.to_datetime(df["打包完成时间"], errors="coerce")
        df = df.dropna(subset=["打包完成时间"])

        # ======================
        # 业务日期（5点切日）
        # ======================
        def get_business_date(dt):
            if pd.isna(dt):
                return pd.NaT
            return (dt - pd.Timedelta(days=1)).date() if dt.hour < 5 else dt.date()

        df["业务日期"] = df["打包完成时间"].apply(get_business_date)

        # ======================
        # 数据清洗
        # ======================
        df["件数"] = pd.to_numeric(df["件数"], errors="coerce").fillna(0)

        df = df[df["状态"] == "交接完成"]

        if "是否取消" in df.columns:
            df = df[df["是否取消"] != "是"]

        # ======================
        # 每日件量
        # ======================
        daily = (
            df.groupby("业务日期")
            .agg(
                单量=("京东订单号", "nunique"),
                件量=("件数", "sum"),
            )
            .reset_index()
        )

        # ======================
        # 🆕 处理 ISC 工时
        # ======================
        if file_isc:
            df_isc = pd.read_excel(file_isc, sheet_name="日-考勤-日")
            df_isc.columns = df_isc.columns.str.strip()

            # ⚠️ 根据你实际字段调整
            required_isc_cols = ["工作日", "考勤时长"]

            missing_isc = [c for c in required_isc_cols if c not in df_isc.columns]

            if missing_isc:
                st.warning(f"⚠️ ISC表缺少字段：{', '.join(missing_isc)}")
            else:
                df_isc["工作日"] = pd.to_datetime(df_isc["工作日"], errors="coerce").dt.date
                df_isc["考勤时长"] = pd.to_numeric(df_isc["考勤时长"], errors="coerce").fillna(0)

                work_daily = (
                    df_isc.groupby("工作日")
                    .agg(
                        总工时=("考勤时长", "sum"),
                        出勤人数=("出勤人数", "max"),
                        加班时长=("加班时长", "sum"),
                        OT占比=("OT占比", "mean"),
                    )
                    .reset_index()
                    .rename(columns={"工作日": "业务日期"})
                )

                # 合并
                daily = pd.merge(daily, work_daily, on="业务日期", how="left")

                # UPPH
                daily["UPPH"] = (daily["件量"] / daily["总工时"]).round(2)

        daily = daily.sort_values("业务日期")

        # ======================
        # 🆕 读取仓库名称（明细表）
        # ======================
        try:
            df_wh = pd.read_excel(file_isc, sheet_name="明细&汇总-图表")
            df_wh.columns = df_wh.columns.str.strip()

            if "仓库名称" in df_wh.columns:
                wh_name_full = df_wh["仓库名称"].dropna().iloc[0]

                # 提取 “5号仓”
                import re
                match = re.search(r"(\d+号仓)", str(wh_name_full))
                warehouse_name = match.group(1) if match else wh_name_full
            else:
                warehouse_name = "未知仓库"

        except:
            warehouse_name = "未知仓库"

        # ======================
        # KPI
        # ======================
        # ======================
        # 🎯 仓库目标UPPH（写死规则）
        # ======================
        TARGET_MAP = {
            "1号仓": 34.57,
            "2号仓": 37.11,
            "5号仓": 11.3
        }

        target_upph = TARGET_MAP.get(warehouse_name, 11.3)
        col1, col2, col3 = st.columns(3)

        total_orders = int(daily["单量"].sum())
        total_units = int(daily["件量"].sum())

        col1, col2, col3 = st.columns(3)

        col1.metric("总单量", total_orders)
        col2.metric("总件量", total_units)
        col3.metric("仓库", warehouse_name)
        if file_isc:
            st.metric("目标UPPH", target_upph)

        # ======================
        # 下载
        # ======================
        def to_excel(df):
            output = BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                df.to_excel(writer, index=False, sheet_name="每日明细")
            return output.getvalue()

        st.download_button(
            "📥 下载Excel",
            data=to_excel(daily),
            file_name="单量件量UPPH.xlsx",
        )

        # ======================
        # 表格
        # ======================
        def highlight_low_upph(row):
            if "UPPH" in row and pd.notna(row["UPPH"]):
                if row["UPPH"] < target_upph:
                    return ["background-color: #ffe6e6"] * len(row)
            return [""] * len(row)
        st.subheader("📋 每日明细")
        styled_df = daily.style.apply(highlight_low_upph, axis=1)

        st.dataframe(styled_df, use_container_width=True)

    except Exception as e:
        st.error("❌ 文件解析失败")
        st.caption(str(e))

else:
    st.info("请上传销售单综合数据")