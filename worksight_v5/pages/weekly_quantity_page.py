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

        # ======================
        # 👤 上传拣货表（iwms拣货结果查询）
        # ======================
        st.markdown("### 👤 每人每日拣货效率")
        st.caption("📌 来源：iWMS拣货结果查询")
        pick_file = st.file_uploader(
            "",
            type=["xlsx"],
            key="pick",
            label_visibility="collapsed"
        )

        if pick_file and file_isc:
            try:
                # ======================
                # 👤 每人每日拣货效率（完整版）
                # ======================

                df_pick = pd.read_excel(pick_file)
                df_pick.columns = df_pick.columns.str.strip()

                df_pick["拣货完成时间"] = pd.to_datetime(df_pick["拣货完成时间"], errors="coerce")
                df_pick = df_pick.dropna(subset=["拣货完成时间"])

                df_pick["日期"] = df_pick["拣货完成时间"].dt.date
                df_pick["实际拣货量"] = pd.to_numeric(df_pick["实际拣货量"], errors="coerce").fillna(0)

                # ======================
                # 1️⃣ 爆品识别
                # ======================
                grouped = df_pick.groupby(['储位', '拣货完成时间', '任务单号']).size().reset_index(name='count')
                burst_groups = grouped[grouped['count'] > 1]

                df_burst = df_pick.merge(
                    burst_groups[['储位', '拣货完成时间', '任务单号']],
                    on=['储位', '拣货完成时间', '任务单号'],
                    how='inner'
                )

                df_pick['是否爆品'] = df_pick.merge(
                    df_burst[['订单号', '储位', '拣货完成时间', '任务单号']],
                    on=['订单号', '储位', '拣货完成时间', '任务单号'],
                    how='left',
                    indicator=True
                )['_merge'] == 'both'

                # ======================
                # 2️⃣ 总件数
                # ======================
                person_qty_raw = (
                    df_pick.groupby(['日期', '工号'])['实际拣货量']
                    .sum()
                    .reset_index(name='总件数')
                )

                # ======================
                # 3️⃣ 去重件数
                # ======================
                df_unique = df_pick.drop_duplicates(
                    subset=['储位', '拣货完成时间', '任务单号', '工号']
                )

                person_qty = (
                    df_unique.groupby(['日期', '工号'])['实际拣货量']
                    .sum()
                    .reset_index(name='件数')
                )

                # ======================
                # 4️⃣ 有效工时
                # ======================
                df_pick = df_pick.sort_values(['工号', '拣货完成时间'])
                df_pick['next_time'] = df_pick.groupby('工号')['拣货完成时间'].shift(-1)

                if '拣货开始时间' in df_pick.columns:
                    df_pick['拣货开始时间'] = pd.to_datetime(df_pick['拣货开始时间'], errors='coerce')
                    df_pick['任务耗时'] = (df_pick['拣货完成时间'] - df_pick['拣货开始时间']).dt.total_seconds()
                    df_pick['任务耗时'] = df_pick['任务耗时'].clip(lower=0, upper=600)
                else:
                    df_pick['任务耗时'] = 0

                df_pick['间隔'] = (df_pick['next_time'] - df_pick['拣货完成时间']).dt.total_seconds()
                df_pick['间隔'] = df_pick['间隔'].clip(lower=0, upper=600).fillna(0)

                df_pick['有效秒'] = df_pick['任务耗时'] + df_pick['间隔']
                df_pick['有效工时'] = df_pick['有效秒'] / 3600

                person_time = (
                    df_pick.groupby(['日期', '工号'])['有效工时']
                    .sum()
                    .reset_index()
                )

                # ======================
                # 5️⃣ ISC工时
                # ======================
                df_att = pd.read_excel(file_isc, sheet_name="明细&汇总-图表")
                df_att.columns = df_att.columns.str.strip()
                df_att["日期"] = pd.to_datetime(df_att["日"], errors="coerce").dt.date

                person_time_isc = (
                    df_att.groupby(["日期", "员工号"])["考勤时长"]
                    .sum()
                    .reset_index()
                )

                person_time_isc.rename(columns={"员工号": "工号"}, inplace=True)

                # ======================
                # 6️⃣ 合并
                # ======================
                person_eff = person_qty.merge(person_time, on=['日期', '工号'], how='left')
                person_eff = person_eff.merge(person_qty_raw, on=['日期', '工号'], how='left')
                person_eff = person_eff.merge(person_time_isc, on=['日期', '工号'], how='left')
                name_map = df_pick[['工号', '姓名']].drop_duplicates()
                person_eff = person_eff.merge(name_map, on='工号', how='left')

                person_eff = person_eff.fillna(0)
                person_eff = person_eff.fillna(0)

                # ======================
                # 7️⃣ 效率
                # ======================
                person_eff['拣非爆品效率'] = person_eff['件数'] / person_eff['有效工时']
                person_eff['总效率'] = person_eff['总件数'] / person_eff['有效工时']

                person_eff = person_eff.replace([float("inf")], 0).fillna(0)

                # ======================
                # 有效工时占比（简化版）
                # ======================
                person_eff["有效工时占比"] = person_eff["有效工时"] / person_eff["考勤时长"]
                person_eff["有效工时占比"] = person_eff["有效工时占比"].replace([float("inf")], 0).fillna(0)
                # ======================
                # 是否低于目标
                # ======================
                person_eff["低于目标"] = person_eff["拣非爆品效率"] < target_upph
                # ======================
                # 日期筛选
                # ======================
                date_list = sorted(person_eff["日期"].dropna().unique())
                selected_date = st.selectbox("选择日期", date_list, key="person_date")

                result = person_eff[person_eff["日期"] == selected_date]
                result = result[
                    [
                        "日期",
                        "工号",
                        "姓名",
                        "总件数",
                        "件数",
                        "有效工时",
                        "考勤时长",
                        "有效工时占比",
                        "拣非爆品效率",
                        "总效率",
                    ]
                ]
                # ======================
                # 初始化缓存（支持切换日期）
                # ======================
                if "edited_cache" not in st.session_state or "current_date" not in st.session_state:
                    st.session_state.edited_cache = result.copy()
                    st.session_state.current_date = selected_date

                if st.session_state.current_date != selected_date:
                    st.session_state.edited_cache = result.copy()
                    st.session_state.current_date = selected_date
                
                # 新增删除列
                if "删除" not in st.session_state.edited_cache.columns:
                    st.session_state.edited_cache["删除"] = False
                # ======================
                # 交互表（可勾选删除）
                # ======================
                edited_df = st.data_editor(
                    st.session_state.edited_cache,
                    use_container_width=True,
                    key="editor",
                    column_config={
                        "删除": st.column_config.CheckboxColumn("删除")
                    },
                    disabled=[
                        "日期", "工号", "姓名", "总件数", "件数",
                        "有效工时", "考勤时长", "有效工时占比",
                        "拣非爆品效率", "总效率"
                    ],
                    hide_index=True
                )
                apply_btn = st.button("应用删除")
                if apply_btn:
                    # ✅ 防止删除列异常（关键修复）
                    edited_df["删除"] = edited_df["删除"].fillna(False).astype(bool)

                    valid_df = edited_df[edited_df["删除"] == False]
                    deleted_df = edited_df[edited_df["删除"] == True]

                    valid_df = valid_df.sort_values(by="拣非爆品效率", ascending=False)

                    st.session_state.edited_cache = pd.concat(
                        [valid_df, deleted_df],
                        ignore_index=True
                    )
                # ======================
                # 灰色显示已删除行
                # ======================
                def highlight_deleted(row):
                    if row["删除"]:
                        return ["background-color: #f2f2f2"] * len(row)
                    return [""] * len(row)

                styled = st.session_state.edited_cache.style.apply(highlight_deleted, axis=1)
                st.dataframe(styled, use_container_width=True, hide_index=True)
            except Exception as e:
                st.error("❌ 生成人效失败")
                st.caption(str(e))

    except Exception as e:
        st.error("❌ 文件解析失败")
        st.caption(str(e))

else:
    st.info("请上传销售单综合数据")