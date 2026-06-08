"""大津京こと薬局 シフト自動作成アプリ
起動: streamlit run app.py  （または launch.pyw をダブルクリック）
"""
import calendar
import datetime

import streamlit as st
from shift_engine import NAMES, SHOTEIKOJI, generate_shift

st.set_page_config(
    page_title="大津京こと薬局 シフト作成",
    page_icon="💊",
    layout="wide",
)
st.title("💊 大津京こと薬局　シフト自動作成")

# ============================================================
# サイドバー
# ============================================================
with st.sidebar:
    st.header("📅 基本設定")
    today      = datetime.date.today()
    next_month = today.month % 12 + 1
    next_year  = today.year + (1 if today.month == 12 else 0)
    year  = int(st.number_input("年", min_value=2024, max_value=2035, value=next_year, step=1))
    month = int(st.number_input("月", min_value=1,    max_value=12,   value=next_month, step=1))

    days_in_month = calendar.monthrange(year, month)[1]
    WD = ["月", "火", "水", "木", "金", "土", "日"]

    def label(d):
        return f"{d}日（{WD[datetime.date(year, month, d).weekday()]}）"

    all_days  = list(range(1, days_in_month + 1))
    open_days = [d for d in all_days if datetime.date(year, month, d).weekday() != 6]

    st.subheader("臨時休業")
    hol_input = st.multiselect("臨時休業日", options=open_days, format_func=label)
    holidays  = set(hol_input)

    shoteikoji_val = SHOTEIKOJI.get(year, {}).get(month)
    if shoteikoji_val:
        st.divider()
        st.caption(f"正社員 所定労働時間: {shoteikoji_val}h")

    # 日祝担当 上書き
    st.divider()
    with st.expander("日祝担当 上書き（空欄=自動）"):
        sun_hol_preview = [d for d in all_days
                           if datetime.date(year, month, d).weekday() == 6 or d in holidays]
        sun_override_ph: dict = {}
        sun_override_jm: dict = {}
        for d in sun_hol_preview:
            lbl = "祝" if d in holidays else "日"
            wd  = WD[datetime.date(year, month, d).weekday()]
            ph_sel = st.selectbox(
                f"{d}日({wd}/{lbl}) 薬剤師",
                options=["自動", "B", "C", "D"],
                key=f"ph_{d}",
            )
            jm_sel = st.selectbox(
                f"{d}日({wd}/{lbl}) 事務",
                options=["自動", "H", "I", "J"],
                key=f"jm_{d}",
            )
            if ph_sel != "自動": sun_override_ph[d] = ph_sel
            if jm_sel != "自動": sun_override_jm[d] = jm_sel

# ============================================================
# 入力エリア（タブ）
# ============================================================
requested_off    = {k: set() for k in "ABCDEFGHIJK"}
yukyu_per_person = {k: set() for k in "ABCDEFGHIJK"}
extra_off_config = {k: set() for k in "BCDHIJ"}
b_fri = set(); c_fri = set(); d_fri = set()
b_sat = set(); c_sat = set(); d_sat = set()

tab1, tab2, tab3 = st.tabs(["📌 希望休 / 有休", "🗓 追加休日（B〜J）", "📋 金土出勤指定（B/C/D）"])

with tab1:
    col_l, col_r = st.columns(2)
    with col_l:
        st.subheader("希望休（×）")
        for k in "ABCDEFGHIJK":
            days_sel = st.multiselect(
                f"{k}（{NAMES[k]}）", options=open_days, format_func=label, key=f"off_{k}")
            requested_off[k] = set(days_sel)
    with col_r:
        st.subheader("有休")
        for k in "ABCDEFGHIJK":
            days_sel = st.multiselect(
                f"{k}（{NAMES[k]}）", options=open_days, format_func=label, key=f"yu_{k}")
            yukyu_per_person[k] = set(days_sel)

with tab2:
    st.subheader("追加休日")
    st.caption("日曜補償は自動追加されます。それ以外に休ませたい日を指定してください。")
    cols2 = st.columns(3)
    for i, k in enumerate("BCDHIJ"):
        with cols2[i % 3]:
            days_sel = st.multiselect(
                f"{k}（{NAMES[k]}）", options=open_days, format_func=label, key=f"ex_{k}")
            extra_off_config[k] = set(days_sel)

with tab3:
    st.subheader("金曜出勤日指定")
    st.caption("全て空欄の場合、所定労働時間に合わせて自動配置します。")
    fri_days = [d for d in open_days if datetime.date(year, month, d).weekday() == 4]
    sat_days_list = [d for d in open_days if datetime.date(year, month, d).weekday() == 5]
    c1, c2, c3 = st.columns(3)
    with c1:
        b_fri = set(st.multiselect("B（中本）金曜", options=fri_days, format_func=label, key="bfri"))
        b_sat = set(st.multiselect("B（中本）土曜", options=sat_days_list, format_func=label, key="bsat"))
    with c2:
        c_fri = set(st.multiselect("C（今堀）金曜", options=fri_days, format_func=label, key="cfri"))
        c_sat = set(st.multiselect("C（今堀）土曜", options=sat_days_list, format_func=label, key="csat"))
    with c3:
        d_fri = set(st.multiselect("D（安井）金曜", options=fri_days, format_func=label, key="dfri"))
        d_sat = set(st.multiselect("D（安井）土曜", options=sat_days_list, format_func=label, key="dsat"))

# ============================================================
# 生成ボタン
# ============================================================
st.divider()
if st.button("🚀　シフトを生成する", type="primary", use_container_width=True):
    config = {
        "year": year, "month": month,
        "holidays":         holidays,
        "sun_override_ph":  sun_override_ph,
        "sun_override_jm":  sun_override_jm,
        "b_fri": b_fri, "c_fri": c_fri, "d_fri": d_fri,
        "b_sat": b_sat, "c_sat": c_sat, "d_sat": d_sat,
        "b_extra_off": extra_off_config["B"],
        "c_extra_off": extra_off_config["C"],
        "d_extra_off": extra_off_config["D"],
        "h_extra_off": extra_off_config["H"],
        "i_extra_off": extra_off_config["I"],
        "j_extra_off": extra_off_config["J"],
        "requested_off":    requested_off,
        "yukyu_per_person": yukyu_per_person,
    }

    with st.spinner("シフトを計算中..."):
        try:
            shift_data, excel_bytes, verif = generate_shift(config)
        except Exception as ex:
            st.error(f"生成エラー: {ex}")
            st.exception(ex)
            st.stop()

    st.success("✅ 生成完了！")

    # 日祝ローテーション
    with st.expander("📅 日祝担当ローテーション", expanded=True):
        for d in verif["sun_hol_days"]:
            lbl     = "祝" if d in holidays else "日"
            wd_lbl  = WD[datetime.date(year, month, d).weekday()]
            ph_key  = verif["ph_assign"].get(d, "-")
            jm_key  = verif["jm_assign"].get(d, "-")
            st.write(
                f"{d}日（{wd_lbl}/{lbl}）　"
                f"薬: {NAMES.get(ph_key, '-')}　事務: {NAMES.get(jm_key, '-')}"
            )

    # 労働時間（正社員のみ）
    st.subheader("📊 労働時間（正社員）")
    main_staff = list("BCDHIJ")
    cols_h = st.columns(6)
    for i, n in enumerate(main_staff):
        h_info = verif["hours"].get(n)
        if not h_info:
            continue
        diff = h_info["diff"]
        cols_h[i].metric(
            label=f"{n}（{NAMES[n]}）",
            value=f"{h_info['total']:.2f}h",
            delta=f"{diff:+.2f}h（目標 {h_info['target']:.0f}h）",
            delta_color="normal" if abs(diff) <= 0.5 else ("inverse" if diff > 0 else "off"),
        )

    # 連勤チェック
    st.subheader("📅 連勤チェック")
    cols_s = st.columns(6)
    for i, n in enumerate(main_staff):
        mx = verif["streaks"].get(n, 0)
        cols_s[i].metric(
            f"{n}（{NAMES[n]}）",
            f"最大 {mx} 日",
            delta="⚠ 要確認" if mx >= 6 else "OK",
            delta_color="inverse" if mx >= 6 else "normal",
        )

    # チェック結果
    cov = verif["coverage_issues"]
    sun = verif["sunhol_issues"]
    if cov or sun:
        st.subheader("⚠ チェック結果")
        for msg in cov: st.error(msg)
        for msg in sun: st.error(msg)
    else:
        st.success("✅ 薬剤師カバレッジ・日祝配置：全日程クリア")

    # ダウンロード
    st.divider()
    filename = f"大津京こと薬局シフト{year}年{month:02d}月.xlsx"
    st.download_button(
        label=f"⬇　{filename} をダウンロード",
        data=excel_bytes,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )
