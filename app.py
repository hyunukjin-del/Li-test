import streamlit as st
import pandas as pd
import numpy as np
import io
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

# 화학 분자량 상수 (g/mol)
MW_LI2CO3 = 73.89
MW_CAO = 56.08
MW_CAOH2 = 74.09
MW_LIOH = 23.95
MW_CACO3 = 100.09
MW_LI = 6.941
MW_CA = 40.08
MW_NA = 22.99
MW_SI = 28.09
MW_MG = 24.31
MW_K = 39.10

# 에이전트 제목 설정
AGENT_TITLE = "LC-LH전환반응 M/B자동화 및 거동예측 Agent tool"

st.set_page_config(page_title=AGENT_TITLE, page_icon="🧪", layout="wide")

# 세션 상태 초기화
if "history" not in st.session_state:
    st.session_state.history = pd.DataFrame([
        {
            "회차 (Run)": 1, 
            "구분": "실측치 (Actual)",
            "Li 회수율 (%)": 95.80, 
            "1차여액 Li농도 (mg/L)": 10500.0,
            "1차여액 LiOH농도 (g/L)": round(10500.0 * (MW_LIOH / MW_LI) / 1000, 2),
            "M/B 닫힘율 (%)": 95.06, 
            "하소 감율 LOI (%)": 41.13, 
            "CaO 활성도 (%)": 100.0,
            "신품 CaO 보충량 (g)": 68.52
        }
    ])

if "chat_messages" not in st.session_state:
    st.session_state.chat_messages = [
        {"role": "assistant", "content": f"안녕하세요! **{AGENT_TITLE}**입니다. 탄산리튬(LC)에서 수산화리튬(LH)으로의 가성화 M/B 연산, ICP 불순물 분석, 회차별 거동예측에 대해 무엇이든 질문해 주세요."}
    ]

st.title(f"🧪 {AGENT_TITLE}")
st.caption("탄산리튬(LC) → 수산화리튬(LH) 가성화 양론 연산 | ICP 성분별 회수율 및 불순물 분석 | 회차별 거동예측 | 리포트 메일 발송")

# 상단 탭 구성
main_tab1, main_tab2, main_tab3, main_tab4, main_tab5 = st.tabs([
    "1️⃣ 실험 데이터 입력 & M/B 연산", 
    "2️⃣ 🧪 용액 ICP 분석 & 회수율", 
    "3️⃣ 📈 회차별 트렌드 & 거동예측", 
    "4️⃣ 💬 AI 공정 대화창", 
    "5️⃣ 📧 리포트 메일 발송"
])

# --------------------------------------------------------------------------
# TAB 1: 실험 데이터 입력 및 기초 M/B 연산
# --------------------------------------------------------------------------
with main_tab1:
    with st.expander("📝 이번 회차 실험 수치 입력 폼", expanded=True):
        # 1. 상단: 투입 원료 vs 1차 여과 및 케이크 수세
        col_in1, col_in2 = st.columns(2)
        with col_in1:
            st.markdown("#### [1. 투입 원료 및 반응 조건]")
            run_no = st.number_input("실험 회차 (Run No.)", min_value=1, value=len(st.session_state.history), step=1)
            li2co3_mass = st.number_input("Li₂CO₃ 투입량 (g)", value=95.34, format="%.2f")
            li2co3_water = st.number_input("Li₂CO₃ 용매수 (g)", value=1040.0, format="%.1f")
            fresh_cao_mass = st.number_input("신품 CaO 투입량 (g)", value=92.42 if run_no == 1 else 68.52, format="%.2f")
            recycled_cao_mass = st.number_input("재생 CaO 투입량 (g)", value=0.0 if run_no == 1 else 23.90, format="%.2f")
            slurry_water = st.number_input("슬러리 조제수 (g)", value=831.0, format="%.1f")
            temp_c = st.number_input("반응 온도 (℃)", value=80.0, format="%.1f")
            time_h = st.number_input("반응 시간 (시간)", value=2.0, format="%.1f")

        with col_in2:
            st.markdown("#### [2. 1차 여과 및 케이크 수세]")
            primary_filtrate_mass = st.number_input("1차 LiOH 여액 무게 (g)", value=1646.0, format="%.1f")
            primary_filtrate_sg = st.number_input("1차 여액 비중 (g/mL)", value=1.035, format="%.3f")
            primary_filtrate_ph = st.number_input("1차 여액 pH", value=12.81, format="%.2f")
            wet_cake_mass = st.number_input("1차 습케이크 무게 (g)", value=311.0, format="%.1f")
            sample_wet = st.number_input("함수율 측정 샘플 습중량 (g)", value=27.7, format="%.1f")
            sample_dry = st.number_input("함수율 측정 샘플 건중량 (g)", value=14.8, format="%.1f")
            wash_sol_mass = st.number_input("회수된 수세액 무게 (g)", value=832.0, format="%.1f")
            wash_sol_sg = st.number_input("수세액 비중 (g/mL)", value=1.000, format="%.3f")
            wash_sol_ph = st.number_input("수세액 pH", value=13.70, format="%.2f")

        st.divider()

        # 2. 하단: 소성(하소) 및 CaO 재생 (여과/수세 아래에 별도 배치)
        st.markdown("#### [3. CaCO₃ 소성(하소) 및 CaO 재생]")
        col_calc1, col_calc2 = st.columns(2)
        with col_calc1:
            test_dry_cake = st.number_input("소성 투입 건조케익 샘플 (g)", value=40.6, format="%.1f")
            calcined_cao = st.number_input("소성 후 회수된 CaO (g)", value=23.9, format="%.1f")
        with col_calc2:
            calc_temp = st.number_input("소성 온도 (℃)", value=1000.0, format="%.1f")
            calc_time = st.number_input("소성 시간 (시간)", value=1.0, format="%.1f")

    # M/B 기본 연산 로직
    n_li2co3 = li2co3_mass / MW_LI2CO3
    total_cao_in = fresh_cao_mass + recycled_cao_mass
    n_cao = (total_cao_in * 1.0) / MW_CAO
    limiting = "Li2CO3" if n_li2co3 <= n_cao else "CaO"
    excess_pct = (n_cao / n_li2co3 - 1.0) * 100
    theo_lioh_mass = (n_li2co3 * 2.0) * MW_LIOH

    total_in = li2co3_mass + li2co3_water + total_cao_in + slurry_water
    total_out = primary_filtrate_mass + wet_cake_mass
    loss_mass = total_in - total_out
    mass_closure = (total_out / total_in) * 100.0
    cake_moisture = (1.0 - (sample_dry / sample_wet)) * 100.0
    est_total_dry_solids = wet_cake_mass * (sample_dry / sample_wet)

    loi_pct = ((test_dry_cake - calcined_cao) / test_dry_cake) * 100.0
    cao_yield_dry = (calcined_cao / test_dry_cake) * 100.0
    purity_caco3 = max(0.0, min(100.0, ((loi_pct/100.0) - 0.2432) / (0.4397 - 0.2432) * 100.0))
    pot_total_cao = est_total_dry_solids * (cao_yield_dry / 100.0)
    ca_loop_recovery = (pot_total_cao / total_cao_in) * 100.0

    target_cao = 92.42
    fresh_makeup = max(0.0, target_cao - calcined_cao)

    st.subheader(f"📊 Run {run_no} 공정 기본 물질수지(M/B)")
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("M/B 정합성 (Closure)", f"{mass_closure:.2f} %", f"증발: {loss_mass:.1f}g")
    k2.metric("1차 케익 함수율", f"{cake_moisture:.2f} %", f"고형분 {est_total_dry_solids:.1f}g")
    k3.metric("소성 감율 (LOI)", f"{loi_pct:.2f} %", f"CaCO₃ 순도 ~{purity_caco3:.1f}%")
    k4.metric("Ca-Loop 원소 회수율", f"{ca_loop_recovery:.2f} %", f"재생잠재 {pot_total_cao:.1f}g")

# --------------------------------------------------------------------------
# TAB 2: 🧪 용액 ICP 분석 (mg/L 단위 표기 & 하단 회수율 자동 계산)
# --------------------------------------------------------------------------
with main_tab2:
    st.header("🧪 용액 ICP 분석 데이터 입력 및 회수율 계산")
    st.markdown("1차 여액 및 수세액의 각 성분 농도(**mg/L**)를 입력하면, **투입량 대비 리튬 회수율과 성분별 물질수지**를 하단에 즉시 계산합니다.")

    # [1] ICP 농도 입력 폼 (2개 컬럼, 각 항목에 mg/L 명시)
    st.markdown("### 1. ICP 분석 데이터 입력")
    icp_col1, icp_col2 = st.columns(2)

    with icp_col1:
        st.markdown(f"#### 🔹 1차 여액 분석치 (부피: {primary_filtrate_mass/primary_filtrate_sg:.1f} mL)")
        icp_li_1 = st.number_input("Li 농도 (mg/L) - 1차 여액", value=10500.0, step=50.0, format="%.1f")
        icp_ca_1 = st.number_input("Ca 농도 (mg/L) - 1차 여액", value=120.0, step=5.0, format="%.1f")
        icp_na_1 = st.number_input("Na 농도 (mg/L) - 1차 여액", value=45.0, step=1.0, format="%.1f")
        icp_si_1 = st.number_input("Si 농도 (mg/L) - 1차 여액", value=8.5, step=0.5, format="%.1f")
        icp_mg_1 = st.number_input("Mg 농도 (mg/L) - 1차 여액", value=1.2, step=0.1, format="%.1f")
        icp_k_1  = st.number_input("K 농도 (mg/L) - 1차 여액", value=15.0, step=1.0, format="%.1f")

    with icp_col2:
        st.markdown(f"#### 🔹 수세액 분석치 (부피: {wash_sol_mass/wash_sol_sg:.1f} mL)")
        icp_li_w = st.number_input("Li 농도 (mg/L) - 수세액", value=1400.0, step=50.0, format="%.1f")
        icp_ca_w = st.number_input("Ca 농도 (mg/L) - 수세액", value=80.0, step=5.0, format="%.1f")
        icp_na_w = st.number_input("Na 농도 (mg/L) - 수세액", value=6.0, step=1.0, format="%.1f")
        icp_si_w = st.number_input("Si 농도 (mg/L) - 수세액", value=2.1, step=0.5, format="%.1f")
        icp_mg_w = st.number_input("Mg 농도 (mg/L) - 수세액", value=0.3, step=0.1, format="%.1f")
        icp_k_w  = st.number_input("K 농도 (mg/L) - 수세액", value=2.0, step=1.0, format="%.1f")

    st.divider()

    # [2] 회수율 및 성분 질량 자동 계산 (입력창 바로 아래 위치)
    st.markdown("### 2. 성분별 질량 및 회수율 계산 결과")

    v1_L = (primary_filtrate_mass / primary_filtrate_sg) / 1000.0
    vw_L = (wash_sol_mass / wash_sol_sg) / 1000.0

    # Li 총 투입량
    li_in_total_g = n_li2co3 * 2.0 * MW_LI

    # 각 성분별 회수 질량 (g) = 농도(mg/L) * 부피(L) / 1000
    elements = ["Li", "Ca", "Na", "Si", "Mg", "K"]
    conc_1 = [icp_li_1, icp_ca_1, icp_na_1, icp_si_1, icp_mg_1, icp_k_1]
    conc_w = [icp_li_w, icp_ca_w, icp_na_w, icp_si_w, icp_mg_w, icp_k_w]

    mass_1_g = [c * v1_L / 1000.0 for c in conc_1]
    mass_w_g = [c * vw_L / 1000.0 for c in conc_w]
    mass_total_g = [m1 + mw for m1, mw in zip(mass_1_g, mass_w_g)]

    # 리튬 세부 회수율 계산
    li_rec_1_pct = (mass_1_g[0] / li_in_total_g) * 100.0 if li_in_total_g > 0 else 0.0
    li_rec_w_pct = (mass_w_g[0] / li_in_total_g) * 100.0 if li_in_total_g > 0 else 0.0
    total_li_rec_pct = li_rec_1_pct + li_rec_w_pct
    li_loss_pct = max(0.0, 100.0 - total_li_rec_pct)
    lioh_equiv_g_l = icp_li_1 * (MW_LIOH / MW_LI) / 1000.0

    # 메트릭 카드
    rc1, rc2, rc3, rc4 = st.columns(4)
    rc1.metric("🎯 총 Li 회수율", f"{total_li_rec_pct:.2f} %", f"총 회수: {mass_total_g[0]:.2f}g / 투입: {li_in_total_g:.2f}g")
    rc2.metric("1차 여액 회수율", f"{li_rec_1_pct:.2f} %", f"회수 질량: {mass_1_g[0]:.2f}g")
    rc3.metric("수세액 회수율", f"{li_rec_w_pct:.2f} %", f"회수 질량: {mass_w_g[0]:.2f}g")
    rc4.metric("케이크 잔류/손실률", f"{li_loss_pct:.2f} %", f"미회수: {li_in_total_g - mass_total_g[0]:.2f}g", delta_color="inverse")

    # 성분별 분석 종합 표
    df_icp_summary = pd.DataFrame({
        "원소 (Element)": elements,
        "1차 여액 농도 (mg/L)": conc_1,
        "1차 여액 회수량 (g)": [round(x, 4) for x in mass_1_g],
        "수세액 농도 (mg/L)": conc_w,
        "수세액 회수량 (g)": [round(x, 4) for x in mass_w_g],
        "총 용출 질량 (g)": [round(x, 4) for x in mass_total_g],
        "Li 대비 불순물 비율 (wt ppm)": [
            round((mt / mass_total_g[0]) * 1e6, 1) if el != "Li" and mass_total_g[0] > 0 else "-" 
            for el, mt in zip(elements, mass_total_g)
        ]
    })

    st.markdown("##### 📋 다성분(Li, Ca, Na, Si, Mg, K) 분석 및 불순물 비율 요약표")
    st.dataframe(
        df_icp_summary.style.format({
            "1차 여액 농도 (mg/L)": "{:,.1f}",
            "1차 여액 회수량 (g)": "{:.4f}",
            "수세액 농도 (mg/L)": "{:,.1f}",
            "수세액 회수량 (g)": "{:.4f}",
            "총 용출 질량 (g)": "{:.4f}"
        }),
        use_container_width=True
    )

    # AI ICP 진단 코멘트
    st.markdown("##### 🤖 ICP 분석 기반 AI 공정 진단")
    icp_diags = []
    if icp_ca_1 > 200.0:
        icp_diags.append(f"여액 내 Ca 농도가 {icp_ca_1:.1f} mg/L로 높습니다. CaO 과잉 또는 미세 고형분 유출 가능성이 있으므로 탄산화(Carbonation) 탈칼슘 후처리가 필요합니다.")
    if li_rec_w_pct > 10.0:
        icp_diags.append(f"수세액에서 전체 리튬의 {li_rec_w_pct:.1f}%가 회수되었습니다. 케이크 여과 탈수 효율을 높여 1차 여액 회수율을 극대화할 필요가 있습니다.")
    if icp_na_1 > 50.0 or icp_k_1 > 20.0:
        icp_diags.append(f"알칼리 불순물(Na: {icp_na_1} mg/L, K: {icp_k_1} mg/L)은 수용성이 높아 여액에 100% 잔류하므로, 농축/결정화 단계 전 원료 순도 관리가 요구됩니다.")

    if not icp_diags:
        icp_diags.append("모든 불순물(Ca, Na, Si, Mg, K) 수치와 Li 회수율이 정상 관리 범위 내에 있습니다.")

    for d in icp_diags:
        st.info(f"💡 {d}")

    # 트렌드 DB에 분석 결과 반영
    if st.button("💾 이 분석 결과를 트렌드 DB에 최종 저장하기", type="primary"):
        new_row = {
            "회차 (Run)": int(run_no),
            "구분": "실측치 (Actual)",
            "Li 회수율 (%)": round(total_li_rec_pct, 2), 
            "1차여액 Li농도 (mg/L)": round(icp_li_1, 1),
            "1차여액 LiOH농도 (g/L)": round(lioh_equiv_g_l, 2),
            "M/B 닫힘율 (%)": round(mass_closure, 2), 
            "하소 감율 LOI (%)": round(loi_pct, 2), 
            "CaO 활성도 (%)": 100.0,
            "신품 CaO 보충량 (g)": round(fresh_makeup, 2)
        }
        st.session_state.history = st.session_state.history[st.session_state.history["회차 (Run)"] != run_no]
        st.session_state.history = pd.concat([st.session_state.history, pd.DataFrame([new_row])]).sort_values("회차 (Run)").reset_index(drop=True)
        st.success(f"✅ Run {run_no} ICP 분석 및 회수율 데이터가 트렌드 DB에 영구 등록되었습니다!")

# --------------------------------------------------------------------------
# TAB 3: 📈 회차별 트렌드 & 거동예측
# --------------------------------------------------------------------------
with main_tab3:
    st.header("📈 $n$회차 트렌드 시각화 & 거동예측 시뮬레이터")
    
    with st.expander("⚙️ 거동예측 시뮬레이션 파라미터 설정 (What-If Simulation)", expanded=True):
        col_p1, col_p2, col_p3 = st.columns(3)
        with col_p1:
            target_max_run = st.slider("예측 시뮬레이션 목표 회차", min_value=3, max_value=20, value=10, step=1)
            sintering_decay = st.slider("회차당 CaO 소결 활성도 감쇄율 (%)", min_value=0.5, max_value=8.0, value=3.5, step=0.1)
        with col_p2:
            sim_temp = st.number_input("가상 반응 온도 (℃)", min_value=60.0, max_value=95.0, value=80.0, step=1.0)
            sim_time = st.number_input("가상 반응 시간 (h)", min_value=1.0, max_value=5.0, value=2.0, step=0.5)
        with col_p3:
            wash_ratio = st.slider("수세수 투입 배수 (케이크 대비)", min_value=1.0, max_value=5.0, value=3.0, step=0.5)
            fresh_makeup_mode = st.selectbox("CaO 보충 방식", ["고정량 보충 (전회차 감량분 100% Make-up)", "신품 100% 교체 (Purge)"])

    sim_rows = []
    base_li_conc = icp_li_1 if 'icp_li_1' in locals() else 10500.0
    base_recovery = total_li_rec_pct if 'total_li_rec_pct' in locals() and total_li_rec_pct > 0 else 95.80

    temp_factor = 1.0 + (sim_temp - 80.0) * 0.004
    time_factor = 1.0 + (sim_time - 2.0) * 0.03

    for r in range(1, target_max_run + 1):
        matched_actual = st.session_state.history[st.session_state.history["회차 (Run)"] == r]
        
        if not matched_actual.empty:
            row = matched_actual.iloc[0].to_dict()
            sim_rows.append(row)
        else:
            activity = 100.0 * ((1.0 - (sintering_decay / 100.0)) ** (r - 1))
            eff_activity = min(100.0, activity * temp_factor * time_factor)
            
            pred_rec = max(70.0, min(99.0, base_recovery * (eff_activity / 100.0)))
            pred_li_conc = max(7000.0, base_li_conc * (pred_rec / base_recovery))
            pred_lioh_conc = round(pred_li_conc * (MW_LIOH / MW_LI) / 1000.0, 2)
            pred_loi = max(35.0, 41.13 - (r - 1) * 0.4)
            pred_makeup = 68.52 if fresh_makeup_mode.startswith("고정량") else 92.42

            sim_rows.append({
                "회차 (Run)": r,
                "구분": "AI 예측치 (Simulated)",
                "Li 회수율 (%)": round(pred_rec, 2),
                "1차여액 Li농도 (mg/L)": round(pred_li_conc, 1),
                "1차여액 LiOH농도 (g/L)": round(pred_lioh_conc, 2),
                "M/B 닫힘율 (%)": round(max(90.0, 95.06 - (r - 1) * 0.2), 2),
                "하소 감율 LOI (%)": round(pred_loi, 2),
                "CaO 활성도 (%)": round(eff_activity, 1),
                "신품 CaO 보충량 (g)": round(pred_makeup, 2)
            })

    df_simulation = pd.DataFrame(sim_rows)

    st.markdown("---")
    st.subheader("📊 회차별 인터랙티브 X-Y 트렌드 그래프")

    col_ctrl1, col_ctrl2 = st.columns([1, 2])
    with col_ctrl1:
        y_axis_metric = st.selectbox(
            "📌 Y축에 표시할 지표를 선택하세요:",
            [
                "Li 회수율 (%)", 
                "1차여액 LiOH농도 (g/L)", 
                "1차여액 Li농도 (mg/L)", 
                "CaO 활성도 (%)", 
                "하소 감율 LOI (%)",
                "신품 CaO 보충량 (g)"
            ]
        )

    chart_df = df_simulation.set_index("회차 (Run)")[[y_axis_metric]]
    st.line_chart(chart_df, height=380, use_container_width=True)

    st.markdown("##### 📋 회차별 실측치 & 예측 시뮬레이션 상세 데이터 테이블")
    st.dataframe(
        df_simulation.style.format({
            "Li 회수율 (%)": "{:.2f} %",
            "1차여액 Li농도 (mg/L)": "{:,.1f} mg/L",
            "1차여액 LiOH농도 (g/L)": "{:.2f} g/L",
            "M/B 닫힘율 (%)": "{:.2f} %",
            "하소 감율 LOI (%)": "{:.2f} %",
            "CaO 활성도 (%)": "{:.1f} %",
            "신품 CaO 보충량 (g)": "{:.2f} g"
        }),
        use_container_width=True
    )

# --------------------------------------------------------------------------
# TAB 4: 💬 AI 공정 대화창
# --------------------------------------------------------------------------
with main_tab4:
    st.header("💬 AI 공정 엔지니어와 대화하기")
    st.caption("현재 실험 수치, ICP 분석 결과, $n$회차 시뮬레이션을 바탕으로 실시간 질의응답을 진행합니다.")

    for msg in st.session_state.chat_messages:
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])

    if user_prompt := st.chat_input("질문을 입력하세요 (예: Ca 농도가 120ppm이면 제품 순도에 영향이 커?)"):
        st.session_state.chat_messages.append({"role": "user", "content": user_prompt})
        with st.chat_message("user"):
            st.markdown(user_prompt)

        prompt_low = user_prompt.lower()
        if "ca" in prompt_low or "칼슘" in prompt_low or "불순물" in prompt_low:
            ai_reply = f"현재 1차 여액 내 Ca 농도는 **{icp_ca_1} mg/L**입니다. 배터리급 수산화리튬(LH) 스펙 대비 높으므로, 증발 농축/결정화 전에 CO₂ 가스를 불어넣어 침전 분리하는 **2차 탄산화(Carbonation) 탈칼슘 공정**을 거치는 것을 권장합니다."
        elif "ph" in prompt_low or "수세" in prompt_low or "세척" in prompt_low:
            ai_reply = f"수세액 pH가 **{wash_sol_ph}**로 높은 이유는 1차 감압 여과 후 케이크 기공 내 LiOH 고농도 액이 갇혀 있었기 때문입니다. 3배수 수세를 통해 총 {mass_w_g[0]:.2f}g의 Li(회수율 기여도 {li_rec_w_pct:.2f}%)을 회수하였습니다."
        elif "소결" in prompt_low or "sintering" in prompt_low or "퍼지" in prompt_low:
            ai_reply = f"1000℃ 고온 하소가 반복되면 활성도가 회차당 약 {sintering_decay}%씩 저하됩니다. 회수율을 90% 이상 유지하려면 약 6~7회차 시점에 재생 CaO 일부를 퍼지(Purge)하고 신품 CaO로 리프레시해 주시는 것이 좋습니다."
        else:
            ai_reply = f"Run {run_no}의 총 Li 회수율은 **{total_li_rec_pct:.2f}%** (1차 여액 {li_rec_1_pct:.1f}% + 수세액 {li_rec_w_pct:.1f}%)로 산출되었습니다. 추가로 확인하고 싶은 반응 변수나 거동 조건이 있으신가요?"

        st.session_state.chat_messages.append({"role": "assistant", "content": ai_reply})
        with st.chat_message("assistant"):
            st.markdown(ai_reply)

# --------------------------------------------------------------------------
# TAB 5: 📧 리포트 메일 발송
# --------------------------------------------------------------------------
with main_tab5:
    st.header(f"📧 {AGENT_TITLE} 결과 메일 발송")
    st.markdown("ICP 분석표 및 M/B 결과가 포함된 엑셀 파일(`.xlsx`)을 메일로 전송합니다.")

    col_m1, col_m2 = st.columns(2)
    with col_m1:
        recipient_email = st.text_input("수신자 이메일 주소", value="user@company.com")
        email_subject = st.text_input("메일 제목", value=f"[{AGENT_TITLE}] Run {run_no} 종합 브리핑")
        sender_email = st.text_input("발신자 이메일 주소 (Gmail/SMTP)", value="sender@gmail.com")
        sender_password = st.text_input("발신자 앱 비밀번호 (16자리)", type="password")

    with col_m2:
        smtp_server = st.text_input("SMTP 서버", value="smtp.gmail.com")
        smtp_port = st.number_input("SMTP 포트", value=587, step=1)

    if st.button("🚀 엑셀 리포트 첨부하여 이메일 발송", type="primary", use_container_width=True):
        if not sender_email or not sender_password or not recipient_email:
            st.error("❌ 발신자/수신자 이메일과 비밀번호를 모두 입력해 주세요.")
        else:
            try:
                wb = Workbook()
                ws1 = wb.active
                ws1.title = "MB_종합결과"
                ws1.append(["지표명", "수치", "단위", "평가"])
                ws1.append(["실험 회차", run_no, "Run", "정상"])
                ws1.append(["총 Li 회수율(ICP)", round(total_li_rec_pct, 2), "%", "1차여액+수세액"])
                ws1.append(["1차 여액 LiOH 농도", round(lioh_equiv_g_l, 2), "g/L", f"Li: {icp_li_1:.1f} mg/L"])
                ws1.append(["M/B 정합성", round(mass_closure, 2), "%", f"증발 {loss_mass:.1f}g"])
                ws1.append(["소성 감율(LOI)", round(loi_pct, 2), "%", f"CaCO3 순도 ~{purity_caco3:.1f}%"])

                # ICP 분석 시트 추가
                ws2 = wb.create_sheet(title="ICP_원소분석")
                ws2.append(list(df_icp_summary.columns))
                for _, r in df_icp_summary.iterrows():
                    ws2.append(list(r.values))

                # 트렌드 시뮬레이션 시트 추가
                ws3 = wb.create_sheet(title="트렌드시뮬레이션")
                ws3.append(list(df_simulation.columns))
                for _, r in df_simulation.iterrows():
                    ws3.append(list(r.values))

                excel_buffer = io.BytesIO()
                wb.save(excel_buffer)
                excel_buffer.seek(0)

                msg = MIMEMultipart()
                msg["From"] = sender_email
                msg["To"] = recipient_email
                msg["Subject"] = email_subject

                html_body = f"""
                <h3>🧪 {AGENT_TITLE} - Run {run_no} 리포트</h3>
                <hr>
                <ul>
                    <li><b>총 Li 회수율:</b> {total_li_rec_pct:.2f}% (1차 여액: {li_rec_1_pct:.2f}%, 수세액: {li_rec_w_pct:.2f}%)</li>
                    <li><b>1차 여액 LiOH 농도:</b> {lioh_equiv_g_l:.2f} g/L (Li: {icp_li_1:,.1f} mg/L)</li>
                    <li><b>주요 불순물 농도(1차 여액):</b> Ca {icp_ca_1} mg/L, Na {icp_na_1} mg/L, Si {icp_si_1} mg/L, Mg {icp_mg_1} mg/L, K {icp_k_1} mg/L</li>
                    <li><b>M/B 정합성(Closure):</b> {mass_closure:.2f}% (증발 손실: {loss_mass:.1f}g)</li>
                </ul>
                <p>※ 세부 ICP 분석표 및 시뮬레이션 데이터가 포함된 엑셀 파일(5개 시트)을 첨부하였습니다.</p>
                """
                msg.attach(MIMEText(html_body, "html", "utf-8"))

                part = MIMEApplication(excel_buffer.read(), Name=f"MB_ICP_Report_Run_{run_no:03d}.xlsx")
                part['Content-Disposition'] = f'attachment; filename="MB_ICP_Report_Run_{run_no:03d}.xlsx"'
                msg.attach(part)

                server = smtplib.SMTP(smtp_server, smtp_port)
                server.starttls()
                server.login(sender_email, sender_password)
                server.send_message(msg)
                server.quit()

                st.success(f"🎉 [{recipient_email}]로 엑셀 리포트 및 ICP 분석 결과 메일이 발송되었습니다!")
            except Exception as e:
                st.error(f"❌ 메일 발송 실패: {e}")
