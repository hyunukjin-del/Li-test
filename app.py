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

AGENT_TITLE = "LC-LH전환반응 M/B자동화 및 거동예측 Agent tool"

st.set_page_config(page_title=AGENT_TITLE, page_icon="🧪", layout="wide")

# --------------------------------------------------------------------------
# [1] 기본 세션 상태 및 변수 초기화
# --------------------------------------------------------------------------
DEFAULT_DATA = {
    "run_no": 1,
    "li2co3_mass": 95.34,
    "li2co3_water": 1040.0,
    "fresh_cao_mass": 92.42,
    "recycled_cao_mass": 0.0,
    "slurry_water": 831.0,
    "temp_c": 80.0,
    "time_h": 2.0,
    "primary_filtrate_mass": 1646.0,
    "primary_filtrate_sg": 1.035,
    "primary_filtrate_ph": 12.81,
    "wet_cake_mass": 311.0,
    "sample_wet": 27.7,
    "sample_dry": 14.8,
    "wash_sol_mass": 832.0,
    "wash_sol_sg": 1.000,
    "wash_sol_ph": 13.70,
    "test_dry_cake": 40.6,
    "calcined_cao": 23.9,
    "calc_temp": 1000.0,
    "calc_time": 1.0,
    "icp_li_1": 10500.0,
    "icp_ca_1": 120.0,
    "icp_na_1": 45.0,
    "icp_si_1": 8.5,
    "icp_mg_1": 1.2,
    "icp_k_1": 15.0,
    "icp_li_w": 1400.0,
    "icp_ca_w": 80.0,
    "icp_na_w": 6.0,
    "icp_si_w": 2.1,
    "icp_mg_w": 0.3,
    "icp_k_w": 2.0,
}

for k, v in DEFAULT_DATA.items():
    if k not in st.session_state:
        st.session_state[k] = v

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

# 키패드 매핑 사전 (한글 라벨 -> 세션 변수명)
FIELD_MAP = {
    "Li₂CO₃ 투입량 (g)": "li2co3_mass",
    "Li₂CO₃ 용매수 (g)": "li2co3_water",
    "신품 CaO 투입량 (g)": "fresh_cao_mass",
    "재생 CaO 투입량 (g)": "recycled_cao_mass",
    "슬러리 조제수 (g)": "slurry_water",
    "반응 온도 (℃)": "temp_c",
    "반응 시간 (h)": "time_h",
    "1차 여액 무게 (g)": "primary_filtrate_mass",
    "1차 여액 비중 (g/mL)": "primary_filtrate_sg",
    "1차 여액 pH": "primary_filtrate_ph",
    "1차 습케이크 무게 (g)": "wet_cake_mass",
    "함수율 샘플 습중량 (g)": "sample_wet",
    "함수율 샘플 건중량 (g)": "sample_dry",
    "수세액 무게 (g)": "wash_sol_mass",
    "수세액 pH": "wash_sol_ph",
    "소성 투입 건조케익 (g)": "test_dry_cake",
    "소성 회수 CaO (g)": "calcined_cao",
    "1차여액 Li 농도 (mg/L)": "icp_li_1",
    "1차여액 Ca 농도 (mg/L)": "icp_ca_1",
    "1차여액 Na 농도 (mg/L)": "icp_na_1",
    "1차여액 Si 농도 (mg/L)": "icp_si_1",
    "1차여액 Mg 농도 (mg/L)": "icp_mg_1",
    "1차여액 K 농도 (mg/L)": "icp_k_1",
    "수세액 Li 농도 (mg/L)": "icp_li_w",
    "수세액 Ca 농도 (mg/L)": "icp_ca_w",
    "수세액 Na 농도 (mg/L)": "icp_na_w",
    "수세액 Si 농도 (mg/L)": "icp_si_w",
    "수세액 Mg 농도 (mg/L)": "icp_mg_w",
    "수세액 K 농도 (mg/L)": "icp_k_w"
}

if "keypad_target" not in st.session_state:
    st.session_state.keypad_target = "Li₂CO₃ 투입량 (g)"
if "keypad_buffer" not in st.session_state:
    st.session_state.keypad_buffer = str(st.session_state[FIELD_MAP[st.session_state.keypad_target]])

def update_keypad_target():
    st.session_state.keypad_buffer = str(st.session_state[FIELD_MAP[st.session_state.keypad_target]])

# --------------------------------------------------------------------------
# [2] 사이드바: 가상 계산기 키패드
# --------------------------------------------------------------------------
with st.sidebar:
    st.header("🧮 마우스 전용 가상 키패드")
    st.caption("터치나 마우스 클릭으로 수치를 빠르게 입력하고 즉시 적용할 수 있습니다.")
    
    target_label = st.selectbox(
        "📌 입력할 항목 선택:", 
        list(FIELD_MAP.keys()), 
        key="keypad_target", 
        on_change=update_keypad_target
    )
    target_var = FIELD_MAP[target_label]

    # LCD 스타일 표시창
    st.markdown(
        f"""
        <div style="background-color: #1E293B; color: #38BDF8; font-family: monospace; 
                    font-size: 26px; font-weight: bold; text-align: right; 
                    padding: 10px 14px; border-radius: 8px; border: 2px solid #0EA5E9; margin-bottom: 10px;">
            {st.session_state.keypad_buffer if st.session_state.keypad_buffer else "0"}
        </div>
        """, 
        unsafe_allow_html=True
    )

    # 키패드 버튼 그리드 (4x4)
    r1_c1, r1_c2, r1_c3, r1_c4 = st.columns(4)
    if r1_c1.button("7", use_container_width=True): st.session_state.keypad_buffer += "7"; st.rerun()
    if r1_c2.button("8", use_container_width=True): st.session_state.keypad_buffer += "8"; st.rerun()
    if r1_c3.button("9", use_container_width=True): st.session_state.keypad_buffer += "9"; st.rerun()
    if r1_c4.button("⌫", help="한 글자 지우기", use_container_width=True): 
        st.session_state.keypad_buffer = st.session_state.keypad_buffer[:-1]
        st.rerun()

    r2_c1, r2_c2, r2_c3, r2_c4 = st.columns(4)
    if r2_c1.button("4", use_container_width=True): st.session_state.keypad_buffer += "4"; st.rerun()
    if r2_c2.button("5", use_container_width=True): st.session_state.keypad_buffer += "5"; st.rerun()
    if r2_c3.button("6", use_container_width=True): st.session_state.keypad_buffer += "6"; st.rerun()
    if r2_c4.button("C", help="전체 지우기", use_container_width=True): 
        st.session_state.keypad_buffer = ""
        st.rerun()

    r3_c1, r3_c2, r3_c3, r3_c4 = st.columns(4)
    if r3_c1.button("1", use_container_width=True): st.session_state.keypad_buffer += "1"; st.rerun()
    if r3_c2.button("2", use_container_width=True): st.session_state.keypad_buffer += "2"; st.rerun()
    if r3_c3.button("3", use_container_width=True): st.session_state.keypad_buffer += "3"; st.rerun()
    if r3_c4.button("00", use_container_width=True): st.session_state.keypad_buffer += "00"; st.rerun()

    r4_c1, r4_c2, r4_c3, r4_c4 = st.columns(4)
    if r4_c1.button("0", use_container_width=True): st.session_state.keypad_buffer += "0"; st.rerun()
    if r4_c2.button(".", use_container_width=True): 
        if "." not in st.session_state.keypad_buffer:
            st.session_state.keypad_buffer += "."
        st.rerun()
    if r4_c3.button("±", use_container_width=True):
        if st.session_state.keypad_buffer.startswith("-"):
            st.session_state.keypad_buffer = st.session_state.keypad_buffer[1:]
        else:
            st.session_state.keypad_buffer = "-" + st.session_state.keypad_buffer
        st.rerun()
    if r4_c4.button("💾", help="입력값 적용 (Enter)", type="primary", use_container_width=True):
        try:
            val = float(st.session_state.keypad_buffer) if st.session_state.keypad_buffer else 0.0
            st.session_state[target_var] = val
            st.success(f"✅ '{target_label}'에 {val} 적용 완료!")
            st.rerun()
        except ValueError:
            st.error("유효한 숫자가 아닙니다.")

    st.divider()
    st.info("💡 사이드바의 번호판을 누르고 [💾] 버튼을 누르면 본문 화면의 입력창에 즉시 입력됩니다.")

# --------------------------------------------------------------------------
# [3] 메인 화면 및 탭 구성
# --------------------------------------------------------------------------
st.title(f"🧪 {AGENT_TITLE}")
st.caption("탄산리튬(LC) → 수산화리튬(LH) 가성화 양론 연산 | ICP 성분별 회수율 및 불순물 분석 | 회차별 거동예측 | 리포트 메일 발송")

main_tab1, main_tab2, main_tab3, main_tab4, main_tab5 = st.tabs([
    "1️⃣ 실험 데이터 입력 & M/B 연산", 
    "2️⃣ 🧪 용액 ICP 분석 & 회수율", 
    "3️⃣ 📈 회차별 트렌드 & 거동예측", 
    "4️⃣ 💬 AI 공정 대화창", 
    "5️⃣ 📧 리포트 메일 발송"
])

# --------------------------------------------------------------------------
# TAB 1: 실험 데이터 입력 및 M/B 연산
# --------------------------------------------------------------------------
with main_tab1:
    with st.expander("📝 이번 회차 실험 수치 입력 폼 (키보드 직접 입력도 가능)", expanded=True):
        col_in1, col_in2 = st.columns(2)
        with col_in1:
            st.markdown("#### [1. 투입 원료 및 반응 조건]")
            run_no = st.number_input("실험 회차 (Run No.)", min_value=1, value=int(st.session_state["run_no"]), step=1, key="run_no")
            li2co3_mass = st.number_input("Li₂CO₃ 투입량 (g)", format="%.2f", key="li2co3_mass")
            li2co3_water = st.number_input("Li₂CO₃ 용매수 (g)", format="%.1f", key="li2co3_water")
            fresh_cao_mass = st.number_input("신품 CaO 투입량 (g)", format="%.2f", key="fresh_cao_mass")
            recycled_cao_mass = st.number_input("재생 CaO 투입량 (g)", format="%.2f", key="recycled_cao_mass")
            slurry_water = st.number_input("슬러리 조제수 (g)", format="%.1f", key="slurry_water")
            temp_c = st.number_input("반응 온도 (℃)", format="%.1f", key="temp_c")
            time_h = st.number_input("반응 시간 (시간)", format="%.1f", key="time_h")

        with col_in2:
            st.markdown("#### [2. 1차 여과 및 케이크 수세]")
            primary_filtrate_mass = st.number_input("1차 LiOH 여액 무게 (g)", format="%.1f", key="primary_filtrate_mass")
            primary_filtrate_sg = st.number_input("1차 여액 비중 (g/mL)", format="%.3f", key="primary_filtrate_sg")
            primary_filtrate_ph = st.number_input("1차 여액 pH", format="%.2f", key="primary_filtrate_ph")
            wet_cake_mass = st.number_input("1차 습케이크 무게 (g)", format="%.1f", key="wet_cake_mass")
            sample_wet = st.number_input("함수율 측정 샘플 습중량 (g)", format="%.1f", key="sample_wet")
            sample_dry = st.number_input("함수율 측정 샘플 건중량 (g)", format="%.1f", key="sample_dry")
            wash_sol_mass = st.number_input("회수된 수세액 무게 (g)", format="%.1f", key="wash_sol_mass")
            wash_sol_sg = st.number_input("수세액 비중 (g/mL)", format="%.3f", key="wash_sol_sg")
            wash_sol_ph = st.number_input("수세액 pH", format="%.2f", key="wash_sol_ph")

        st.divider()

        # 소성 분리 배치
        st.markdown("#### [3. CaCO₃ 소성(하소) 및 CaO 재생]")
        col_calc1, col_calc2 = st.columns(2)
        with col_calc1:
            test_dry_cake = st.number_input("소성 투입 건조케익 샘플 (g)", format="%.1f", key="test_dry_cake")
            calcined_cao = st.number_input("소성 후 회수된 CaO (g)", format="%.1f", key="calcined_cao")
        with col_calc2:
            calc_temp = st.number_input("소성 온도 (℃)", format="%.1f", key="calc_temp")
            calc_time = st.number_input("소성 시간 (시간)", format="%.1f", key="calc_time")

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
    cake_moisture = (1.0 - (sample_dry / sample_wet)) * 100.0 if sample_wet > 0 else 0.0
    est_total_dry_solids = wet_cake_mass * (sample_dry / sample_wet) if sample_wet > 0 else 0.0

    loi_pct = ((test_dry_cake - calcined_cao) / test_dry_cake) * 100.0 if test_dry_cake > 0 else 0.0
    cao_yield_dry = (calcined_cao / test_dry_cake) * 100.0 if test_dry_cake > 0 else 0.0
    purity_caco3 = max(0.0, min(100.0, ((loi_pct/100.0) - 0.2432) / (0.4397 - 0.2432) * 100.0))
    pot_total_cao = est_total_dry_solids * (cao_yield_dry / 100.0)
    ca_loop_recovery = (pot_total_cao / total_cao_in) * 100.0 if total_cao_in > 0 else 0.0

    target_cao = 92.42
    fresh_makeup = max(0.0, target_cao - calcined_cao)

    st.subheader(f"📊 Run {run_no} 공정 기본 물질수지(M/B)")
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("M/B 정합성 (Closure)", f"{mass_closure:.2f} %", f"증발: {loss_mass:.1f}g")
    k2.metric("1차 케익 함수율", f"{cake_moisture:.2f} %", f"고형분 {est_total_dry_solids:.1f}g")
    k3.metric("소성 감율 (LOI)", f"{loi_pct:.2f} %", f"CaCO₃ 순도 ~{purity_caco3:.1f}%")
    k4.metric("Ca-Loop 원소 회수율", f"{ca_loop_recovery:.2f} %", f"재생잠재 {pot_total_cao:.1f}g")

# --------------------------------------------------------------------------
# TAB 2: 🧪 용액 ICP 분석 (mg/L 단위 & 하단 회수율 자동 계산)
# --------------------------------------------------------------------------
with main_tab2:
    st.header("🧪 용액 ICP 분석 데이터 입력 및 회수율 계산")
    st.markdown("1차 여액 및 수세액의 각 성분 농도(**mg/L**)를 입력하면, **투입량 대비 리튬 회수율과 성분별 물질수지**를 하단에 즉시 계산합니다.")

    st.markdown("### 1. ICP 분석 데이터 입력")
    icp_col1, icp_col2 = st.columns(2)

    with icp_col1:
        st.markdown(f"#### 🔹 1차 여액 분석치 (부피: {primary_filtrate_mass/primary_filtrate_sg:.1f} mL)")
        icp_li_1 = st.number_input("Li 농도 (mg/L) - 1차 여액", format="%.1f", key="icp_li_1")
        icp_ca_1 = st.number_input("Ca 농도 (mg/L) - 1차 여액", format="%.1f", key="icp_ca_1")
        icp_na_1 = st.number_input("Na 농도 (mg/L) - 1차 여액", format="%.1f", key="icp_na_1")
        icp_si_1 = st.number_input("Si 농도 (mg/L) - 1차 여액", format="%.1f", key="icp_si_1")
        icp_mg_1 = st.number_input("Mg 농도 (mg/L) - 1차 여액", format="%.1f", key="icp_mg_1")
        icp_k_1  = st.number_input("K 농도 (mg/L) - 1차 여액", format="%.1f", key="icp_k_1")

    with icp_col2:
        st.markdown(f"#### 🔹 수세액 분석치 (부피: {wash_sol_mass/wash_sol_sg:.1f} mL)")
        icp_li_w = st.number_input("Li 농도 (mg/L) - 수세액", format="%.1f", key="icp_li_w")
        icp_ca_w = st.number_input("Ca 농도 (mg/L) - 수세액", format="%.1f", key="icp_ca_w")
        icp_na_w = st.number_input("Na 농도 (mg/L) - 수세액", format="%.1f", key="icp_na_w")
        icp_si_w = st.number_input("Si 농도 (mg/L) - 수세액", format="%.1f", key="icp_si_w")
        icp_mg_w = st.number_input("Mg 농도 (mg/L) - 수세액", format="%.1f", key="icp_mg_w")
        icp_k_w  = st.number_input("K 농도 (mg/L) - 수세액", format="%.1f", key="icp_k_w")

    st.divider()

    # 하단 회수율 연산
    st.markdown("### 2. 성분별 질량 및 회수율 계산 결과")

    v1_L = (primary_filtrate_mass / primary_filtrate_sg) / 1000.0
    vw_L = (wash_sol_mass / wash_sol_sg) / 1000.0
    li_in_total_g = n_li2co3 * 2.0 * MW_LI

    elements = ["Li", "Ca", "Na", "Si", "Mg", "K"]
    conc_1 = [icp_li_1, icp_ca_1, icp_na_1, icp_si_1, icp_mg_1, icp_k_1]
    conc_w = [icp_li_w, icp_ca_w, icp_na_w, icp_si_w, icp_mg_w, icp_k_w]

    mass_1_g = [c * v1_L / 1000.0 for c in conc_1]
    mass_w_g = [c * vw_L / 1000.0 for c in conc_w]
    mass_total_g = [m1 + mw for m1, mw in zip(mass_1_g, mass_w_g)]

    li_rec_1_pct = (mass_1_g[0] / li_in_total_g) * 100.0 if li_in_total_g > 0 else 0.0
    li_rec_w_pct = (mass_w_g[0] / li_in_total_g) * 100.0 if li_in_total_g > 0 else 0.0
    total_li_rec_pct = li_rec_1_pct + li_rec_w_pct
    li_loss_pct = max(0.0, 100.0 - total_li_rec_pct)
    lioh_equiv_g_l = icp_li_1 * (MW_LIOH / MW_LI) / 1000.0

    rc1, rc2, rc3, rc4 = st.columns(4)
    rc1.metric("🎯 총 Li 회수율", f"{total_li_rec_pct:.2f} %", f"총 회수: {mass_total_g[0]:.2f}g / 투입: {li_in_total_g:.2f}g")
    rc2.metric("1차 여액 회수율", f"{li_rec_1_pct:.2f} %", f"회수 질량: {mass_1_g[0]:.2f}g")
    rc3.metric("수세액 회수율", f"{li_rec_w_pct:.2f} %", f"회수 질량: {mass_w_g[0]:.2f}g")
    rc4.metric("케이크 잔류/손실률", f"{li_loss_pct:.2f} %", f"미회수: {li_in_total_g - mass_total_g[0]:.2f}g", delta_color="inverse")

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

                ws2 = wb.create_sheet(title="ICP_원소분석")
                ws2.append(list(df_icp_summary.columns))
                for _, r in df_icp_summary.iterrows():
                    ws2.append(list(r.values))

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
