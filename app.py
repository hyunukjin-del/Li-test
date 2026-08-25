import streamlit as st
import pandas as pd
import io
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

st.set_page_config(page_title="습식반응 M/B 자동화 Agent", page_icon="🧪", layout="wide")

st.title("🧪 습식반응 및 Ca-Loop M/B 자동화 Agent")
st.markdown("수치를 입력하면 **화학양론 계산, Mass Balance 정합성 검증, AI 공정 진단 및 엑셀 리포트**가 자동 생성됩니다.")

# 사이드바: 기본 설정
with st.sidebar:
    st.header("⚙️ 실험 회차 설정")
    run_no = st.number_input("실험 회차 (Run No.)", min_value=1, value=1, step=1)
    st.info("💡 이전 회차의 재생 CaO를 투입하는 경우 아래 원료 탭에서 입력하세요.")

# 메인 입력 폼 (탭 구분)
tab1, tab2, tab3, tab4 = st.tabs(["1️⃣ 투입 원료", "2️⃣ 여과 및 수세", "3️⃣ 소성(재생)", "4️⃣ ICP 분석(선택)"])

with tab1:
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("탄산리튬 용액")
        li2co3_mass = st.number_input("Li₂CO₃ 투입량 (g)", value=95.34, format="%.2f")
        li2co3_water = st.number_input("Li₂CO₃ 용매수 (g)", value=1040.0, format="%.1f")
        temp_c = st.number_input("반응 온도 (℃)", value=80.0, format="%.1f")
        time_h = st.number_input("반응 시간 (시간)", value=2.0, format="%.1f")
    with col2:
        st.subheader("CaO 슬러리")
        fresh_cao_mass = st.number_input("신품 CaO 투입량 (g)", value=92.42, format="%.2f")
        recycled_cao_mass = st.number_input("재생 CaO 투입량 (g)", value=0.0, format="%.2f")
        slurry_water = st.number_input("슬러리 조제수 (g)", value=831.0, format="%.1f")
        cao_purity = st.slider("CaO 순도 (%)", min_value=80.0, max_value=100.0, value=100.0) / 100.0

with tab2:
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("1차 반응 및 여과")
        primary_filtrate_mass = st.number_input("1차 LiOH 여액 무게 (g)", value=1646.0, format="%.1f")
        primary_filtrate_sg = st.number_input("1차 여액 비중 (g/mL)", value=1.035, format="%.3f")
        primary_filtrate_ph = st.number_input("1차 여액 pH", value=12.81, format="%.2f")
        wet_cake_mass = st.number_input("1차 습케이크 무게 (g)", value=311.0, format="%.1f")
        sample_wet = st.number_input("함수율 측정 샘플 습중량 (g)", value=27.7, format="%.1f")
        sample_dry = st.number_input("함수율 측정 샘플 건중량 (g)", value=14.8, format="%.1f")
    with col2:
        st.subheader("케이크 3배수 수세")
        wash_input_cake = st.number_input("수세 투입 습케이크 (g)", value=283.0, format="%.1f")
        washed_dry_cake = st.number_input("수세 후 최종 건조케이크 (g)", value=141.0, format="%.1f")
        wash_sol_mass = st.number_input("회수된 수세액 무게 (g)", value=832.0, format="%.1f")
        wash_sol_sg = st.number_input("수세액 비중 (g/mL)", value=1.000, format="%.3f")
        wash_sol_ph = st.number_input("수세액 pH", value=13.70, format="%.2f")

with tab3:
    st.subheader("CaCO₃ 소성 (1000℃ 하소 및 CaO 재생)")
    col1, col2 = st.columns(2)
    with col1:
        test_dry_cake = st.number_input("소성 테스트 건조케이크 샘플 (g)", value=40.6, format="%.1f")
        calcined_cao = st.number_input("소성 후 회수된 CaO (g)", value=23.9, format="%.1f")
    with col2:
        calc_temp = st.number_input("소성 온도 (℃)", value=1000.0, format="%.1f")
        calc_time = st.number_input("소성 시간 (시간)", value=1.0, format="%.1f")

with tab4:
    st.subheader("ICP 분석 농도 (mg/L)")
    col1, col2 = st.columns(2)
    with col1:
        primary_li = st.number_input("1차 여액 Li 농도 (mg/L)", value=10500.0, format="%.1f")
        primary_ca = st.number_input("1차 여액 Ca 농도 (mg/L)", value=120.0, format="%.1f")
    with col2:
        wash_li = st.number_input("수세액 Li 농도 (mg/L)", value=1400.0, format="%.1f")
        wash_ca = st.number_input("수세액 Ca 농도 (mg/L)", value=80.0, format="%.1f")

st.divider()

# 연산 실행
if st.button("🚀 M/B 계산 및 AI 공정 진단 실행", type="primary", use_container_width=True):
    # 1. 양론 계산
    n_li2co3 = li2co3_mass / MW_LI2CO3
    total_cao_in = fresh_cao_mass + recycled_cao_mass
    n_cao = (total_cao_in * cao_purity) / MW_CAO
    limiting = "Li2CO3" if n_li2co3 <= n_cao else "CaO"
    excess_pct = (n_cao / n_li2co3 - 1.0) * 100
    theo_lioh_mass = (n_li2co3 * 2.0) * MW_LIOH
    
    # 2. 물질수지
    total_in = li2co3_mass + li2co3_water + total_cao_in + slurry_water
    total_out = primary_filtrate_mass + wet_cake_mass
    loss_mass = total_in - total_out
    mass_closure = (total_out / total_in) * 100.0
    cake_moisture = (1.0 - (sample_dry / sample_wet)) * 100.0
    est_total_dry_solids = wet_cake_mass * (sample_dry / sample_wet)

    # 3. 소성 및 재생
    loi_pct = ((test_dry_cake - calcined_cao) / test_dry_cake) * 100.0
    cao_yield_dry = (calcined_cao / test_dry_cake) * 100.0
    purity_caco3 = max(0.0, min(100.0, ((loi_pct/100.0) - 0.2432) / (0.4397 - 0.2432) * 100.0))
    pot_total_cao = est_total_dry_solids * (cao_yield_dry / 100.0)
    ca_loop_recovery = (pot_total_cao / total_cao_in) * 100.0

    # 4. Li 회수율
    li_in_g = n_li2co3 * 2 * MW_LI
    v_primary_l = (primary_filtrate_mass / primary_filtrate_sg) / 1000.0
    v_wash_l = (wash_sol_mass / wash_sol_sg) / 1000.0
    li_p_g = (primary_li * v_primary_l) / 1000.0
    li_w_g = (wash_li * v_wash_l) / 1000.0
    total_li_rec_g = li_p_g + li_w_g
    li_recovery_pct = (total_li_rec_g / li_in_g) * 100.0 if primary_li > 0 else 0.0

    # 5. 차기 배합비
    target_cao = 92.42
    fresh_makeup = max(0.0, target_cao - calcined_cao)

    # 결과 표시 대시보드
    st.header(f"📊 Run {run_no} M/B 연산 결과")
    kpi1, kpi2, kpi3, kpi4 = st.columns(4)
    kpi1.metric("M/B 정합성 (Closure)", f"{mass_closure:.2f} %", f"손실: {loss_mass:.1f}g", delta_color="normal")
    kpi2.metric("총 Li 회수율 (ICP)", f"{li_recovery_pct:.2f} %", f"여액 {li_p_g:.1f}g + 수세 {li_w_g:.1f}g")
    kpi3.metric("소성 감율 (LOI)", f"{loi_pct:.2f} %", f"CaCO₃ 순도 ~{purity_caco3:.1f}%")
    kpi4.metric("Ca-Loop 원소 회수율", f"{ca_loop_recovery:.2f} %", f"재생잠재 {pot_total_cao:.1f}g")

    # AI 공정 진단 카드
    st.subheader("🤖 AI 공정 진단 및 이상징후 체크포인트")
    if loss_mass > 50.0:
        st.warning(f"⚠️ **[증발 손실]** 반응 중 약 **{loss_mass:.1f}g**의 수분이 증발했습니다. 농도 일정 유지를 위해 환류 냉각기 장착 또는 증발 보충수(Make-up water) 투입을 권장합니다.")
    if wash_sol_ph >= 13.5:
        st.info(f"💡 **[케이크 잔류 용질 회수]** 수세액 pH가 **{wash_sol_ph:.2f}**로 매우 높습니다. 1차 여과 후 케이크 기공 내 LiOH 농축액이 다량 잔류했음을 의미하며, 3배수 수세로 성공적으로 회수되었습니다.")
    if loi_pct < 43.0:
        st.info(f"🔍 **[케이크 순도 역산]** 하소 감율({loi_pct:.1f}%)이 순수 탄산칼슘(43.97%)보다 낮습니다. 반응 시 과잉 투입된 Ca(OH)₂가 케이크 내 약 **{100-purity_caco3:.1f}%** 공침/잔류한 결과입니다.")
    if calc_temp >= 1000.0:
        st.warning(f"⚠️ **[소결 주의]** 1000℃ 고온 소성된 재생 CaO는 비표면적이 감소하여 소화 속도가 지연될 수 있습니다. Run {run_no+1} 슬러리 조제 시 교반 시간을 15분 이상 충분히 확보하세요.")

    # 차기 회차 가이드
    st.subheader(f"📋 차기 회차 (Run {run_no+1}) 추천 배합비")
    guide_col1, guide_col2, guide_col3 = st.columns(3)
    guide_col1.metric("재생 CaO 사용량", f"{calcined_cao:.2f} g")
    guide_col2.metric("신품 CaO 보충량 (Make-up)", f"{fresh_makeup:.2f} g")
    guide_col3.metric("슬러리 물 투입량", f"{slurry_water:.1f} g")

    # 엑셀 파일 생성 (메모리 버퍼)
    wb = Workbook()
    ws_calc = wb.active
    ws_calc.title = "MB계산결과"
    ws_calc.append(["대분류", "지표명", "계산값", "단위", "평가/비고"])
    ws_calc.append(["화학양론", "제한반응물", limiting, "-", f"Li2CO3 {n_li2co3:.2f}mol vs CaO {n_cao:.2f}mol"])
    ws_calc.append(["화학양론", "이론 LiOH 생성량", round(theo_lioh_mass, 2), "g", "100% 전환 기준"])
    ws_calc.append(["물질수지", "M/B 정합성(Closure)", round(mass_closure, 2), "%", f"증발손실: {loss_mass:.1f}g"])
    ws_calc.append(["소성/재생", "실측 하소 감율(LOI)", round(loi_pct, 2), "%", f"CaCO3 순도 약 {purity_caco3:.1f}%"])
    ws_calc.append(["ICP/수율", "총 Li 회수율", round(li_recovery_pct, 2), "%", f"1차여액 {li_p_g:.2f}g + 수세액 {li_w_g:.2f}g"])
    ws_calc.append(["차기투입", "Run n+1 신품 CaO 보충량", round(fresh_makeup, 2), "g", f"재생분 {calcined_cao:.2f}g 활용"])

    excel_buffer = io.BytesIO()
    wb.save(excel_buffer)
    excel_buffer.seek(0)

    st.download_button(
        label=f"📥 Run {run_no} 엑셀 리포트 다운로드 (.xlsx)",
        data=excel_buffer,
        file_name=f"MB_Report_Run_{run_no:03d}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary"
    )
