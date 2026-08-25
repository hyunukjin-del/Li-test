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

st.set_page_config(page_title="습식반응 M/B & 트렌드 예측 Agent", page_icon="🧪", layout="wide")

# 세션 상태 초기화 (실측 데이터 이력 누적)
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
        {"role": "assistant", "content": "안녕하세요! 습식 가성화 및 Ca-Loop 공정 전문 AI 에이전트입니다. 이번 회차 계산 결과, $n$회차 농도/회수율 예측, 공정 변수 최적화에 대해 무엇이든 물어보세요."}
    ]

st.title("🧪 습식반응 및 Ca-Loop M/B 자동화 & 트렌드 예측 Agent")
st.caption("화학양론 M/B 연산 | 회차별 농도·회수율 예측 시뮬레이션 | AI 공정 진단 챗봇 | 리포트 메일 발송")

# 상단 메인 탭
main_tab1, main_tab2, main_tab3, main_tab4 = st.tabs([
    "1️⃣ 실험 데이터 입력 & M/B 연산", 
    "2️⃣ 📈 회차별 트렌드 & 가상 예측", 
    "3️⃣ 💬 AI 공정 대화창", 
    "4️⃣ 📧 리포트 메일 발송"
])

# --------------------------------------------------------------------------
# TAB 1: 실험 데이터 입력 및 M/B 연산
# --------------------------------------------------------------------------
with main_tab1:
    with st.expander("📝 이번 회차 실험 수치 입력 폼", expanded=True):
        col_in1, col_in2 = st.columns(2)
        with col_in1:
            st.markdown("#### [투입 원료 및 반응 조건]")
            run_no = st.number_input("실험 회차 (Run No.)", min_value=1, value=len(st.session_state.history), step=1)
            li2co3_mass = st.number_input("Li₂CO₃ 투입량 (g)", value=95.34, format="%.2f")
            li2co3_water = st.number_input("Li₂CO₃ 용매수 (g)", value=1040.0, format="%.1f")
            fresh_cao_mass = st.number_input("신품 CaO 투입량 (g)", value=92.42 if run_no == 1 else 68.52, format="%.2f")
            recycled_cao_mass = st.number_input("재생 CaO 투입량 (g)", value=0.0 if run_no == 1 else 23.90, format="%.2f")
            slurry_water = st.number_input("슬러리 조제수 (g)", value=831.0, format="%.1f")
            temp_c = st.number_input("반응 온도 (℃)", value=80.0, format="%.1f")
            time_h = st.number_input("반응 시간 (시간)", value=2.0, format="%.1f")

        with col_in2:
            st.markdown("#### [여과, 수세 및 소성(하소)]")
            primary_filtrate_mass = st.number_input("1차 LiOH 여액 무게 (g)", value=1646.0, format="%.1f")
            primary_filtrate_sg = st.number_input("1차 여액 비중 (g/mL)", value=1.035, format="%.3f")
            primary_filtrate_ph = st.number_input("1차 여액 pH", value=12.81, format="%.2f")
            wet_cake_mass = st.number_input("1차 습케이크 무게 (g)", value=311.0, format="%.1f")
            sample_wet = st.number_input("함수율 측정 샘플 습중량 (g)", value=27.7, format="%.1f")
            sample_dry = st.number_input("함수율 측정 샘플 건중량 (g)", value=14.8, format="%.1f")
            wash_sol_mass = st.number_input("수세액 무게 (g)", value=832.0, format="%.1f")
            wash_sol_ph = st.number_input("수세액 pH", value=13.70, format="%.2f")
            test_dry_cake = st.number_input("소성 투입 건조케익 (g)", value=40.6, format="%.1f")
            calcined_cao = st.number_input("소성 회수 CaO (g)", value=23.9, format="%.1f")

        st.markdown("#### [ICP 분석 농도 (mg/L)]")
        col_icp1, col_icp2 = st.columns(2)
        with col_icp1:
            primary_li = st.number_input("1차 여액 Li 농도 (mg/L)", value=10500.0, format="%.1f")
        with col_icp2:
            wash_li = st.number_input("수세액 Li 농도 (mg/L)", value=1400.0, format="%.1f")

    # M/B 연산 로직
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

    li_in_g = n_li2co3 * 2 * MW_LI
    v_primary_l = (primary_filtrate_mass / primary_filtrate_sg) / 1000.0
    v_wash_l = (wash_sol_mass / 1.0) / 1000.0
    li_p_g = (primary_li * v_primary_l) / 1000.0
    li_w_g = (wash_li * v_wash_l) / 1000.0
    total_li_rec_g = li_p_g + li_w_g
    li_recovery_pct = (total_li_rec_g / li_in_g) * 100.0 if primary_li > 0 else 0.0

    lioh_conc_g_l = primary_li * (MW_LIOH / MW_LI) / 1000.0

    target_cao = 92.42
    fresh_makeup = max(0.0, target_cao - calcined_cao)

    st.subheader(f"📊 Run {run_no} M/B 연산 및 진단 결과")
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("M/B 정합성 (Closure)", f"{mass_closure:.2f} %", f"증발: {loss_mass:.1f}g")
    k2.metric("총 Li 회수율 (ICP)", f"{li_recovery_pct:.2f} %", f"여액 {li_p_g:.1f}g + 수세 {li_w_g:.1f}g")
    k3.metric("1차 여액 LiOH 농도", f"{lioh_conc_g_l:.2f} g/L", f"Li: {primary_li:.0f} mg/L")
    k4.metric("Ca-Loop 원소 회수율", f"{ca_loop_recovery:.2f} %", f"재생잠재 {pot_total_cao:.1f}g")

    # AI 진단 알림
    diagnostics = []
    if loss_mass > 50.0:
        diagnostics.append(f"반응 중 약 {loss_mass:.1f}g의 수분 증발 손실 발생 (환류 냉각기 또는 보충수 투입 권장)")
    if wash_sol_ph >= 13.5:
        diagnostics.append(f"수세액 pH({wash_sol_ph:.2f}) 강알칼리 확인: 케이크 내부 잔류 LiOH 농축액이 세척수로 정상 회수됨")
    if loi_pct < 43.0:
        diagnostics.append(f"하소 감율({loi_pct:.1f}%) 편차: 잉여 Ca(OH)₂ 약 {100-purity_caco3:.1f}% 공침/잔류 확인")
    if temp_c >= 80.0:
        diagnostics.append("1000℃ 고온 소성에 따른 입자 소결(Sintering) 주의: 다음 회차 소화 시간 15분 이상 확보 권장")

    st.markdown("##### 🤖 AI 공정 진단 코멘트")
    for d in diagnostics:
        st.info(f"💡 {d}")

    # 현재 결과를 트렌드 DB에 반영
    if st.button("💾 이번 회차 실측치를 트렌드 DB에 저장", type="primary"):
        new_row = {
            "회차 (Run)": int(run_no),
            "구분": "실측치 (Actual)",
            "Li 회수율 (%)": round(li_recovery_pct, 2), 
            "1차여액 Li농도 (mg/L)": round(primary_li, 1),
            "1차여액 LiOH농도 (g/L)": round(lioh_conc_g_l, 2),
            "M/B 닫힘율 (%)": round(mass_closure, 2), 
            "하소 감율 LOI (%)": round(loi_pct, 2), 
            "CaO 활성도 (%)": 100.0,
            "신품 CaO 보충량 (g)": round(fresh_makeup, 2)
        }
        st.session_state.history = st.session_state.history[st.session_state.history["회차 (Run)"] != run_no]
        st.session_state.history = pd.concat([st.session_state.history, pd.DataFrame([new_row])]).sort_values("회차 (Run)").reset_index(drop=True)
        st.success(f"✅ Run {run_no} 실측 데이터가 트렌드 DB에 등록되었습니다!")

# --------------------------------------------------------------------------
# TAB 2: 사용자 지정 임의 변수 기반 트렌드 예측 & X-Y 시각화
# --------------------------------------------------------------------------
with main_tab2:
    st.header("📈 $n$회차 트렌드 시각화 & 임의 조건 예측 시뮬레이터")
    st.markdown("임의의 공정 파라미터(소결율, 반응온도, 목표 회차)를 입력하여 **미래 회차의 농도 및 회수율 변화를 예측 시뮬레이션**합니다.")

    # [1] 시뮬레이션 파라미터 설정 패널
    with st.expander("⚙️ 임의 조건 예측 파라미터 설정 (What-If Simulation)", expanded=True):
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

    # [2] 예측 데이터프레임 자동 생성
    sim_rows = []
    base_li_conc = 10500.0  # Run 1 기준 농도
    base_recovery = 95.80   # Run 1 기준 회수율

    # 온도/시간에 따른 반응성 보정 팩터
    temp_factor = 1.0 + (sim_temp - 80.0) * 0.004
    time_factor = 1.0 + (sim_time - 2.0) * 0.03

    for r in range(1, target_max_run + 1):
        # 기존 실측치가 있는 회차인지 확인
        matched_actual = st.session_state.history[st.session_state.history["회차 (Run)"] == r]
        
        if not matched_actual.empty:
            # 실측 데이터 유지
            row = matched_actual.iloc[0].to_dict()
            sim_rows.append(row)
        else:
            # AI 예측 모델 연산
            # 활성도: (1 - 감쇄율)^(r-1)
            activity = 100.0 * ((1.0 - (sintering_decay / 100.0)) ** (r - 1))
            eff_activity = min(100.0, activity * temp_factor * time_factor)
            
            # 예측 회수율 및 농도
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

    # [3] X축(회차) - Y축 지표 선택 시각화
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

    # 차트 데이터 재구성 (X축: 회차 (Run))
    chart_df = df_simulation.set_index("회차 (Run)")[[y_axis_metric]]

    # Streamlit 인터랙티브 라인 차트 렌더링
    st.line_chart(chart_df, height=380, use_container_width=True)

    # [4] 종합 비교 및 예측 데이터 테이블
    st.markdown("##### 📋 회차별 실측치 & 예측 시뮬레이션 상세 데이터 테이블")
    
    # 구분 컬럼 스타일링
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

    # [5] AI 분석 요약 인사이트
    st.markdown("##### 🤖 시뮬레이션 기반 AI 공정 제언")
    drop_run = df_simulation[df_simulation["Li 회수율 (%)"] < 90.0]
    if not drop_run.empty:
        first_drop = drop_run.iloc[0]["회차 (Run)"]
        st.warning(f"⚠️ 설정하신 조건(감쇄율 {sintering_decay}%) 적용 시, **Run {int(first_drop)}부터 Li 회수율이 90% 이하로 저하**될 것으로 예측됩니다. Run {int(first_drop)} 이전 시점에 반응 온도를 {sim_temp+5:.0f}℃로 승온하거나 CaO 퍼지(전량 교체)를 권고합니다.")
    else:
        st.success(f"✅ Run 1부터 Run {target_max_run}까지 모든 회차에서 Li 회수율이 90% 이상을 유지하는 안정적인 공정 조건입니다.")

# --------------------------------------------------------------------------
# TAB 3: AI 공정 엔지니어 대화창 (Chatbot)
# --------------------------------------------------------------------------
with main_tab3:
    st.header("💬 AI 공정 엔지니어와 대화하기")
    st.caption("현재 실험 수치 및 시뮬레이션 트렌드를 바탕으로 실시간 질의응답을 진행합니다.")

    for msg in st.session_state.chat_messages:
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])

    if user_prompt := st.chat_input("질문을 입력하세요 (예: 회수율을 95% 이상 유지하려면 몇 회차에 CaO를 교체해야 해?)"):
        st.session_state.chat_messages.append({"role": "user", "content": user_prompt})
        with st.chat_message("user"):
            st.markdown(user_prompt)

        prompt_low = user_prompt.lower()
        if "ph" in prompt_low or "수세" in prompt_low or "세척" in prompt_low:
            ai_reply = f"현재 1차 여액 pH는 {primary_filtrate_ph}이지만, 수세액 pH가 **{wash_sol_ph}**로 측정되었습니다. 이는 1차 여과 후 케이크 내부에 고농도 LiOH 액이 물리적으로 갇혀(Entrainment) 있었기 때문이며, 3배수 수세를 통해 리튬 성분이 정상 회수되었음을 의미합니다."
        elif "소결" in prompt_low or "sintering" in prompt_low or "교체" in prompt_low or "퍼지" in prompt_low:
            ai_reply = f"시뮬레이션 모델에 따르면 회차당 약 {sintering_decay}%의 소결 감쇄가 발생합니다. 회수율 저하를 막기 위해 약 5~7회차 주기로 재생 CaO의 일부를 폐기하고 신품 CaO로 100% 교체(Purge)하거나 소화 시간을 20분 이상 연장하는 것을 추천합니다."
        elif "농도" in prompt_low or "회수율" in prompt_low or "예측" in prompt_low:
            ai_reply = f"현재 설정된 조건 기준 1차 여액의 이론 LiOH 농도는 약 **{lioh_conc_g_l:.2f} g/L (Li 기준 {primary_li:.0f} mg/L)**입니다. 회차가 누적되면서 소결로 인해 농도가 소폭 감소할 수 있으나, 반응 온도를 85℃로 상향 시 약 1.5%의 회수율 보상 효과가 있습니다."
        elif "배합" in prompt_low or "보충" in prompt_low:
            ai_reply = f"다음 회차에서는 소성 회수된 **재생 CaO {calcined_cao:.2f}g**에 **신품 CaO {fresh_makeup:.2f}g**을 보충하여 총 92.42g을 맞추시면 됩니다."
        else:
            ai_reply = f"입력하신 실험 조건(Run {run_no}, Li 회수율 {li_recovery_pct:.1f}%, LiOH 농도 {lioh_conc_g_l:.2f} g/L)을 바탕으로 분석한 결과, 공정이 양호한 양론 밸런스를 유지하고 있습니다. 추가로 확인하고 싶은 세부 반응 변수가 있으신가요?"

        st.session_state.chat_messages.append({"role": "assistant", "content": ai_reply})
        with st.chat_message("assistant"):
            st.markdown(ai_reply)

# --------------------------------------------------------------------------
# TAB 4: 리포트 이메일 발송
# --------------------------------------------------------------------------
with main_tab4:
    st.header("📧 M/B 엑셀 리포트 및 시뮬레이션 결과 메일 발송")
    st.markdown("실험 결과 요약문과 생성된 엑셀 리포트 파일(`.xlsx`)을 담당자 이메일로 전송합니다.")

    col_m1, col_m2 = st.columns(2)
    with col_m1:
        recipient_email = st.text_input("수신자 이메일 주소", value="user@company.com")
        email_subject = st.text_input("메일 제목", value=f"[M/B 리포트] Run {run_no} 습식반응 및 Ca-Loop 트렌드 브리핑")
        sender_email = st.text_input("발신자 이메일 주소 (Gmail/SMTP)", value="sender@gmail.com")
        sender_password = st.text_input("발신자 앱 비밀번호 (16자리)", type="password")

    with col_m2:
        smtp_server = st.text_input("SMTP 서버 호스트", value="smtp.gmail.com")
        smtp_port = st.number_input("SMTP 포트", value=587, step=1)

    if st.button("🚀 엑셀 리포트 첨부하여 이메일 발송", type="primary", use_container_width=True):
        if not sender_email or not sender_password or not recipient_email:
            st.error("❌ 발신자/수신자 이메일과 비밀번호를 모두 입력해 주세요.")
        else:
            try:
                wb = Workbook()
                ws1 = wb.active
                ws1.title = "당회차_MB결과"
                ws1.append(["지표명", "수치", "단위", "평가"])
                ws1.append(["실험 회차", run_no, "Run", "정상"])
                ws1.append(["M/B 정합성", round(mass_closure, 2), "%", f"증발 {loss_mass:.1f}g"])
                ws1.append(["총 Li 회수율", round(li_recovery_pct, 2), "%", "1차여액+수세액"])
                ws1.append(["1차 여액 LiOH 농도", round(lioh_conc_g_l, 2), "g/L", f"Li: {primary_li:.1f} mg/L"])
                ws1.append(["소성 감율(LOI)", round(loi_pct, 2), "%", f"CaCO3 순도 ~{purity_caco3:.1f}%"])
                ws1.append(["차기 신품 CaO 보충량", round(fresh_makeup, 2), "g", f"재생분 {calcined_cao:.1f}g"])

                # 2번 시트에 시뮬레이션 데이터 추가
                ws2 = wb.create_sheet(title="트렌드시뮬레이션")
                ws2.append(list(df_simulation.columns))
                for _, r in df_simulation.iterrows():
                    ws2.append(list(r.values))

                excel_buffer = io.BytesIO()
                wb.save(excel_buffer)
                excel_buffer.seek(0)

                msg = MIMEMultipart()
                msg["From"] = sender_email
                msg["To"] = recipient_email
                msg["Subject"] = email_subject

                html_body = f"""
                <h3>🧪 습식반응 및 Ca-Loop Run {run_no} M/B & 트렌드 리포트</h3>
                <hr>
                <ul>
                    <li><b>총 Li 회수율:</b> {li_recovery_pct:.2f}%</li>
                    <li><b>1차 여액 LiOH 농도:</b> {lioh_conc_g_l:.2f} g/L (Li: {primary_li:,.1f} mg/L)</li>
                    <li><b>M/B 정합성(Closure):</b> {mass_closure:.2f}% (증발 손실: {loss_mass:.1f}g)</li>
                    <li><b>차기(Run {run_no+1}) 신품 CaO 보충량:</b> {fresh_makeup:.2f}g</li>
                </ul>
                <h4>🤖 AI 진단 코멘트</h4>
                <ul>
                    {''.join([f'<li>{d}</li>' for d in diagnostics])}
                </ul>
                <p>※ 회차별 트렌드 시뮬레이션 표가 포함된 엑셀 파일을 첨부하였습니다.</p>
                """
                msg.attach(MIMEText(html_body, "html", "utf-8"))

                part = MIMEApplication(excel_buffer.read(), Name=f"MB_Report_Run_{run_no:03d}.xlsx")
                part['Content-Disposition'] = f'attachment; filename="MB_Report_Run_{run_no:03d}.xlsx"'
                msg.attach(part)

                server = smtplib.SMTP(smtp_server, smtp_port)
                server.starttls()
                server.login(sender_email, sender_password)
                server.send_message(msg)
                server.quit()

                st.success(f"🎉 [{recipient_email}]로 엑셀 리포트 및 브리핑 메일이 성공적으로 발송되었습니다!")
            except Exception as e:
                st.error(f"❌ 메일 발송 실패: {e}")
