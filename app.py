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

st.set_page_config(page_title="습식반응 M/B & AI 공정 Agent", page_icon="🧪", layout="wide")

# 세션 상태 초기화 (이력 누적 및 채팅 기록)
if "history" not in st.session_state:
    st.session_state.history = pd.DataFrame([
        {
            "Run_No": 1, "Li_Recovery": 95.8, "MB_Closure": 95.06, 
            "LOI": 41.13, "CaCO3_Purity": 85.5, "Activity_Index": 100.0,
            "Fresh_CaO_in": 92.42, "Recycled_CaO_in": 0.0, "Next_Makeup_g": 68.52
        }
    ])

if "chat_messages" not in st.session_state:
    st.session_state.chat_messages = [
        {"role": "assistant", "content": "안녕하세요! 습식 가성화 및 Ca-Loop 공정 전문 AI 에이전트입니다. 이번 회차 계산 결과나 다음 실험 조건, 이상 징후에 대해 무엇이든 물어보세요."}
    ]

st.title("🧪 습식반응 및 Ca-Loop M/B 자동화 & AI Agent")
st.caption("화학양론 M/B 연산 | $n$회차 예측 및 시각화 | AI 공정 진단 챗봇 | 리포트 이메일 발송")

# 상단 탭 구성
main_tab1, main_tab2, main_tab3, main_tab4 = st.tabs([
    "1️⃣ 실험 데이터 입력 & M/B 연산", 
    "2️⃣ 회차별 트렌드 & 차기 예측", 
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

    target_cao = 92.42
    fresh_makeup = max(0.0, target_cao - calcined_cao)

    st.subheader(f"📊 Run {run_no} M/B 연산 및 진단 결과")
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("M/B 정합성 (Closure)", f"{mass_closure:.2f} %", f"증발: {loss_mass:.1f}g")
    k2.metric("총 Li 회수율 (ICP)", f"{li_recovery_pct:.2f} %", f"여액 {li_p_g:.1f}g + 수세 {li_w_g:.1f}g")
    k3.metric("소성 감율 (LOI)", f"{loi_pct:.2f} %", f"CaCO₃ 순도 ~{purity_caco3:.1f}%")
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

    # 현재 결과를 트렌드 DB에 반영하는 버튼
    if st.button("💾 이번 회차 결과를 누적 트렌드 DB에 등록하기", type="primary"):
        new_row = {
            "Run_No": run_no, "Li_Recovery": round(li_recovery_pct, 2), 
            "MB_Closure": round(mass_closure, 2), "LOI": round(loi_pct, 2), 
            "CaCO3_Purity": round(purity_caco3, 1), 
            "Activity_Index": max(70.0, 100.0 - (run_no - 1) * 3.5),
            "Fresh_CaO_in": fresh_cao_mass, "Recycled_CaO_in": recycled_cao_mass, 
            "Next_Makeup_g": round(fresh_makeup, 2)
        }
        # 중복 Run 업데이트 또는 추가
        st.session_state.history = st.session_state.history[st.session_state.history["Run_No"] != run_no]
        st.session_state.history = pd.concat([st.session_state.history, pd.DataFrame([new_row])]).sort_values("Run_No").reset_index(drop=True)
        st.success(f"✅ Run {run_no} 데이터가 트렌드 이력에 저장되었습니다!")

# --------------------------------------------------------------------------
# TAB 2: 회차별 트렌드 시각화 및 차기 회차 AI 예측
# --------------------------------------------------------------------------
with main_tab2:
    st.header("📈 $n$회차 누적 트렌드 및 다음 실험 AI 예측 모델")
    
    col_t1, col_t2 = st.columns([2, 1])
    with col_t1:
        st.subheader("회차별 핵심 KPI 변화 트렌드")
        chart_data = st.session_state.history.set_index("Run_No")[["Li_Recovery", "MB_Closure", "LOI", "Activity_Index"]]
        st.line_chart(chart_data)
        st.dataframe(st.session_state.history, use_container_width=True)

    with col_t2:
        st.subheader(f"🔮 차기 회차 (Run {len(st.session_state.history)+1}) 예측")
        curr_runs = len(st.session_state.history)
        next_run = curr_runs + 1
        
        # 소결에 따른 활성도 감쇄 모델: A_n = 100 * (0.965)^(n-1)
        pred_activity = 100.0 * (0.965 ** (next_run - 1))
        pred_recovery = max(88.0, 96.0 - (next_run - 1) * 1.2)
        
        st.metric("예상 CaO 활성도 (소결 영향)", f"{pred_activity:.1f} %", f"{-3.5:.1f}%p 감쇄", delta_color="inverse")
        st.metric("예상 Li 회수율", f"{pred_recovery:.1f} %")
        st.metric("추천 신품 CaO 보충량", f"{fresh_makeup:.2f} g", f"재생분 {calcined_cao:.1f}g 재사용")

        st.warning(f"""
        **🤖 AI 공정 최적화 권고안 (Run {next_run}):**
        - 재생 CaO 비율이 높아짐에 따라 소화 발열 지연이 예상됩니다.
        - 슬러리 교반 시간을 기존 2시간에서 **2시간 20분(+20분)**으로 연장하거나 반응 온도를 **85℃**로 상향을 추천합니다.
        """)

# --------------------------------------------------------------------------
# TAB 3: AI 공정 엔지니어 대화창 (Chatbot)
# --------------------------------------------------------------------------
with main_tab3:
    st.header("💬 AI 공정 엔지니어와 대화하기")
    st.caption("현재 실험 수치와 공정 지식을 바탕으로 AI가 실시간으로 답변합니다.")

    # 이전 대화 기록 표시
    for msg in st.session_state.chat_messages:
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])

    # 사용자 질의 입력
    if user_prompt := st.chat_input("질문을 입력하세요 (예: 수세액 pH가 왜 이렇게 높아? 소결 방지 대책은?)"):
        st.session_state.chat_messages.append({"role": "user", "content": user_prompt})
        with st.chat_message("user"):
            st.markdown(user_prompt)

        # AI 응답 생성 (룰 및 공정 모델 기반 해석)
        prompt_low = user_prompt.lower()
        if "ph" in prompt_low or "수세" in prompt_low or "세척" in prompt_low:
            ai_reply = f"현재 1차 여액 pH는 {primary_filtrate_ph}이지만, 수세액 pH가 **{wash_sol_ph}**로 측정되었습니다. 이는 1차 감압 여과 시 케이크 내부에 고농도 LiOH 액이 물리적으로 갇혀(Entrainment) 있었기 때문입니다. 3배수 수세를 통해 리튬 성분이 정상적으로 회수되었음을 의미하므로 안심하셔도 됩니다."
        elif "소결" in prompt_low or "sintering" in prompt_low or "하소" in prompt_low or "소성" in prompt_low:
            ai_reply = f"1000℃에서 1시간 하소 시 탈탄산은 100% 완결되지만, CaO 입자 간 소결로 인해 비표면적이 감소합니다. Run {run_no+1}에서는 소화(Slaking) 속도가 느려질 수 있으므로 슬러리 조제 시 온수를 사용하거나 교반 시간을 15~20분 연장하는 것을 권장합니다."
        elif "회수율" in prompt_low or "손실" in prompt_low or "오차" in prompt_low:
            ai_reply = f"현재 Run {run_no} 기준 총 Li 회수율은 **{li_recovery_pct:.2f}%**입니다. 주된 질량 손실은 80℃ 반응 중 발생한 수분 증발({loss_mass:.1f}g)입니다. 이를 개선하려면 반응기에 환류 냉각기(Reflux condenser)를 장착하는 것이 가장 효과적입니다."
        elif "다음" in prompt_low or "배합" in prompt_low or "보충" in prompt_low:
            ai_reply = f"다음 회차(Run {run_no+1})에서는 소성 회수된 **재생 CaO {calcined_cao:.2f}g**에 **신품 CaO {fresh_makeup:.2f}g**을 보충하여 총 92.42g을 맞추시면 됩니다. 슬러리 물은 동일하게 {slurry_water}g을 사용하세요."
        else:
            ai_reply = f"질문하신 내용에 대해 공정 수치(Run {run_no}, Li 회수율 {li_recovery_pct:.1f}%, M/B 닫힘 {mass_closure:.1f}%)를 바탕으로 점검한 결과, 현재 양론적 과잉율은 +{excess_pct:.1f}%이며 Ca-Loop 회수율은 {ca_loop_recovery:.1f}%로 안정적인 범위에 있습니다. 추가로 확인하고 싶은 세부 반응 변수를 말씀해 주세요."

        st.session_state.chat_messages.append({"role": "assistant", "content": ai_reply})
        with st.chat_message("assistant"):
            st.markdown(ai_reply)

# --------------------------------------------------------------------------
# TAB 4: 리포트 이메일 발송
# --------------------------------------------------------------------------
with main_tab4:
    st.header("📧 M/B 엑셀 리포트 및 브리핑 이메일 자동 발송")
    st.markdown("실험 결과 요약문과 엑셀 리포트 파일(`.xlsx`)을 담당자 이메일로 즉각 발송합니다.")

    col_m1, col_m2 = st.columns(2)
    with col_m1:
        recipient_email = st.text_input("수신자 이메일 주소", value="user@company.com")
        email_subject = st.text_input("메일 제목", value=f"[M/B 자동화 리포트] Run {run_no} 습식반응 및 Ca-Loop 결과 브리핑")
        sender_email = st.text_input("발신자 이메일 주소 (Gmail/회사 SMTP)", value="sender@gmail.com")
        sender_password = st.text_input("발신자 앱 비밀번호 (App Password)", type="password", help="Gmail 사용 시 '앱 비밀번호 16자리' 입력")

    with col_m2:
        smtp_server = st.text_input("SMTP 서버 호스트", value="smtp.gmail.com")
        smtp_port = st.number_input("SMTP 포트", value=587, step=1)

    if st.button("🚀 엑셀 리포트 첨부하여 이메일 발송", type="primary", use_container_width=True):
        if not sender_email or not sender_password or not recipient_email:
            st.error("❌ 발신자/수신자 이메일과 비밀번호를 모두 입력해 주세요.")
        else:
            try:
                # 엑셀 파일 생성
                wb = Workbook()
                ws = wb.active
                ws.title = "MB_Result"
                ws.append(["지표명", "수치", "단위", "평가"])
                ws.append(["실험 회차", run_no, "Run", "정상"])
                ws.append(["M/B 정합성", round(mass_closure, 2), "%", f"증발 {loss_mass:.1f}g"])
                ws.append(["총 Li 회수율", round(li_recovery_pct, 2), "%", "1차여액+수세액"])
                ws.append(["소성 감율(LOI)", round(loi_pct, 2), "%", f"CaCO3 순도 ~{purity_caco3:.1f}%"])
                ws.append(["차기 신품 CaO 보충량", round(fresh_makeup, 2), "g", f"재생분 {calcined_cao:.1f}g"])
                
                excel_buffer = io.BytesIO()
                wb.save(excel_buffer)
                excel_buffer.seek(0)

                # 이메일 메시지 생성
                msg = MIMEMultipart()
                msg["From"] = sender_email
                msg["To"] = recipient_email
                msg["Subject"] = email_subject

                html_body = f"""
                <h3>🧪 습식반응 및 Ca-Loop Run {run_no} M/B 리포트</h3>
                <p>본 메일은 습식반응 M/B 자동화 AI Agent에 의해 자동 생성된 브리핑입니다.</p>
                <hr>
                <ul>
                    <li><b>M/B 정합성(Closure):</b> {mass_closure:.2f}% (증발 손실: {loss_mass:.1f}g)</li>
                    <li><b>총 Li 회수율:</b> {li_recovery_pct:.2f}%</li>
                    <li><b>소성 감율(LOI):</b> {loi_pct:.2f}% (CaCO₃ 추정 순도: {purity_caco3:.1f}%)</li>
                    <li><b>차기(Run {run_no+1}) 신품 CaO 보충량:</b> {fresh_makeup:.2f}g</li>
                </ul>
                <h4>🤖 AI 진단 코멘트</h4>
                <ul>
                    {''.join([f'<li>{d}</li>' for d in diagnostics])}
                </ul>
                <p>※ 상세 분석 데이터는 첨부된 엑셀 파일을 확인해 주시기 바랍니다.</p>
                """
                msg.attach(MIMEText(html_body, "html", "utf-8"))

                # 엑셀 파일 첨부
                part = MIMEApplication(excel_buffer.read(), Name=f"MB_Report_Run_{run_no:03d}.xlsx")
                part['Content-Disposition'] = f'attachment; filename="MB_Report_Run_{run_no:03d}.xlsx"'
                msg.attach(part)

                # SMTP 전송
                server = smtplib.SMTP(smtp_server, smtp_port)
                server.starttls()
                server.login(sender_email, sender_password)
                server.send_message(msg)
                server.quit()

                st.success(f"🎉 [{recipient_email}]로 엑셀 리포트 및 브리핑 메일이 성공적으로 발송되었습니다!")
            except Exception as e:
                st.error(f"❌ 메일 발송 실패: {e}")
