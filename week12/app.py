# 실행 방법 : 터미널에서 `streamlit run app.py` 를 실행하면 웹 브라우저가 열리면서 Streamlit 앱이 실행된다. (http://localhost:8501)
# 아래 단계별로 실행해보려면 단계별 코드를 주석 해제하고 app.py를 실행시킨 상태에서 브라우저를 새로고침 하면 된다. 

# 1단계 : st.write()
# import streamlit as st
# st.write("안녕하세요, Streamlit")

# 2단계 : st.title()
# import streamlit as st
# st.title("첫 Streamlit 화면")
# st.write("안녕하세요, Streamlit")

# 3단계 : st.text_input()
# import streamlit as st
# st.title("첫 Streamlit 화면")
# st.write("안녕하세요, Streamlit")
# name = st.text_input("이름을 입력하세요", value="홍길동")
# st.write(f"현재 입력값: {name}")

# 4단계 : st.button()
# import streamlit as st
# st.title("첫 Streamlit 화면")
# st.write("안녕하세요, Streamlit")
# name = st.text_input("이름을 입력하세요", value="홍길동")
# clicked = st.button("인사하기")
# if clicked:
#     st.write(f"{name}님, 반갑습니다.")
# else:
#     st.write("아직 버튼을 누르지 않았습니다.")

# 5단계 : st.sidebar
# import streamlit as st
# st.title("첫 Streamlit 화면")
# st.write("안녕하세요, Streamlit")
# name = st.text_input("이름을 입력하세요", value="홍길동")
# clicked = st.button("인사하기")
# with st.sidebar:
#     st.write("여기는 사이드바입니다.")
#     st.write("나중에는 종목코드나 주문 입력을 둘 수 있습니다.")
# if clicked:
#     st.write(f"{name}님, 반갑습니다.")
# else:
#     st.write("아직 버튼을 누르지 않았습니다.")

# 6단계 : st.metric()
# import streamlit as st
# st.title("첫 Streamlit 화면")
# st.write("안녕하세요, Streamlit")
# name = st.text_input("이름을 입력하세요", value="홍길동")
# clicked = st.button("인사하기")
# with st.sidebar:
#     st.write("여기는 사이드바입니다.")
#     st.write("나중에는 종목코드나 주문 입력을 둘 수 있습니다.")
# st.metric("현재 입력 글자 수", len(name), "입력 중")
# if clicked:
#     st.write(f"{name}님, 반갑습니다.")
# else:
#     st.write("아직 버튼을 누르지 않았습니다.")


# 최종 코드 통합
# import streamlit as st

# st.title("첫 Streamlit 화면")
# st.write("안녕하세요, Streamlit")

# name = st.text_input("이름을 입력하세요", value="홍길동")
# clicked = st.button("인사하기")

# with st.sidebar:
#     st.write("여기는 사이드바입니다.")
#     st.write("나중에는 종목코드나 주문 입력을 둘 수 있습니다.")

# st.metric("현재 입력 글자 수", len(name), "입력 중")

# if clicked:
#     st.write(f"{name}님, 반갑습니다.")
# else:
#     st.write("아직 버튼을 누르지 않았습니다.")


# 위젯과 이벤트 흐름 =====================================================
import streamlit as st # 여기는 계속 주석 해제해놓고 있기 

# 1 - st.button() - key 파라미터
# 버튼 위젯. 개개 버튼에는 고유한 key 값을 넣어 주어야 한다.
# x = st.button("클릭~", key="a")
# st.write(x)

# 2 - st.radio()
# 라디오 버튼 위젯.
# my_choice = st.radio("라디오 선택", ["부먹", "찍먹"])
# st.write(my_choice)

# 3 - st.checkbox()
# 체크박스 위젯.
# x = st.checkbox("위 내용에 동의합니다!")
# st.write(x)

# 4 - st.color_picker()
# 컬러 피커 위젯. 선택 결과는 16진수 컬러값.
# my_choice = st.color_picker("컬러 선택")
# st.write(my_choice)

# 5 - st.selectbox()
# 드롭다운 선택 위젯.
# my_choice = st.selectbox("다음 중 한가지 선택", ["부먹", "찍먹"])
# st.write(my_choice)

# 6 - st.multiselect()
# 멀티셀렉트 위젯.
# my_choice = st.multiselect("후식 선택", ["티라미수", "아이스크림", "과일샐러드", "수정과"])
# st.write(my_choice)

# 심화 위젯 -------------------------------------------------------

# 7 - st.text_input()
# text_input 위젯.
# my_name = st.text_input("당신의 이름은?")
# st.write(my_name)

# 8 - st.date_input()
# date_input 위젯.
# my_date = st.date_input("미팅 날짜는?")
# st.write(my_date)

# 9 - st.time_input()
# time_input 위젯.
# my_time = st.time_input("미팅 시각은?")
# st.write(my_time)

# 10 - st.slider()
# slider 위젯.
# x = st.slider("만족도 점수는?", min_value=0, max_value=10, value=5, step=1)
# st.write(x)

# 11 - st.select_slider()
# select_slider 위젯.
# x = st.select_slider("서비스 만족도는?", ["매우 불만족", "불만족", "보통", "만족", "매우 만족"])
# st.write(x)

# 12 - st.number_input()
# number_input 위젯.
# my_score = st.number_input("시험 점수는?", min_value=0, max_value=10, value=10, step=2)
# st.write(my_score)

# 위젯 값 다루기 - 반환값으로 받는다 =====================================
# 여러 위젯의 반환값을 변수로 받아 조합한다.
# 버튼이 True인 그 실행에서만 출력된다.

# name = st.text_input("이름")
# age  = st.slider("나이", 1, 100, 20)
# job  = st.selectbox("직업", ["학생", "개발자"])
# if st.button("제출"):
#     st.write(f"{name} {age} {job}")

# 문제점 : 재실행되면 변수가 초기화된다 =====================================
# count = 0          # 매 실행마다 0으로 초기화
# if st.button("증가"):
#     count += 1     # 0 + 1 = 1 (항상)
# st.write("현재 값:", count)

# Part4 : st.session_state — 재실행을 넘어 상태 지키기 ======================
# if "count" not in st.session_state:
#     st.session_state.count = 0
# if st.button("증가"):
#     st.session_state.count += 1
# st.write("현재 값:", st.session_state.count)

# 채팅 인터페이스 =====================================================
if "messages" not in st.session_state:
    st.session_state.messages = []
for m in st.session_state.messages:
    with st.chat_message(m["role"]):
        st.markdown(m["content"])
if prompt := st.chat_input("입력"):
    st.session_state.messages.append(
        {"role": "user", "content": prompt})
    with st.chat_message("user"):
        st.markdown(prompt)
    reply = f"echo: {prompt}"
    st.session_state.messages.append(
        {"role": "assistant", "content": reply})
    with st.chat_message("assistant"):
        st.markdown(reply)
