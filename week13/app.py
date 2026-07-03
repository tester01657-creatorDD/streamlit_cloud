import re
import io
import concurrent.futures
import requests

import streamlit as st
import google.generativeai as genai
from dotenv import load_dotenv
import os

from langchain_core.messages import AIMessage, HumanMessage
from langchain_core.tools import tool
from langgraph.graph import StateGraph, START, END, MessagesState
from langgraph.errors import GraphRecursionError

# ── 환경 설정 ──────────────────────────────────────────────────────────────
load_dotenv(dotenv_path=os.path.join(os.path.dirname(__file__), "..", ".env"))
genai.configure(api_key=os.getenv("GOOGLE_API_KEY"))

# ── 가드레일 패턴 (week11 Part 8) ──────────────────────────────────────────
BLOCK = re.compile(r"(ignore|무시).*(instruction|지시)", re.I)
WEATHER_KWS = ["날씨", "weather", "기온", "비", "맑"]

# ── 날씨 도구 (week11 Part 3) ───────────────────────────────────────────────
@tool
def get_weather(city: str) -> str:
    """도시의 현재 날씨를 조회한다."""
    try:
        r = requests.get(f"https://wttr.in/{city}?format=3", timeout=5)
        r.raise_for_status()
        return r.text
    except requests.Timeout:
        return "오류: 날씨 서버 응답 시간 초과"
    except Exception as e:
        return f"오류: 날씨 조회 실패 ({e})"

# ── 날씨 전용 LangGraph 에이전트 (week11 Part 6) ───────────────────────────
class AgentState(MessagesState):
    error: str

def weather_node(state: AgentState):
    text = state["messages"][-1].content
    for kw in WEATHER_KWS + ["알려줘", "?"]:
        text = text.replace(kw, "")
    city = text.strip()
    if not city:
        return {"error": "도시명 없음",
                "messages": [AIMessage(content="도시명을 찾을 수 없습니다. 예: 'Seoul 날씨'")]}
    result = get_weather.invoke({"city": city})
    if result.startswith("오류"):
        return {"error": "조회 실패", "messages": [AIMessage(content=result)]}
    return {"error": "", "messages": [AIMessage(content=result)]}

def fallback_node(state: AgentState):
    return {"messages": [AIMessage(content="지금은 날씨를 조회할 수 없습니다.")]}

def check_error(state: AgentState) -> str:
    return "fallback" if state["error"] else END

builder = StateGraph(AgentState)
builder.add_node("weather",  weather_node)
builder.add_node("fallback", fallback_node)
builder.add_edge(START, "weather")
builder.add_conditional_edges("weather", check_error, {"fallback": "fallback", END: END})
builder.add_edge("fallback", END)
weather_agent = builder.compile()

def run_weather_agent(question: str, timeout: int = 10) -> str:
    inputs = {"messages": [HumanMessage(content=question)], "error": ""}
    with concurrent.futures.ThreadPoolExecutor(max_workers=1) as executor:
        future = executor.submit(weather_agent.invoke, inputs, {"recursion_limit": 10})
        try:
            result = future.result(timeout=timeout)
            return result["messages"][-1].content
        except GraphRecursionError:
            return "처리 단계가 너무 길어 중단했습니다."
        except concurrent.futures.TimeoutError:
            return "응답 시간이 초과되었습니다."

# ── 스트리밍 제너레이터 (토큰 흘려보내기) ────────────────────────────────────
def stream_gemini(prompt: str, cfg: dict):
    model = genai.GenerativeModel(
        "gemini-2.5-flash",
        generation_config=genai.GenerationConfig(
            temperature=cfg["temp"],
            max_output_tokens=int(cfg["max_tokens"]),
        ),
    )
    for chunk in model.generate_content(prompt, stream=True):
        if chunk.text:
            yield chunk.text  # 토큰 하나씩 yield

def stream_ollama(prompt: str, cfg: dict):
    import ollama
    for chunk in ollama.chat(
        model="llama3.2:3b",
        messages=[{"role": "user", "content": prompt}],
        stream=True,
    ):
        yield chunk["message"]["content"]

# ── PDF 텍스트 추출 ────────────────────────────────────────────────────────
def extract_pdf_text(uploaded_file) -> str:
    try:
        import pypdf
        reader = pypdf.PdfReader(io.BytesIO(uploaded_file.getvalue()))
        return "\n".join(page.extract_text() or "" for page in reader.pages)
    except Exception:
        return ""

# ── 사이드바 설정 패널 (Part 4) ────────────────────────────────────────────
with st.sidebar:
    st.header("설정")
    model_choice = st.selectbox("모델", ["Gemini", "Ollama"])
    temp      = st.slider("창의성", 0.0, 1.0, 0.7)
    max_tok   = st.number_input("최대 토큰", 256, 4096, 1024)
    pdf       = st.file_uploader("문서", type=["pdf"])
    if st.button("대화 초기화"):
        st.session_state.messages = []
        st.rerun()

# 설정값을 session_state에 저장 → 에이전트에 전달
st.session_state.cfg = dict(model=model_choice, temp=temp, max_tokens=max_tok, pdf=pdf)

# ── 메인 UI ────────────────────────────────────────────────────────────────
st.title("LangGraph 에이전트 챗봇")
st.caption("날씨 질문 → LangGraph + wttr.in  |  일반 질문 → 선택 모델 스트리밍")

# 이력 초기화
if "messages" not in st.session_state:
    st.session_state.messages = []

# 기존 이력 전체 렌더링
for m in st.session_state.messages:
    with st.chat_message(m["role"]):
        st.markdown(m["content"])

# ── 입력 → 에이전트 → 응답 ─────────────────────────────────────────────────
if prompt := st.chat_input("메시지를 입력하세요 (예: Seoul 날씨, 파이썬이란?)"):
    cfg = st.session_state.cfg

    # 1. 사용자 입력을 먼저 이력에 추가
    st.session_state.messages.append({"role": "user", "content": prompt})
    with st.chat_message("user"):
        st.markdown(prompt)

    # ── 가드레일: 프롬프트 인젝션 차단 ──────────────────────────────────
    if BLOCK.search(prompt):
        reply = "허용되지 않은 요청입니다."
        st.session_state.messages.append({"role": "assistant", "content": reply})
        with st.chat_message("assistant"):
            st.markdown(reply)

    # ── 날씨 질문 → LangGraph 에이전트 ──────────────────────────────────
    elif any(kw in prompt.lower() for kw in WEATHER_KWS):
        with st.spinner("날씨 조회 중..."):
            reply = run_weather_agent(prompt)
        # 4. 최종 응답을 이력에 append
        st.session_state.messages.append({"role": "assistant", "content": reply})
        with st.chat_message("assistant"):
            st.markdown(reply)

    # ── 일반 질문 → 스트리밍 ─────────────────────────────────────────────
    else:
        # PDF가 업로드된 경우 문서 내용을 프롬프트 앞에 추가
        full_prompt = prompt
        if cfg["pdf"]:
            doc_text = extract_pdf_text(cfg["pdf"])
            if doc_text.strip():
                full_prompt = f"[문서 내용]\n{doc_text[:3000]}\n\n[질문]\n{prompt}"

        with st.chat_message("assistant"):
            # 2. 제너레이터가 토큰을 하나씩 yield → 화면에 실시간 출력
            # 3. 스트리밍이 끝나면 st.write_stream이 완성된 텍스트를 반환
            if cfg["model"] == "Gemini":
                reply = st.write_stream(stream_gemini(full_prompt, cfg))
            else:
                reply = st.write_stream(stream_ollama(full_prompt, cfg))

        # 4. 최종 응답을 이력에 append → 다음 턴의 문맥으로 재사용
        st.session_state.messages.append({"role": "assistant", "content": reply})
