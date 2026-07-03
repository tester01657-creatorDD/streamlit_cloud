# 실행: streamlit run week13/rag_agent.py

import os
import tempfile
import requests
import numpy as np
import streamlit as st
import fitz                                      # PyMuPDF
import easyocr
import chromadb
from dotenv import load_dotenv
from langchain_core.tools import tool
from langchain_core.messages import HumanMessage
from langchain_google_genai import ChatGoogleGenerativeAI
from langchain.agents import create_agent
from langchain_ollama import ChatOllama          # Part 6: Ollama 로컬 백엔드

load_dotenv(dotenv_path=os.path.join(os.path.dirname(__file__), "..", ".env"))

DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")


# 5-1. parse_pdf  :  PDF 경로 → 텍스트 청크 리스트
@st.cache_resource
def get_ocr_reader():
    """EasyOCR 리더를 한 번만 초기화해 재사용한다 (모델 로딩 수 초 소요)."""
    return easyocr.Reader(['ko', 'en'], gpu=False, verbose=False)

def parse_pdf(pdf_path: str, chunk_size: int = 500) -> list[str]:
    """각 페이지에서 텍스트를 추출한다. 텍스트가 없으면 EasyOCR로 대체한다."""
    doc = fitz.open(pdf_path)
    pages_text = []
    for page in doc:
        text = page.get_text().strip()
        if not text:                             # 이미지 기반 PDF → OCR
            pix = page.get_pixmap(dpi=150)
            img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.h, pix.w, pix.n)
            results = get_ocr_reader().readtext(img, detail=0)
            text = "\n".join(results)
        pages_text.append(text)
    doc.close()
    full_text = "\n".join(pages_text)
    chunks = [full_text[i : i + chunk_size] for i in range(0, len(full_text), chunk_size)]
    return [c for c in chunks if c.strip()]


# 5-1. build_db  :  청크 리스트 → ChromaDB 인메모리 컬렉션
def build_db(chunks: list[str]) -> chromadb.Collection:
    """청크를 ChromaDB에 임베딩(벡터화)해 저장하고 컬렉션을 반환한다."""
    client = chromadb.Client()                   # 인메모리 (세션 안에서만 유지)
    col = client.get_or_create_collection("pdf_docs")
    col.upsert(
        documents=chunks,
        ids=[f"chunk_{i}" for i in range(len(chunks))],
    )
    return col


# 5-2. make_tools  :  도구 2개 — RAG 검색 + 날씨 API
def make_tools(collection: chromadb.Collection) -> list:
    """벡터 DB를 검색하는 도구와 날씨 API 도구를 생성해 리스트로 반환한다."""

    @tool
    def search_documents(query: str) -> str:
        """업로드된 PDF 문서에서 질문과 관련된 내용을 검색한다."""
        res = collection.query(query_texts=[query], n_results=3)
        docs = res["documents"][0]
        return "\n\n".join(docs) if docs else "관련 내용을 찾을 수 없습니다."

    @tool
    def get_weather(city: str) -> str:
        """도시 이름을 받아 현재 날씨를 조회한다. (예: Seoul, Tokyo, 부산)"""
        try:
            r = requests.get(f"https://wttr.in/{city}?format=3", timeout=5)
            r.raise_for_status()
            return r.text
        except requests.Timeout:
            return "오류: 날씨 서버 응답 시간 초과"
        except Exception as e:
            return f"오류: 날씨 조회 실패 ({e})"

    return [search_documents, get_weather]


# 5-3. make_agent  :  LLM + 도구 → ReAct 에이전트
def make_agent(tools: list, cfg: dict):
    """Gemini LLM에 도구를 결합해 ReAct 에이전트를 반환한다."""
    llm = ChatGoogleGenerativeAI(
        model="gemini-2.5-flash",
        temperature=cfg["temp"],
        max_output_tokens=int(cfg["max_tokens"]),
    )
    return create_agent(llm, tools)


# 5-4. run_chat  :  에이전트 실행 → 토큰 스트리밍
def run_chat(question: str, agent):
    """에이전트로 질문을 처리하고 최종 답변을 토큰 단위로 yield한다."""
    inputs = {"messages": [HumanMessage(content=question)]}
    for msg, metadata in agent.stream(inputs, stream_mode="messages"):
        if (
            metadata.get("langgraph_node") == "model"
            and hasattr(msg, "content")
            and isinstance(msg.content, str)   # 도구 호출 청크(list)는 제외
            and msg.content
        ):
            yield msg.content


# Part 6. ask_agent — 교체 지점: LLM 한 줄만 바꿔 백엔드를 전환
# make_agent + run_chat를 하나로 묶은 스트리밍 제너레이터
def ask_agent(question: str, tools: list, backend: str, cfg: dict):
    """backend 값 하나로 Gemini ↔ Ollama를 전환한다."""
    if backend == "Gemini":
        llm = ChatGoogleGenerativeAI(
            model="gemini-2.5-flash",
            temperature=cfg["temp"],
            max_output_tokens=int(cfg["max_tokens"]),
        )
    else:  # Ollama — 로컬 실행, 비용 0, 외부 전송 없음
        llm = ChatOllama(
            model="qwen2.5",
            base_url="http://localhost:11434",
            temperature=cfg["temp"],
            num_predict=int(cfg["max_tokens"]),
        )
    agent = create_agent(llm, tools)
    for msg, metadata in agent.stream(
        {"messages": [HumanMessage(content=question)]},
        stream_mode="messages",
    ):
        if (
            metadata.get("langgraph_node") == "model"
            and hasattr(msg, "content")
            and isinstance(msg.content, str)
            and msg.content
        ):
            yield msg.content


# 사이드바 설정 패널
with st.sidebar:
    st.header("설정")

    # data/ 폴더에서 PDF 목록 자동 탐지
    pdf_files = (
        sorted(f for f in os.listdir(DATA_DIR) if f.lower().endswith(".pdf"))
        if os.path.isdir(DATA_DIR) else []
    )
    selected = st.selectbox("PDF 선택 (data/ 폴더)", pdf_files) if pdf_files else None
    uploaded = st.file_uploader("또는 직접 업로드", type=["pdf"])

    temp    = st.slider("창의성", 0.0, 1.0, 0.7)
    max_tok = st.number_input("최대 토큰", 256, 4096, 1024)

    if st.button("대화 초기화"):
        st.session_state.messages = []
        st.rerun()

cfg = dict(temp=temp, max_tokens=max_tok)

# Part 6: 백엔드 선택 — 이 한 줄만 바꾸면 Gemini ↔ Ollama 전환
with st.sidebar:
    st.divider()
    backend = st.selectbox("백엔드 (Part 6)", ["Gemini", "Ollama"])

# PDF가 바뀔 때만 파싱 → 벡터화 재실행
pdf_key = uploaded.name if uploaded else selected

if pdf_key and st.session_state.get("pdf_key") != pdf_key:
    with st.spinner("PDF 파싱 및 벡터화 중..."):
        if uploaded:
            # 업로드 파일 → 임시 파일로 저장 후 파싱
            with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
                tmp.write(uploaded.getvalue())
                tmp_path = tmp.name
            chunks = parse_pdf(tmp_path)
            os.unlink(tmp_path)
        else:
            chunks = parse_pdf(os.path.join(DATA_DIR, selected))

        st.session_state.collection = build_db(chunks)
        st.session_state.pdf_key    = pdf_key
        st.session_state.messages   = []

    st.sidebar.success(f"{len(chunks)}개 청크 저장 완료")

# 메인 UI
st.title("PDF RAG + 외부 API 에이전트")
st.caption("PDF 질문 → ChromaDB RAG 검색  |  날씨 질문 → wttr.in API")

if "messages" not in st.session_state:
    st.session_state.messages = []

# 기존 이력 렌더링
for m in st.session_state.messages:
    with st.chat_message(m["role"]):
        st.markdown(m["content"])

# PDF 미선택 안내
if "collection" not in st.session_state:
    st.info("사이드바에서 PDF를 선택하거나 업로드하면 대화를 시작할 수 있습니다.")

elif prompt := st.chat_input("질문을 입력하세요 (예: 이 문서에서 안전 수칙은?, Seoul 날씨)"):
    # 1. 사용자 입력을 먼저 이력에 추가
    st.session_state.messages.append({"role": "user", "content": prompt})
    with st.chat_message("user"):
        st.markdown(prompt)

    # 에이전트는 cfg가 바뀔 수 있으므로 매번 생성
    tools = make_tools(st.session_state.collection)
    agent = make_agent(tools, cfg)

    # 2. 토큰 스트리밍으로 실시간 출력
    # 3. st.write_stream이 완성 텍스트를 반환
    # 4. 이력에 append → 다음 턴에 재사용
    with st.chat_message("assistant"):
        reply = st.write_stream(run_chat(prompt, agent))
    st.session_state.messages.append({"role": "assistant", "content": reply})
