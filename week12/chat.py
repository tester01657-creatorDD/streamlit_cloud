import streamlit as st
import google.generativeai as genai
from dotenv import load_dotenv
import os

# Gemini 백엔드
load_dotenv(dotenv_path=os.path.join(os.path.dirname(__file__), "..", ".env"))
genai.configure(api_key=os.getenv("GOOGLE_API_KEY"))
model = genai.GenerativeModel("gemini-2.5-flash")

# save(role, content) : 메시지 한 줄을 messages 리스트에 더하는 헬퍼
def save(role, content):
    st.session_state.messages.append({"role": role, "content": content})
    with st.chat_message(role):
        st.markdown(content)

# ask_llm(prompt) : LLM을 호출해 응답을 반환하는 함수

# --- Gemini 버전 ---
def ask_llm(prompt):
    return model.generate_content(prompt).text

# # --- Ollama 버전 (로컬) ---
# import ollama
# def ask_llm(prompt):
#     res = ollama.chat(
#         model="llama3.2:3b",
#         messages=[{
#             "role": "user",
#             "content": prompt}])
#     return res["message"]["content"]

# 1) 이력 초기화 (없으면 빈 리스트)
if "messages" not in st.session_state:
    st.session_state.messages = []

# 2) 기존 이력 전체를 다시 그린다
for m in st.session_state.messages:
    with st.chat_message(m["role"]):
        st.markdown(m["content"])

# 3~6) 입력 -> LLM 호출 -> 응답 저장
if prompt := st.chat_input("메시지"):
    save("user", prompt)
    reply = ask_llm(prompt)
    save("assistant", reply)
