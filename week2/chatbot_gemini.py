import os
from dotenv import load_dotenv
from google import genai
from google.genai import types  # ← 추가

load_dotenv()
client = genai.Client(api_key=os.environ.get('GOOGLE_API_KEY'))

# 페르소나 정의 (자유롭게 바꿔보기)
SYSTEM_PROMPT = 'You are a friendly Python tutor. Always explain in simple terms for beginners.'

history = []  # 대화 히스토리 리스트

print("챗봇을 시작합니다. 종료하려면 'quit'을 입력하세요.")

while True:
    # 사용자 입력
    user_input = input("You: ").strip()

    # 빈 입력 건너뜀
    if not user_input:
        continue

    # 종료
    if user_input == 'quit':
        print("대화를 종료합니다.")
        break

    # reset 커맨드
    if user_input == 'reset':
        history.clear()
        print("대화 히스토리를 초기화했습니다.")
        continue

    # user 메시지를 히스토리에 추가
    history.append({'role': 'user', 'parts': [user_input]})
    
    # API 호출
    response = client.models.generate_content(
        model='gemini-2.5-flash',
        contents=history,     # contents에 단순 문자열 대신 history 전달
        config=types.GenerateContentConfig(        # system_instruction으로 페르소나 적용
            system_instruction=SYSTEM_PROMPT,
            temperature=0.7,
            max_output_tokens=500
        )
    )

    # AI 응답도 히스토리에 추가
    history.append({'role': 'model', 'parts': [response.text]})

    # 응답 출력
    print(f"Gemini: {response.text}")
    
