import os
import time
from dotenv import load_dotenv
from google import genai
from google.genai import types

load_dotenv()
client = genai.Client(api_key=os.environ.get('GOOGLE_API_KEY'))

MODEL = 'gemini-2.5-flash'

# 1. temperature 차이 비교
question = "Please recommend me a lunch menu for today"

for temp in [0.0, 0.5, 1.0]:
    response = client.models.generate_content(
        model=MODEL,
        contents=question,
        config=types.GenerateContentConfig(
            temperature=temp,         # Ollama의 temperature와 동일한 역할
            max_output_tokens=100     # Ollama의 num_predict와 동일한 역할
        )
    )
    print(f"[temperature {temp}]\n{response.text.strip()}\n")
    time.sleep(15)  # 할당량 초과 방지


# 2. system 파라미터 비교
question = 'Who are you?'

# system 없을 때
response = client.models.generate_content(
    model=MODEL,
    contents=question
)
print(f"[system 없음]\n{response.text.strip()}\n")

# system 있을 때
response = client.models.generate_content(
    model=MODEL,
    contents=question,
    config=types.GenerateContentConfig(
        system_instruction='You are a blunt pirate. Every answer begins with "Sailor!"',
        temperature=0.7,
        max_output_tokens=200
    )
)
print(f"[system 있음]\n{response.text.strip()}\n")


# 3. 단일 메시지 vs 멀티턴 비교

# 단일 메시지 (기억 못함)
for msg in ['My name is Wonbin. Please remember.', 'What''s my name?']:
    response = client.models.generate_content(model=MODEL, contents=msg)
    print(f"[단일] {msg}\n→ {response.text.strip()}\n")

# 멀티턴 (chats 사용 - 히스토리 자동 관리)
chat = client.chats.create(model=MODEL)

for msg in ['My name is Wonbin. Please remember.', 'What''s my name?']:
    response = chat.send_message(msg)
    print(f"[멀티턴] {msg}\n→ {response.text.strip()}\n")