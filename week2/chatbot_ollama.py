import ollama

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
    history.append({'role': 'user', 'content': user_input})

    # API 호출
    response = ollama.chat(
        model='llama3.2:3b',
        # 변경: system을 messages 첫 번째에 추가
        messages=[
            {'role': 'system', 'content': SYSTEM_PROMPT},
            *history   # 히스토리 펼치기
        ],
        options={'temperature': 0.7, 'num_predict': 500}
    )

    # AI 응답도 히스토리에 추가
    history.append({'role': 'assistant', 'content': response.message.content})

    # 응답 출력
    print(f"Ollama: {response.message.content}")
