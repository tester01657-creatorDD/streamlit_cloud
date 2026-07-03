import ollama

response = ollama.chat(
    model='llama3.2:3b',                    # ollama pull로 받은 모델명
    messages=[
        {'role': 'user', 'content': '한국의 수도는 어디인가요?'}
    ]
)

print(response.message.content)
print('모델:', response.model)
print('입력 토큰:', response.prompt_eval_count)
print('출력 토큰:', response.eval_count)

