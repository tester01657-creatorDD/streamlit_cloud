import ollama

# 1. temperature 차이 비교
question = "한국은 어떤가요?"

for temp in [0.0, 0.5, 1.0]:
    response = ollama.chat(
        model='llama3.2:3b',
        messages=[{'role': 'user', 'content': question}],
        options={
            'temperature': temp,  # Gemini의 temperature와 동일한 역할
            'num_predict': 100    # Gemini의 max_output_tokens와 동일한 역할
        }
    )
    print(f"[temperature {temp}]\n{response.message.content.strip()}\n")
