import os
from dotenv import load_dotenv
from google import genai

load_dotenv()

api_key = os.environ.get('GOOGLE_API_KEY')
client = genai.Client(api_key=api_key)

try:
    response = client.models.generate_content(
        model='gemini-2.0-flash',  
        contents='한국의 수도는 어디인가요?'
    )

    print(f"답변: {response.text}")
    print(f"입력 토큰: {response.usage_metadata.prompt_token_count}")
    print(f"출력 토큰: {response.usage_metadata.candidates_token_count}")

except Exception as e:
    print(f"에러 발생 상세: {e}")

