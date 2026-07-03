import google.generativeai as genai

genai.configure(api_key="GOOGLE_API_KEY")

sentences = [
    "오늘 서울은 맑고 따뜻한 봄 날씨입니다",
    "내일도 화창하고 기온이 올라갈 예정입니다",
    "파이썬에서 리스트를 정렬하려면 sort 메서드를 사용합니다",
]

vectors = []
for s in sentences:
    result = genai.embed_content(
        model="models/text-embedding-004",
        content=s
    )
    vectors.append(result['embedding'])
    print(f"'{s}' → 앞 5개: {result['embedding'][:5]}")

print(f"벡터 길이: {len(vectors[0])}")       # 768
print(f"앞 5개만: {vectors[0][:5]}")
