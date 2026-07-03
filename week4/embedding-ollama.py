import ollama

# API 키 불필요 — 로컬 서버가 대신함
# 사전 준비: 터미널에서 ollama pull nomic-embed-text

sentences = [
    "오늘 날씨가 좋다",
    "날씨가 화창하네",
    "주식이 올랐다"
]

vectors = []
for s in sentences:
    result = ollama.embed(
        model="nomic-embed-text",
        input=s
    )
    vectors.append(result['embeddings'][0])
    print(f"'{s}' → 앞 5개: {result['embeddings'][0][:5]}")
    