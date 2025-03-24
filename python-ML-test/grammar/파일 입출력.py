# 파일 쓰기
with open("sample.txt", "w", encoding="utf-8") as f:
    f.write("Hello, Python!\n")

# 파일 읽기
with open("sample.txt", "r", encoding="utf-8") as f:
    print(f.read())
