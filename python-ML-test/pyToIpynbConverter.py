import os
import subprocess
from pathlib import Path
import shutil

# 원본 폴더와 대상 폴더 경로 설정
src_root = Path(r"C:\Users\user\Desktop\python-ML-test")
dst_root = Path(r"C:\Users\user\Desktop\python-ML-test-converted")

# 기존 폴더 구조 유지하면서 .py 파일을 .ipynb로 변환
for root, dirs, files in os.walk(src_root):
    for file in files:
        if file.endswith(".py"):
            src_path = Path(root) / file

            # 상대 경로 계산 (원본 폴더 기준)
            rel_path = src_path.relative_to(src_root)
            dst_dir = dst_root / rel_path.parent
            dst_dir.mkdir(parents=True, exist_ok=True)

            # 출력 .ipynb 파일 경로
            dst_file = dst_dir / (src_path.stem + ".ipynb")

            # ipynb-py-convert 명령어 실행
            subprocess.run([
                "ipynb-py-convert",
                str(src_path),
                str(dst_file)
            ])
