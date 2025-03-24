import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

# 한글 설정
plt.rcParams['font.family'] = 'Malgun Gothic'
plt.rcParams['axes.unicode_minus'] = False

# 데이터 불러오기 및 전처리
file_path = "C:/Users/user/Desktop/python-test/water.csv"
raw_df = pd.read_csv(file_path, skiprows=1)
columns = raw_df.iloc[0, 1:].tolist()
data = raw_df.iloc[1:, :].copy()
data.columns = ['일'] + columns
data = data.dropna(how='all').reset_index(drop=True)

# 문자열 -> 숫자 변환
for col in data.columns[1:]:
    data[col] = data[col].str.replace(",", "").replace("-", np.nan).astype(float)

# 분석
monthly_total = data.iloc[:, 1:].sum()
monthly_mean = data.iloc[:, 1:].mean()
year_total = monthly_total.sum()
year_mean = monthly_mean.mean()

# ✅ 하나의 창에 두 개의 subplot
fig, axes = plt.subplots(2, 1, figsize=(12, 10))  # (행, 열), 창 크기

# 1. 월별 총 공급량
monthly_total.plot(kind='bar', ax=axes[0], color='skyblue')
axes[0].set_title("월별 총 공급량")
axes[0].set_xlabel("월")
axes[0].set_ylabel("공급량 (톤)")
axes[0].grid(True)

# 2. 월별 평균 공급량
monthly_mean.plot(kind='line', marker='o', linestyle='--', ax=axes[1], color='orange')
axes[1].set_title("월별 평균 공급량")
axes[1].set_xlabel("월")
axes[1].set_ylabel("평균 공급량 (톤)")
axes[1].grid(True)

plt.tight_layout()
plt.show()

# 연간 요약 출력
print(f"✅ 연간 총 공급량: {year_total:,.2f} 톤")
print(f"✅ 연간 평균 공급량: {year_mean:,.2f} 톤")
