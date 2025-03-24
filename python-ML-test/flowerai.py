from sklearn.datasets import load_iris
from sklearn.model_selection import train_test_split
from sklearn.ensemble import RandomForestClassifier
from sklearn.metrics import accuracy_score, classification_report
import pandas as pd

# 1. 데이터 불러오기
iris = load_iris()
X = iris.data
y = iris.target
feature_names = iris.feature_names
target_names = iris.target_names

# 2. 학습/테스트 데이터 나누기
X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.3, random_state=42)

# 3. 모델 훈련
model = RandomForestClassifier(random_state=42)
model.fit(X_train, y_train)

# 4. 예측 및 평가
y_pred = model.predict(X_test)

# 5. 결과 출력
print("✅ 정확도:", accuracy_score(y_test, y_pred))
print("\n📊 분류 리포트:\n", classification_report(y_test, y_pred, target_names=target_names))

# 6. 예측 예시
example = [[5.1, 3.5, 1.4, 0.2]]  # 꽃잎, 꽃받침의 길이/너비
predicted = model.predict(example)
print(f"\n🌸 예측 결과: {target_names[predicted[0]]}")
