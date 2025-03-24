from sklearn.datasets import load_breast_cancer
from sklearn.model_selection import train_test_split
from sklearn.preprocessing import StandardScaler
from sklearn.linear_model import LogisticRegression
from sklearn.metrics import (
    accuracy_score, confusion_matrix, classification_report,
    roc_curve, auc, ConfusionMatrixDisplay
)
import matplotlib.pyplot as plt
import seaborn as sns

# 한글 깨짐 방지 (필요시)
plt.rcParams['font.family'] = 'Malgun Gothic'
plt.rcParams['axes.unicode_minus'] = False

# 1. 데이터셋 불러오기
data = load_breast_cancer()
X = data.data
y = data.target

# 2. 정규화
scaler = StandardScaler()
X_scaled = scaler.fit_transform(X)

# 3. 데이터 분할
X_train, X_test, y_train, y_test = train_test_split(X_scaled, y, test_size=0.3, random_state=42)

# 4. 로지스틱 회귀 모델 학습
model = LogisticRegression()
model.fit(X_train, y_train)
y_pred = model.predict(X_test)
y_prob = model.predict_proba(X_test)[:, 1]

# 5. 정확도 및 리포트 출력
print(f"✅ 정확도: {accuracy_score(y_test, y_pred):.4f}")
print("\n📋 분류 리포트:\n")
print(classification_report(y_test, y_pred, target_names=data.target_names))

# 6. 혼동 행렬 시각화
cm = confusion_matrix(y_test, y_pred)
disp = ConfusionMatrixDisplay(confusion_matrix=cm, display_labels=data.target_names)
disp.plot(cmap="Blues")
plt.title("🧩 혼동 행렬")
plt.show()

# 7. ROC 곡선 그리기
fpr, tpr, thresholds = roc_curve(y_test, y_prob)
roc_auc = auc(fpr, tpr)

plt.figure(figsize=(8, 6))
plt.plot(fpr, tpr, label=f'ROC 곡선 (AUC = {roc_auc:.2f})', color='darkorange')
plt.plot([0, 1], [0, 1], linestyle='--', color='gray')
plt.xlabel('False Positive Rate')
plt.ylabel('True Positive Rate')
plt.title('🎯 ROC 곡선')
plt.legend(loc='lower right')
plt.grid(True)
plt.show()
