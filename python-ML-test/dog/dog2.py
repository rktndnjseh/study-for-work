import tensorflow as tf
from tensorflow import keras 
from keras import layers, models
import numpy as np
import matplotlib.pyplot as plt
from keras.datasets import cifar10

# 1. CIFAR-10 데이터셋 로드
(train_images, train_labels), (test_images, test_labels) = cifar10.load_data()

# 2. 'cat' (클래스 3)과 'dog' (클래스 5)만 필터링
cat_dog_train_mask = (train_labels.flatten() == 3) | (train_labels.flatten() == 5)
cat_dog_test_mask = (test_labels.flatten() == 3) | (test_labels.flatten() == 5)

train_images = train_images[cat_dog_train_mask]
train_labels = train_labels[cat_dog_train_mask]
test_images = test_images[cat_dog_test_mask]
test_labels = test_labels[cat_dog_test_mask]

# 3. 라벨을 이진화 (cat: 0, dog: 1)
train_labels = np.where(train_labels == 3, 0, 1)
test_labels = np.where(test_labels == 3, 0, 1)

# 4. 데이터 전처리: 픽셀 값을 0~1로 정규화
train_images = train_images.astype('float32') / 255.0
test_images = test_images.astype('float32') / 255.0

# 5. CNN 모델 정의 (이진 분류용)
model = models.Sequential([
    layers.Conv2D(32, (3, 3), activation='relu', input_shape=(32, 32, 3)),
    layers.MaxPooling2D((2, 2)),
    layers.Conv2D(64, (3, 3), activation='relu'),
    layers.MaxPooling2D((2, 2)),
    layers.Conv2D(64, (3, 3), activation='relu'),
    layers.Flatten(),
    layers.Dense(64, activation='relu'),
    layers.Dense(1, activation='sigmoid')  # 이진 분류: 0(cat) 또는 1(dog)
])

# 6. 모델 컴파일
model.compile(optimizer='adam',
              loss='binary_crossentropy',  # 이진 분류 손실 함수
              metrics=['accuracy'])

# 7. 모델 학습
history = model.fit(train_images, train_labels, epochs=5, 
                    validation_data=(test_images, test_labels))

# 8. 모델 평가
test_loss, test_acc = model.evaluate(test_images, test_labels, verbose=2)
print(f"\n테스트 정확도: {test_acc:.4f}")

# 9. 학습 결과 시각화
plt.plot(history.history['accuracy'], label='Training Accuracy')
plt.plot(history.history['val_accuracy'], label='Validation Accuracy')
plt.xlabel('Epoch')
plt.ylabel('Accuracy')
plt.legend()
plt.show()

# 10. 테스트 이미지 예측 및 시각화
predictions = model.predict(test_images[:5])
class_names = ['cat', 'dog']

plt.figure(figsize=(10, 2))
for i in range(5):
    plt.subplot(1, 5, i+1)
    plt.imshow(test_images[i])
    pred_label = 1 if predictions[i] > 0.5 else 0
    plt.title(f"Pred: {class_names[pred_label]}\nTrue: {class_names[int(test_labels[i].item())]}")
    plt.axis('off')
plt.show()