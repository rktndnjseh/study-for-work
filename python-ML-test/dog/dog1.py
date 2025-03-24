import tensorflow as tf
from tensorflow import keras 
from keras import datasets, layers, models
import matplotlib.pyplot as plt
import numpy as np

# 1. CIFAR-10 데이터셋 로드
(train_images, train_labels), (test_images, test_labels) = datasets.cifar10.load_data()

# 2. 데이터 전처리: 픽셀 값을 0~1 사이로 정규화
train_images = train_images.astype('float32') / 255.0
test_images = test_images.astype('float32') / 255.0

# 클래스 이름 정의
class_names = ['airplane', 'automobile', 'bird', 'cat', 'deer', 
               'dog', 'frog', 'horse', 'ship', 'truck']

# 3. CNN 모델 정의
model = models.Sequential([
    layers.Conv2D(32, (3, 3), activation='relu', input_shape=(32, 32, 3)),
    layers.MaxPooling2D((2, 2)),
    layers.Conv2D(64, (3, 3), activation='relu'),
    layers.MaxPooling2D((2, 2)),
    layers.Conv2D(64, (3, 3), activation='relu'),
    layers.Flatten(),
    layers.Dense(64, activation='relu'),
    layers.Dense(10, activation='softmax')  # 10개 클래스 출력
])

# 4. 모델 컴파일
model.compile(optimizer='adam',
              loss='sparse_categorical_crossentropy',
              metrics=['accuracy'])

# 5. 모델 학습
history = model.fit(train_images, train_labels, epochs=5, 
                    validation_data=(test_images, test_labels))

# 6. 모델 평가
test_loss, test_acc = model.evaluate(test_images, test_labels, verbose=2)
print(f"\n테스트 정확도: {test_acc:.4f}")

# 7. 학습 결과 시각화
plt.plot(history.history['accuracy'], label='Training Accuracy')
plt.plot(history.history['val_accuracy'], label='Validation Accuracy')
plt.xlabel('Epoch')
plt.ylabel('Accuracy')
plt.legend()
plt.show()

# 8. 테스트 이미지 예측
predictions = model.predict(test_images[:5])
for i in range(5):
    pred_label = np.argmax(predictions[i])
    true_label = test_labels[i][0]
    print(f"예측: {class_names[pred_label]}, 실제: {class_names[true_label]}")

# 9. 테스트 이미지 시각화
plt.figure(figsize=(10, 2))
for i in range(5):
    plt.subplot(1, 5, i+1)
    plt.imshow(test_images[i])
    plt.title(f"Pred: {class_names[np.argmax(predictions[i])]}")
    plt.axis('off')
plt.show()