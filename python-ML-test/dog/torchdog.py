import torch
import torch.nn as nn
import torch.optim as optim
import torchvision
import torchvision.transforms as transforms
import matplotlib.pyplot as plt
import numpy as np
from torch.utils.data import DataLoader

# GPU 설정
device = torch.device("cuda" if torch.cuda.is_available() else "cpu")

# CIFAR-10 클래스 이름
class_names = ['airplane', 'automobile', 'bird', 'cat', 'deer',
               'dog', 'frog', 'horse', 'ship', 'truck']

# 데이터 전처리 + 증강
train_transform = transforms.Compose([
    transforms.RandomHorizontalFlip(p=0.5),
    transforms.RandomRotation(10),
    transforms.RandomResizedCrop(32, scale=(0.8, 1.0)),
    transforms.ToTensor(),
    transforms.Normalize(mean=[0.5, 0.5, 0.5], std=[0.5, 0.5, 0.5])
])

test_transform = transforms.Compose([
    transforms.ToTensor(),
    transforms.Normalize(mean=[0.5, 0.5, 0.5], std=[0.5, 0.5, 0.5])
])

# 데이터셋 로드
def load_data():
    trainset = torchvision.datasets.CIFAR10(root='./data', train=True,
                                          download=True, transform=train_transform)
    testset = torchvision.datasets.CIFAR10(root='./data', train=False,
                                         download=True, transform=test_transform)
    return (DataLoader(trainset, batch_size=128, shuffle=True, num_workers=2, pin_memory=True),
            DataLoader(testset, batch_size=128, shuffle=False, num_workers=2, pin_memory=True))

# CNN 모델 정의 (same as before)
class CNN(nn.Module):
    def __init__(self):
        super(CNN, self).__init__()
        self.features = nn.Sequential(
            nn.Conv2d(3, 32, 3, padding=1, bias=False), nn.ReLU(),
            nn.BatchNorm2d(32),
            nn.Conv2d(32, 32, 3, padding=1, bias=False), nn.ReLU(),
            nn.BatchNorm2d(32),
            nn.MaxPool2d(2), nn.Dropout(0.25),
            nn.Conv2d(32, 64, 3, padding=1, bias=False), nn.ReLU(),
            nn.BatchNorm2d(64),
            nn.Conv2d(64, 64, 3, padding=1, bias=False), nn.ReLU(),
            nn.BatchNorm2d(64),
            nn.MaxPool2d(2), nn.Dropout(0.25),
            nn.Conv2d(64, 128, 3, padding=1, bias=False), nn.ReLU(),
            nn.BatchNorm2d(128),
            nn.Conv2d(128, 128, 3, padding=1, bias=False), nn.ReLU(),
            nn.BatchNorm2d(128),
            nn.MaxPool2d(2), nn.Dropout(0.25)
        )
        self.classifier = nn.Sequential(
            nn.Flatten(),
            nn.Linear(128 * 4 * 4, 128, bias=False), nn.ReLU(),
            nn.BatchNorm1d(128),
            nn.Dropout(0.5),
            nn.Linear(128, 10)
        )

    def forward(self, x):
        x = self.features(x)
        return self.classifier(x)

# 학습 및 검증 함수들 (same as before)
def train_epoch(model, loader, criterion, optimizer):
    model.train()
    running_loss = 0.0
    correct = 0
    total = 0
    for inputs, labels in loader:
        inputs, labels = inputs.to(device), labels.to(device)
        optimizer.zero_grad(set_to_none=True)
        outputs = model(inputs)
        loss = criterion(outputs, labels)
        loss.backward()
        optimizer.step()
        running_loss += loss.item()
        _, predicted = outputs.max(1)
        total += labels.size(0)
        correct += predicted.eq(labels).sum().item()
    return running_loss / len(loader), 100. * correct / total

def validate(model, loader, criterion):
    model.eval()
    running_loss = 0.0
    correct = 0
    total = 0
    with torch.no_grad():
        for inputs, labels in loader:
            inputs, labels = inputs.to(device), labels.to(device)
            outputs = model(inputs)
            loss = criterion(outputs, labels)
            running_loss += loss.item()
            _, predicted = outputs.max(1)
            total += labels.size(0)
            correct += predicted.eq(labels).sum().item()
    return running_loss / len(loader), 100. * correct / total

def visualize_predictions(model, loader):
    model.eval()
    images, labels = next(iter(loader))
    images, labels = images.to(device), labels.to(device)
    with torch.no_grad():
        outputs = model(images)
        _, predicted = outputs.max(1)
    print("\n예측 결과:")
    for i in range(5):
        print(f"예측: {class_names[predicted[i]]}, 실제: {class_names[labels[i]]}")
    plt.figure(figsize=(10, 2))
    for i in range(5):
        plt.subplot(1, 5, i+1)
        img = images[i].cpu() / 2 + 0.5
        plt.imshow(np.transpose(img.numpy(), (1, 2, 0)))
        plt.title(f"Pred: {class_names[predicted[i]]}")
        plt.axis('off')
    plt.show()

def main():
    print("Using device:", device)
    trainloader, testloader = load_data()
    
    net = CNN().to(device)
    criterion = nn.CrossEntropyLoss()
    optimizer = optim.Adam(net.parameters(), lr=0.001, weight_decay=1e-4)
    scheduler = optim.lr_scheduler.ReduceLROnPlateau(optimizer, 'min', patience=3, factor=0.5)
    
    epochs = 20
    train_acc_history = []
    val_acc_history = []
    best_val_acc = 0.0
    
    for epoch in range(epochs):
        train_loss, train_acc = train_epoch(net, trainloader, criterion, optimizer)
        val_loss, val_acc = validate(net, testloader, criterion)
        scheduler.step(val_loss)
        train_acc_history.append(train_acc)
        val_acc_history.append(val_acc)
        print(f"Epoch {epoch+1}/{epochs}")
        print(f"Train Loss: {train_loss:.4f} - Train Acc: {train_acc:.2f}%")
        print(f"Val Loss: {val_loss:.4f} - Val Acc: {val_acc:.2f}%\n")
        if val_acc > best_val_acc:
            best_val_acc = val_acc
            torch.save(net.state_dict(), 'best_model.pth')
    
    plt.plot(train_acc_history, label='Train Accuracy')
    plt.plot(val_acc_history, label='Validation Accuracy')
    plt.xlabel('Epoch')
    plt.ylabel('Accuracy (%)')
    plt.legend()
    plt.show()
    
    visualize_predictions(net, testloader)

if __name__ == '__main__':
    main()