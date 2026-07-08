"""
Emotion Detection - Image CNN
Dataset: FER2013 (msambare/fer2013) - Facial Expression Recognition
Classes: angry, happy, sad (3 classes)
Architecture: CNN (Conv2D + MaxPool + Dropout + Dense)
"""

import os
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import seaborn as sns

from tensorflow.keras.models import Sequential
from tensorflow.keras.layers import Conv2D, MaxPooling2D, Flatten, Dense, Dropout, BatchNormalization
from tensorflow.keras.preprocessing.image import ImageDataGenerator
from tensorflow.keras.callbacks import EarlyStopping, ReduceLROnPlateau
from sklearn.metrics import classification_report, confusion_matrix

# ── CONFIG ──────────────────────────────────────────────────────────────────
import kagglehub
DATASET_PATH = kagglehub.dataset_download("msambare/fer2013")

TRAIN_DIR = os.path.join(DATASET_PATH, "train")
TEST_DIR  = os.path.join(DATASET_PATH, "test")
CLASSES   = ["angry", "happy", "sad"]
IMG_SIZE  = 48
BATCH     = 64
EPOCHS    = 25
OUT_DIR   = os.path.dirname(os.path.abspath(__file__))

print(f"Dataset path: {DATASET_PATH}")
print(f"Classes: {CLASSES}")

# ── DATA GENERATORS ──────────────────────────────────────────────────────────
train_datagen = ImageDataGenerator(
    rescale=1.0/255,
    rotation_range=10,
    width_shift_range=0.1,
    height_shift_range=0.1,
    horizontal_flip=True,
    zoom_range=0.1,
    validation_split=0.15
)

test_datagen = ImageDataGenerator(rescale=1.0/255)

train_gen = train_datagen.flow_from_directory(
    TRAIN_DIR,
    target_size=(IMG_SIZE, IMG_SIZE),
    color_mode="grayscale",
    classes=CLASSES,
    class_mode="categorical",
    batch_size=BATCH,
    subset="training",
    shuffle=True
)

val_gen = train_datagen.flow_from_directory(
    TRAIN_DIR,
    target_size=(IMG_SIZE, IMG_SIZE),
    color_mode="grayscale",
    classes=CLASSES,
    class_mode="categorical",
    batch_size=BATCH,
    subset="validation",
    shuffle=False
)

test_gen = test_datagen.flow_from_directory(
    TEST_DIR,
    target_size=(IMG_SIZE, IMG_SIZE),
    color_mode="grayscale",
    classes=CLASSES,
    class_mode="categorical",
    batch_size=BATCH,
    shuffle=False
)

print(f"\nTrain samples: {train_gen.samples}")
print(f"Val   samples: {val_gen.samples}")
print(f"Test  samples: {test_gen.samples}")

# ── MODEL ────────────────────────────────────────────────────────────────────
model = Sequential([
    Conv2D(32, (3,3), activation='relu', padding='same', input_shape=(IMG_SIZE, IMG_SIZE, 1)),
    BatchNormalization(),
    MaxPooling2D(2,2),
    Dropout(0.25),

    Conv2D(64, (3,3), activation='relu', padding='same'),
    BatchNormalization(),
    MaxPooling2D(2,2),
    Dropout(0.25),

    Conv2D(128, (3,3), activation='relu', padding='same'),
    BatchNormalization(),
    MaxPooling2D(2,2),
    Dropout(0.25),

    Flatten(),
    Dense(256, activation='relu'),
    BatchNormalization(),
    Dropout(0.5),
    Dense(len(CLASSES), activation='softmax')
])

model.compile(optimizer='adam', loss='categorical_crossentropy', metrics=['accuracy'])
model.summary()

# ── TRAIN ────────────────────────────────────────────────────────────────────
callbacks = [
    EarlyStopping(monitor='val_loss', patience=5, restore_best_weights=True),
    ReduceLROnPlateau(monitor='val_loss', factor=0.5, patience=3, min_lr=1e-6)
]

history = model.fit(
    train_gen,
    validation_data=val_gen,
    epochs=EPOCHS,
    callbacks=callbacks
)

# ── ACCURACY / LOSS CURVES ───────────────────────────────────────────────────
fig, axes = plt.subplots(1, 2, figsize=(12, 5))

axes[0].plot(history.history['accuracy'], label='Train Accuracy')
axes[0].plot(history.history['val_accuracy'], label='Val Accuracy')
axes[0].set_title('Image CNN — Accuracy')
axes[0].set_xlabel('Epoch')
axes[0].set_ylabel('Accuracy')
axes[0].legend()
axes[0].grid(True)

axes[1].plot(history.history['loss'], label='Train Loss')
axes[1].plot(history.history['val_loss'], label='Val Loss')
axes[1].set_title('Image CNN — Loss')
axes[1].set_xlabel('Epoch')
axes[1].set_ylabel('Loss')
axes[1].legend()
axes[1].grid(True)

plt.tight_layout()
plt.savefig(os.path.join(OUT_DIR, 'image_accuracy_loss.png'), dpi=150)
plt.close()
print("Saved: image_accuracy_loss.png")

# ── EVALUATE ─────────────────────────────────────────────────────────────────
test_loss, test_acc = model.evaluate(test_gen, verbose=0)
print(f"\nTest Accuracy: {test_acc:.4f}  |  Test Loss: {test_loss:.4f}")

# ── CONFUSION MATRIX ─────────────────────────────────────────────────────────
test_gen.reset()
y_pred_probs = model.predict(test_gen, verbose=0)
y_pred = np.argmax(y_pred_probs, axis=1)
y_true = test_gen.classes

cm = confusion_matrix(y_true, y_pred)
plt.figure(figsize=(8, 6))
sns.heatmap(cm, annot=True, fmt='d', cmap='Blues',
            xticklabels=CLASSES, yticklabels=CLASSES)
plt.title('Image CNN — Confusion Matrix')
plt.xlabel('Predicted')
plt.ylabel('True')
plt.tight_layout()
plt.savefig(os.path.join(OUT_DIR, 'image_confusion_matrix.png'), dpi=150)
plt.close()
print("Saved: image_confusion_matrix.png")

# ── CLASSIFICATION REPORT ────────────────────────────────────────────────────
report = classification_report(y_true, y_pred, target_names=CLASSES)
print("\nClassification Report:")
print(report)

# Macro F1
from sklearn.metrics import f1_score
macro_f1 = f1_score(y_true, y_pred, average='macro')
print(f"Macro F1 Score: {macro_f1:.4f}")

# ── SAMPLE PREDICTION ────────────────────────────────────────────────────────
test_gen.reset()
batch_imgs, batch_labels = next(test_gen)
sample_img = batch_imgs[0]
sample_true = CLASSES[np.argmax(batch_labels[0])]
pred_probs = model.predict(sample_img[np.newaxis, ...], verbose=0)[0]
sample_pred = CLASSES[np.argmax(pred_probs)]

plt.figure(figsize=(4, 4))
plt.imshow(sample_img.squeeze(), cmap='gray')
plt.title(f"True: {sample_true} | Pred: {sample_pred}\n"
          + " | ".join([f"{c}: {p:.2f}" for c, p in zip(CLASSES, pred_probs)]))
plt.axis('off')
plt.tight_layout()
plt.savefig(os.path.join(OUT_DIR, 'image_sample_prediction.png'), dpi=150)
plt.close()
print("Saved: image_sample_prediction.png")

print("\n=== Image CNN Complete ===")
print(f"Final Test Accuracy: {test_acc:.4f}")
print(f"Macro F1: {macro_f1:.4f}")
