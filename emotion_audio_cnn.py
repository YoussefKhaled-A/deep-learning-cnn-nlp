"""
Emotion Detection - Audio CNN (Chroma-CQT)
Dataset: TESS - Toronto Emotional Speech Set (ejlok1/toronto-emotional-speech-set-tess)
Classes: angry, happy, sad (3 classes)
Architecture: CNN on Chroma-CQT features
"""

import os
import glob
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import seaborn as sns
import librosa

import tensorflow as tf
from tensorflow.keras.models import Sequential
from tensorflow.keras.layers import Conv2D, MaxPooling2D, Flatten, Dense, Dropout, BatchNormalization
from tensorflow.keras.utils import to_categorical
from tensorflow.keras.callbacks import EarlyStopping, ReduceLROnPlateau
from sklearn.model_selection import train_test_split
from sklearn.metrics import classification_report, confusion_matrix, f1_score
import kagglehub

# ── CONFIG ──────────────────────────────────────────────────────────────────
DATASET_PATH = kagglehub.dataset_download("ejlok1/toronto-emotional-speech-set-tess")
TESS_DIR     = os.path.join(DATASET_PATH, "TESS Toronto emotional speech set data")
CLASSES      = ["angry", "happy", "sad"]
LABEL2ID     = {c: i for i, c in enumerate(CLASSES)}

# Chroma-CQT params (from Lab 3)
SR          = 22050
HOP_LENGTH  = 512
N_CHROMA    = 12
MAX_FRAMES  = 174   # ~4 seconds at 22050 Hz / 512 hop
BATCH       = 32
EPOCHS      = 30
OUT_DIR     = os.path.dirname(os.path.abspath(__file__))

print(f"Dataset path: {TESS_DIR}")
print(f"Classes: {CLASSES}")

# ── LOAD AUDIO & EXTRACT CHROMA-CQT ─────────────────────────────────────────
def extract_chroma_cqt(filepath, sr=SR, hop_length=HOP_LENGTH,
                        n_chroma=N_CHROMA, max_frames=MAX_FRAMES):
    y, _ = librosa.load(filepath, sr=sr)
    chroma = librosa.feature.chroma_cqt(y=y, sr=sr, hop_length=hop_length, n_chroma=n_chroma)
    # Normalize
    mean = chroma.mean(axis=1, keepdims=True)
    std  = chroma.std(axis=1, keepdims=True)
    chroma = (chroma - mean) / (std + 1e-9)
    # Pad or truncate to max_frames
    if chroma.shape[1] < max_frames:
        pad = max_frames - chroma.shape[1]
        chroma = np.pad(chroma, ((0, 0), (0, pad)), mode='constant')
    else:
        chroma = chroma[:, :max_frames]
    return chroma

X, y = [], []
total_files = 0

for folder in os.listdir(TESS_DIR):
    folder_lower = folder.lower()
    emotion = None
    for cls in CLASSES:
        if cls in folder_lower:
            emotion = cls
            break
    if emotion is None:
        continue

    folder_path = os.path.join(TESS_DIR, folder)
    wav_files = glob.glob(os.path.join(folder_path, "*.wav"))
    print(f"  {folder}: {len(wav_files)} files -> '{emotion}'")

    for wav in wav_files:
        try:
            feat = extract_chroma_cqt(wav)
            X.append(feat)
            y.append(LABEL2ID[emotion])
            total_files += 1
        except Exception as e:
            print(f"    Error loading {wav}: {e}")

print(f"\nTotal files loaded: {total_files}")
print(f"Class distribution: {dict(zip(*np.unique(y, return_counts=True)))}")

X = np.array(X)  # shape: (N, 12, 174)
y = np.array(y)

# Add channel dimension for CNN: (N, 12, 174, 1)
X = X[..., np.newaxis]
print(f"X shape: {X.shape}  |  y shape: {y.shape}")

# ── SPLIT ────────────────────────────────────────────────────────────────────
X_train, X_temp, y_train, y_temp = train_test_split(X, y, test_size=0.3, stratify=y, random_state=42)
X_val, X_test, y_val, y_test     = train_test_split(X_temp, y_temp, test_size=0.5, stratify=y_temp, random_state=42)

y_train_cat = to_categorical(y_train, num_classes=len(CLASSES))
y_val_cat   = to_categorical(y_val,   num_classes=len(CLASSES))
y_test_cat  = to_categorical(y_test,  num_classes=len(CLASSES))

print(f"Train: {X_train.shape}  |  Val: {X_val.shape}  |  Test: {X_test.shape}")

# ── MODEL ────────────────────────────────────────────────────────────────────
model = Sequential([
    Conv2D(32, (3,3), activation='relu', padding='same', input_shape=(N_CHROMA, MAX_FRAMES, 1)),
    BatchNormalization(),
    MaxPooling2D(2,2),
    Dropout(0.25),

    Conv2D(64, (3,3), activation='relu', padding='same'),
    BatchNormalization(),
    MaxPooling2D(2,2),
    Dropout(0.25),

    Conv2D(128, (3,3), activation='relu', padding='same'),
    BatchNormalization(),
    MaxPooling2D((1,2)),
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
    EarlyStopping(monitor='val_loss', patience=7, restore_best_weights=True),
    ReduceLROnPlateau(monitor='val_loss', factor=0.5, patience=4, min_lr=1e-6)
]

history = model.fit(
    X_train, y_train_cat,
    validation_data=(X_val, y_val_cat),
    epochs=EPOCHS,
    batch_size=BATCH,
    callbacks=callbacks
)

# ── ACCURACY / LOSS CURVES ───────────────────────────────────────────────────
fig, axes = plt.subplots(1, 2, figsize=(12, 5))

axes[0].plot(history.history['accuracy'], label='Train Accuracy')
axes[0].plot(history.history['val_accuracy'], label='Val Accuracy')
axes[0].set_title('Audio CNN (Chroma-CQT) — Accuracy')
axes[0].set_xlabel('Epoch')
axes[0].set_ylabel('Accuracy')
axes[0].legend()
axes[0].grid(True)

axes[1].plot(history.history['loss'], label='Train Loss')
axes[1].plot(history.history['val_loss'], label='Val Loss')
axes[1].set_title('Audio CNN (Chroma-CQT) — Loss')
axes[1].set_xlabel('Epoch')
axes[1].set_ylabel('Loss')
axes[1].legend()
axes[1].grid(True)

plt.tight_layout()
plt.savefig(os.path.join(OUT_DIR, 'audio_accuracy_loss.png'), dpi=150)
plt.close()
print("Saved: audio_accuracy_loss.png")

# ── EVALUATE ─────────────────────────────────────────────────────────────────
test_loss, test_acc = model.evaluate(X_test, y_test_cat, verbose=0)
print(f"\nTest Accuracy: {test_acc:.4f}  |  Test Loss: {test_loss:.4f}")

# ── CONFUSION MATRIX ─────────────────────────────────────────────────────────
y_pred = np.argmax(model.predict(X_test, verbose=0), axis=1)

cm = confusion_matrix(y_test, y_pred)
plt.figure(figsize=(8, 6))
sns.heatmap(cm, annot=True, fmt='d', cmap='Oranges',
            xticklabels=CLASSES, yticklabels=CLASSES)
plt.title('Audio CNN — Confusion Matrix')
plt.xlabel('Predicted')
plt.ylabel('True')
plt.tight_layout()
plt.savefig(os.path.join(OUT_DIR, 'audio_confusion_matrix.png'), dpi=150)
plt.close()
print("Saved: audio_confusion_matrix.png")

# ── CLASSIFICATION REPORT ────────────────────────────────────────────────────
report = classification_report(y_test, y_pred, target_names=CLASSES)
print("\nClassification Report:")
print(report)

macro_f1 = f1_score(y_test, y_pred, average='macro')
print(f"Macro F1 Score: {macro_f1:.4f}")

# ── SAMPLE PREDICTION ────────────────────────────────────────────────────────
sample_idx  = 0
sample_feat = X_test[sample_idx]
sample_true = CLASSES[y_test[sample_idx]]
pred_probs  = model.predict(sample_feat[np.newaxis, ...], verbose=0)[0]
sample_pred = CLASSES[np.argmax(pred_probs)]

plt.figure(figsize=(10, 3))
plt.imshow(sample_feat.squeeze(), aspect='auto', origin='lower', cmap='viridis')
plt.colorbar(label='Normalized Chroma Energy')
plt.title(f"Sample Chroma-CQT | True: {sample_true} | Pred: {sample_pred}\n"
          + " | ".join([f"{c}: {p:.2f}" for c, p in zip(CLASSES, pred_probs)]))
plt.xlabel('Time Frame')
plt.ylabel('Chroma Bin')
plt.tight_layout()
plt.savefig(os.path.join(OUT_DIR, 'audio_sample_prediction.png'), dpi=150)
plt.close()
print("Saved: audio_sample_prediction.png")

print("\n=== Audio CNN Complete ===")
print(f"Final Test Accuracy: {test_acc:.4f}")
print(f"Macro F1: {macro_f1:.4f}")
