"""
Emotion Detection - Text BERT
Dataset: Tweet Emotions (pashupatigupta/emotion-detection-from-text)
Classes: happiness, sadness, anger (3 classes)
Architecture: BERT (bert-base-uncased) fine-tuned for classification
"""

import os
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import seaborn as sns

import torch
from torch.utils.data import Dataset, DataLoader
from transformers import BertTokenizer, BertForSequenceClassification
from transformers import get_linear_schedule_with_warmup
from torch.optim import AdamW
from sklearn.model_selection import train_test_split
from sklearn.metrics import classification_report, confusion_matrix, f1_score
import kagglehub

# ── CONFIG ──────────────────────────────────────────────────────────────────
DATASET_PATH = kagglehub.dataset_download("pashupatigupta/emotion-detection-from-text")
CSV_FILE     = os.path.join(DATASET_PATH, "tweet_emotions.csv")
CLASSES      = ["happiness", "sadness", "anger"]
LABEL2ID     = {c: i for i, c in enumerate(CLASSES)}
MAX_LEN      = 128
BATCH        = 32
EPOCHS       = 5
LR           = 2e-5
DEVICE       = torch.device("cuda" if torch.cuda.is_available() else "cpu")
OUT_DIR      = os.path.dirname(os.path.abspath(__file__))

print(f"Device: {DEVICE}")
print(f"Classes: {CLASSES}")

# ── LOAD & FILTER DATA ───────────────────────────────────────────────────────
df = pd.read_csv(CSV_FILE)
df = df[df['sentiment'].isin(CLASSES)][['content', 'sentiment']].dropna()
df['label'] = df['sentiment'].map(LABEL2ID)

print(f"\nClass distribution:")
print(df['sentiment'].value_counts())
print(f"Total samples: {len(df)}")

# Balance classes
min_count = df['sentiment'].value_counts().min()
df_balanced = df.groupby('sentiment').sample(n=min_count, random_state=42).reset_index(drop=True)
print(f"\nBalanced dataset: {len(df_balanced)} samples ({min_count} per class)")

train_df, temp_df = train_test_split(df_balanced, test_size=0.3, stratify=df_balanced['label'], random_state=42)
val_df, test_df   = train_test_split(temp_df, test_size=0.5, stratify=temp_df['label'], random_state=42)
print(f"Train: {len(train_df)} | Val: {len(val_df)} | Test: {len(test_df)}")

# ── TOKENIZER & DATASET ──────────────────────────────────────────────────────
tokenizer = BertTokenizer.from_pretrained("bert-base-uncased")

class EmotionDataset(Dataset):
    def __init__(self, texts, labels, tokenizer, max_len):
        self.texts     = texts.reset_index(drop=True)
        self.labels    = labels.reset_index(drop=True)
        self.tokenizer = tokenizer
        self.max_len   = max_len

    def __len__(self):
        return len(self.texts)

    def __getitem__(self, idx):
        enc = self.tokenizer(
            str(self.texts[idx]),
            max_length=self.max_len,
            padding='max_length',
            truncation=True,
            return_tensors='pt'
        )
        return {
            'input_ids':      enc['input_ids'].squeeze(0),
            'attention_mask': enc['attention_mask'].squeeze(0),
            'label':          torch.tensor(self.labels[idx], dtype=torch.long)
        }

train_ds = EmotionDataset(train_df['content'], train_df['label'], tokenizer, MAX_LEN)
val_ds   = EmotionDataset(val_df['content'],   val_df['label'],   tokenizer, MAX_LEN)
test_ds  = EmotionDataset(test_df['content'],  test_df['label'],  tokenizer, MAX_LEN)

train_loader = DataLoader(train_ds, batch_size=BATCH, shuffle=True)
val_loader   = DataLoader(val_ds,   batch_size=BATCH)
test_loader  = DataLoader(test_ds,  batch_size=BATCH)

# ── MODEL ────────────────────────────────────────────────────────────────────
model = BertForSequenceClassification.from_pretrained(
    "bert-base-uncased",
    num_labels=len(CLASSES),
    ignore_mismatched_sizes=True
).to(DEVICE)

optimizer = AdamW(model.parameters(), lr=LR, weight_decay=0.01)
total_steps = len(train_loader) * EPOCHS
scheduler = get_linear_schedule_with_warmup(
    optimizer,
    num_warmup_steps=total_steps // 10,
    num_training_steps=total_steps
)

# ── TRAIN ────────────────────────────────────────────────────────────────────
train_losses, val_losses = [], []
train_accs,   val_accs   = [], []

for epoch in range(1, EPOCHS + 1):
    # --- Train ---
    model.train()
    total_loss, correct, total = 0.0, 0, 0
    for batch in train_loader:
        input_ids      = batch['input_ids'].to(DEVICE)
        attention_mask = batch['attention_mask'].to(DEVICE)
        labels         = batch['label'].to(DEVICE)

        optimizer.zero_grad()
        outputs = model(input_ids=input_ids, attention_mask=attention_mask, labels=labels)
        loss    = outputs.loss
        loss.backward()
        torch.nn.utils.clip_grad_norm_(model.parameters(), 1.0)
        optimizer.step()
        scheduler.step()

        total_loss += loss.item()
        preds = outputs.logits.argmax(dim=1)
        correct += (preds == labels).sum().item()
        total   += labels.size(0)

    train_loss = total_loss / len(train_loader)
    train_acc  = correct / total
    train_losses.append(train_loss)
    train_accs.append(train_acc)

    # --- Validate ---
    model.eval()
    total_loss, correct, total = 0.0, 0, 0
    with torch.no_grad():
        for batch in val_loader:
            input_ids      = batch['input_ids'].to(DEVICE)
            attention_mask = batch['attention_mask'].to(DEVICE)
            labels         = batch['label'].to(DEVICE)
            outputs = model(input_ids=input_ids, attention_mask=attention_mask, labels=labels)
            total_loss += outputs.loss.item()
            preds = outputs.logits.argmax(dim=1)
            correct += (preds == labels).sum().item()
            total   += labels.size(0)

    val_loss = total_loss / len(val_loader)
    val_acc  = correct / total
    val_losses.append(val_loss)
    val_accs.append(val_acc)

    print(f"Epoch {epoch}/{EPOCHS}  "
          f"Train Loss: {train_loss:.4f}  Train Acc: {train_acc:.4f}  "
          f"Val Loss: {val_loss:.4f}  Val Acc: {val_acc:.4f}")

# ── ACCURACY / LOSS CURVES ───────────────────────────────────────────────────
fig, axes = plt.subplots(1, 2, figsize=(12, 5))

axes[0].plot(train_accs, label='Train Accuracy')
axes[0].plot(val_accs,   label='Val Accuracy')
axes[0].set_title('Text BERT — Accuracy')
axes[0].set_xlabel('Epoch')
axes[0].set_ylabel('Accuracy')
axes[0].legend()
axes[0].grid(True)

axes[1].plot(train_losses, label='Train Loss')
axes[1].plot(val_losses,   label='Val Loss')
axes[1].set_title('Text BERT — Loss')
axes[1].set_xlabel('Epoch')
axes[1].set_ylabel('Loss')
axes[1].legend()
axes[1].grid(True)

plt.tight_layout()
plt.savefig(os.path.join(OUT_DIR, 'text_accuracy_loss.png'), dpi=150)
plt.close()
print("Saved: text_accuracy_loss.png")

# ── EVALUATE ─────────────────────────────────────────────────────────────────
model.eval()
all_preds, all_labels = [], []
with torch.no_grad():
    for batch in test_loader:
        input_ids      = batch['input_ids'].to(DEVICE)
        attention_mask = batch['attention_mask'].to(DEVICE)
        labels         = batch['label'].to(DEVICE)
        outputs = model(input_ids=input_ids, attention_mask=attention_mask)
        preds   = outputs.logits.argmax(dim=1)
        all_preds.extend(preds.cpu().numpy())
        all_labels.extend(labels.cpu().numpy())

# ── CONFUSION MATRIX ─────────────────────────────────────────────────────────
cm = confusion_matrix(all_labels, all_preds)
plt.figure(figsize=(8, 6))
sns.heatmap(cm, annot=True, fmt='d', cmap='Greens',
            xticklabels=CLASSES, yticklabels=CLASSES)
plt.title('Text BERT — Confusion Matrix')
plt.xlabel('Predicted')
plt.ylabel('True')
plt.tight_layout()
plt.savefig(os.path.join(OUT_DIR, 'text_confusion_matrix.png'), dpi=150)
plt.close()
print("Saved: text_confusion_matrix.png")

# ── CLASSIFICATION REPORT ────────────────────────────────────────────────────
report = classification_report(all_labels, all_preds, target_names=CLASSES)
print("\nClassification Report:")
print(report)

macro_f1 = f1_score(all_labels, all_preds, average='macro')
test_acc  = np.mean(np.array(all_preds) == np.array(all_labels))
print(f"Test Accuracy: {test_acc:.4f}")
print(f"Macro F1 Score: {macro_f1:.4f}")

# ── SAMPLE PREDICTION ────────────────────────────────────────────────────────
sample_text  = test_df['content'].iloc[0]
sample_label = CLASSES[test_df['label'].iloc[0]]
enc = tokenizer(sample_text, max_length=MAX_LEN, padding='max_length',
                truncation=True, return_tensors='pt')
with torch.no_grad():
    logits = model(input_ids=enc['input_ids'].to(DEVICE),
                   attention_mask=enc['attention_mask'].to(DEVICE)).logits
probs = torch.softmax(logits, dim=1).cpu().numpy()[0]
pred_label = CLASSES[np.argmax(probs)]

print(f"\nSample Prediction:")
print(f"  Text:  {sample_text[:100]}...")
print(f"  True:  {sample_label}")
print(f"  Pred:  {pred_label}")
print(f"  Probs: " + " | ".join([f"{c}: {p:.3f}" for c, p in zip(CLASSES, probs)]))

print("\n=== Text BERT Complete ===")
print(f"Final Test Accuracy: {test_acc:.4f}")
print(f"Macro F1: {macro_f1:.4f}")
