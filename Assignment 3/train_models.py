import os
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from sklearn.model_selection import train_test_split
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.linear_model import LogisticRegression
from sklearn.metrics import classification_report, accuracy_score, precision_recall_fscore_support, confusion_matrix
import torch
import torch.nn as nn
import torch.optim as optim
from torch.utils.data import Dataset, DataLoader
from transformers import BertTokenizer, BertForSequenceClassification
import re
import unicodedata
import emoji
import warnings
warnings.filterwarnings('ignore')

# Set device
device = torch.device('cuda' if torch.cuda.is_available() else 'cpu')
print(f"Using device: {device}")

# Map sentiments to integers
LABEL_MAP = {'negative': 0, 'neutral': 1, 'positive': 2}
INV_LABEL_MAP = {0: 'negative', 1: 'neutral', 2: 'positive'}

# Preprocessing from explore_dataset
def normalize_unicode(text):
    return unicodedata.normalize('NFC', str(text))

def convert_emojis(text):
    return emoji.demojize(text, delimiters=(" ", " "))

def remove_urls(text):
    return re.sub(r'http\S+|www\.\S+', '', text)

def clean_mentions_hashtags(text):
    text = re.sub(r'@\w+', '', text)
    text = re.sub(r'#(\w+)', r'\1', text)
    return text

def lowercase_english(text):
    return text.lower()

def remove_noise(text):
    text = re.sub(r'<.*?>', '', text)
    text = re.sub(r'[^\w\s\u0600-\u06FF!?]', ' ', text)
    return text

def preprocess_text(text):
    text = normalize_unicode(text)
    text = convert_emojis(text)
    text = remove_urls(text)
    text = clean_mentions_hashtags(text)
    text = lowercase_english(text)
    text = remove_noise(text)
    # Return space-separated tokens
    tokens = text.split()
    return " ".join(tokens)

# ============================================================
# PYTORCH DATASET AND MODEL FOR BiLSTM
# ============================================================
class Vocabulary:
    def __init__(self, max_size=10000, min_freq=2):
        self.max_size = max_size
        self.min_freq = min_freq
        self.word2idx = {'<PAD>': 0, '<UNK>': 1}
        self.idx2word = {0: '<PAD>', 1: '<UNK>'}
        self.vocab_size = 2
        
    def build_vocab(self, sentences):
        word_counts = {}
        for sentence in sentences:
            for word in sentence.split():
                word_counts[word] = word_counts.get(word, 0) + 1
                
        # Filter by freq
        sorted_words = sorted([w for w, c in word_counts.items() if c >= self.min_freq], 
                              key=lambda x: word_counts[x], reverse=True)
        
        for word in sorted_words[:self.max_size]:
            self.word2idx[word] = self.vocab_size
            self.idx2word[self.vocab_size] = word
            self.vocab_size += 1
            
    def numericalize(self, sentence):
        return [self.word2idx.get(word, self.word2idx['<UNK>']) for word in sentence.split()]

class TextDataset(Dataset):
    def __init__(self, texts, labels, vocab, max_len=60):
        self.texts = texts
        self.labels = labels
        self.vocab = vocab
        self.max_len = max_len
        
    def __len__(self):
        return len(self.texts)
        
    def __getitem__(self, idx):
        text = self.texts[idx]
        label = self.labels[idx]
        
        numericalized = self.vocab.numericalize(text)
        
        # Pad or truncate
        if len(numericalized) < self.max_len:
            numericalized += [self.vocab.word2idx['<PAD>']] * (self.max_len - len(numericalized))
        else:
            numericalized = numericalized[:self.max_len]
            
        return torch.tensor(numericalized), torch.tensor(label)

class BiLSTMClassifier(nn.Module):
    def __init__(self, vocab_size, embedding_dim=100, hidden_dim=128, output_dim=3, num_layers=2, dropout=0.3):
        super(BiLSTMClassifier, self).__init__()
        self.embedding = nn.Embedding(vocab_size, embedding_dim, padding_idx=0)
        self.lstm = nn.LSTM(embedding_dim, hidden_dim, num_layers=num_layers, 
                            bidirectional=True, batch_first=True, dropout=dropout)
        self.fc = nn.Linear(hidden_dim * 2, output_dim)
        self.dropout = nn.Dropout(dropout)
        
    def forward(self, text):
        embedded = self.dropout(self.embedding(text))
        # lstm output shape: [batch_size, seq_len, hidden_dim * 2]
        # hidden state shape: [num_layers * 2, batch_size, hidden_dim]
        output, (hidden, cell) = self.lstm(embedded)
        
        # Concatenate final forward and backward hidden states
        hidden = self.dropout(torch.cat((hidden[-2,:,:], hidden[-1,:,:]), dim=1))
        return self.fc(hidden)

# ============================================================
# MBERT DATASET AND FINE-TUNING UTILS
# ============================================================
class BERTDataset(Dataset):
    def __init__(self, texts, labels, tokenizer, max_len=60):
        self.texts = texts
        self.labels = labels
        self.tokenizer = tokenizer
        self.max_len = max_len
        
    def __len__(self):
        return len(self.texts)
        
    def __getitem__(self, idx):
        text = str(self.texts[idx])
        label = self.labels[idx]
        
        encoding = self.tokenizer(
            text,
            add_special_tokens=True,
            max_length=self.max_len,
            padding='max_length',
            truncation=True,
            return_attention_mask=True,
            return_tensors='pt'
        )
        
        return {
            'input_ids': encoding['input_ids'].flatten(),
            'attention_mask': encoding['attention_mask'].flatten(),
            'label': torch.tensor(label, dtype=torch.long)
        }

# ============================================================
# MAIN TRAINING AND EVALUATION PIPELINE
# ============================================================
def plot_confusion_matrix(cm, classes, title, filepath):
    plt.figure(figsize=(6, 5))
    sns.heatmap(cm, annot=True, fmt='d', cmap='Blues', 
                xticklabels=classes, yticklabels=classes)
    plt.title(title, fontsize=12, fontweight='bold', color='#1F497D')
    plt.ylabel('True Label')
    plt.xlabel('Predicted Label')
    plt.tight_layout()
    plt.savefig(filepath, dpi=300)
    plt.close()

def main():
    print("Loading clean dataset...")
    df = pd.read_csv("Assignment 3/code_switched_sentiment_dataset_merged.csv")
    
    # Fill any empty values
    df['text'] = df['text'].fillna("")
    
    # Preprocess text for TF-IDF and BiLSTM
    print("Preprocessing texts...")
    df['clean_text'] = df['text'].apply(preprocess_text)
    df['label_idx'] = df['sentiment'].map(LABEL_MAP)
    
    # Split dataset: 80% train, 10% val, 10% test
    print("Splitting datasets...")
    train_df, test_df = train_test_split(df, test_size=0.2, random_state=42, stratify=df['label_idx'])
    val_df, test_df = train_test_split(test_df, test_size=0.5, random_state=42, stratify=test_df['label_idx'])
    
    print(f"Train size: {len(train_df)}")
    print(f"Validation size: {len(val_df)}")
    print(f"Test size: {len(test_df)}")
    
    # Save the split indexes for reproducible experiments
    os.makedirs("Assignment 3/models", exist_ok=True)
    
    results = {}
    
    # ============================================================
    # MODEL 1: TF-IDF + LOGISTIC REGRESSION
    # ============================================================
    print("\n" + "="*50)
    print("TRAINING MODEL 1: TF-IDF + LOGISTIC REGRESSION (Baseline)")
    print("="*50)
    
    vectorizer = TfidfVectorizer(max_features=5000, ngram_range=(1, 2))
    X_train_tfidf = vectorizer.fit_transform(train_df['clean_text'])
    X_test_tfidf = vectorizer.transform(test_df['clean_text'])
    
    lr_model = LogisticRegression(C=1.0, max_iter=300, class_weight='balanced')
    lr_model.fit(X_train_tfidf, train_df['label_idx'])
    
    lr_preds = lr_model.predict(X_test_tfidf)
    lr_acc = accuracy_score(test_df['label_idx'], lr_preds)
    lr_p, lr_r, lr_f1, _ = precision_recall_fscore_support(test_df['label_idx'], lr_preds, average='macro')
    
    print(f"Logistic Regression Accuracy: {lr_acc:.4f}")
    print(f"Logistic Regression Macro F1: {lr_f1:.4f}")
    print("\nClassification Report:")
    print(classification_report(test_df['label_idx'], lr_preds, target_names=['negative', 'neutral', 'positive']))
    
    lr_cm = confusion_matrix(test_df['label_idx'], lr_preds)
    plot_confusion_matrix(lr_cm, ['negative', 'neutral', 'positive'], 
                          'Logistic Regression Confusion Matrix', 
                          'Assignment 3/visualizations/confusion_matrix_lr.png')
    
    results['Logistic Regression'] = {'accuracy': lr_acc, 'precision': lr_p, 'recall': lr_r, 'f1': lr_f1}
    
    # ============================================================
    # MODEL 2: PyTorch BiLSTM
    # ============================================================
    print("\n" + "="*50)
    print("TRAINING MODEL 2: PyTorch BiLSTM (Sequence Model)")
    print("="*50)
    
    # Build vocabulary
    vocab = Vocabulary(max_size=10000, min_freq=2)
    vocab.build_vocab(train_df['clean_text'])
    print(f"Vocab size (including PAD and UNK): {vocab.vocab_size}")
    
    train_dataset = TextDataset(train_df['clean_text'].tolist(), train_df['label_idx'].tolist(), vocab)
    val_dataset = TextDataset(val_df['clean_text'].tolist(), val_df['label_idx'].tolist(), vocab)
    test_dataset = TextDataset(test_df['clean_text'].tolist(), test_df['label_idx'].tolist(), vocab)
    
    train_loader = DataLoader(train_dataset, batch_size=64, shuffle=True)
    val_loader = DataLoader(val_dataset, batch_size=64, shuffle=False)
    test_loader = DataLoader(test_dataset, batch_size=64, shuffle=False)
    
    bilstm_model = BiLSTMClassifier(vocab_size=vocab.vocab_size, embedding_dim=100, hidden_dim=128, output_dim=3)
    bimodal = bilstm_model.to(device)
    
    criterion = nn.CrossEntropyLoss()
    optimizer = optim.Adam(bimodal.parameters(), lr=0.001, weight_decay=1e-5)
    
    epochs = 3
    print(f"Training BiLSTM for {epochs} epochs...")
    for epoch in range(epochs):
        bimodal.train()
        epoch_loss = 0
        correct = 0
        total = 0
        for inputs, targets in train_loader:
            inputs, targets = inputs.to(device), targets.to(device)
            optimizer.zero_grad()
            outputs = bimodal(inputs)
            loss = criterion(outputs, targets)
            loss.backward()
            optimizer.step()
            
            epoch_loss += loss.item()
            _, predicted = outputs.max(1)
            total += targets.size(0)
            correct += predicted.eq(targets).sum().item()
            
        train_acc = correct / total
        
        # Validation
        bimodal.eval()
        val_loss = 0
        val_correct = 0
        val_total = 0
        with torch.no_grad():
            for inputs, targets in val_loader:
                inputs, targets = inputs.to(device), targets.to(device)
                outputs = bimodal(inputs)
                loss = criterion(outputs, targets)
                val_loss += loss.item()
                _, predicted = outputs.max(1)
                val_total += targets.size(0)
                val_correct += predicted.eq(targets).sum().item()
        
        val_acc = val_correct / val_total
        print(f"Epoch {epoch+1}/{epochs} - Loss: {epoch_loss/len(train_loader):.4f} - Train Acc: {train_acc:.4f} - Val Loss: {val_loss/len(val_loader):.4f} - Val Acc: {val_acc:.4f}")
        
    # Evaluate BiLSTM on Test Set
    bimodal.eval()
    bilstm_preds = []
    with torch.no_grad():
        for inputs, _ in test_loader:
            inputs = inputs.to(device)
            outputs = bimodal(inputs)
            _, predicted = outputs.max(1)
            bilstm_preds.extend(predicted.cpu().numpy())
            
    bilstm_acc = accuracy_score(test_df['label_idx'], bilstm_preds)
    bilstm_p, bilstm_r, bilstm_f1, _ = precision_recall_fscore_support(test_df['label_idx'], bilstm_preds, average='macro')
    
    print(f"BiLSTM Accuracy: {bilstm_acc:.4f}")
    print(f"BiLSTM Macro F1: {bilstm_f1:.4f}")
    print("\nClassification Report:")
    print(classification_report(test_df['label_idx'], bilstm_preds, target_names=['negative', 'neutral', 'positive']))
    
    bilstm_cm = confusion_matrix(test_df['label_idx'], bilstm_preds)
    plot_confusion_matrix(bilstm_cm, ['negative', 'neutral', 'positive'], 
                          'BiLSTM Confusion Matrix', 
                          'Assignment 3/visualizations/confusion_matrix_bilstm.png')
    
    results['BiLSTM'] = {'accuracy': bilstm_acc, 'precision': bilstm_p, 'recall': bilstm_r, 'f1': bilstm_f1}
    
    # Save checkpoint
    torch.save({
        'model_state_dict': bimodal.state_dict(),
        'vocab_word2idx': vocab.word2idx,
        'vocab_size': vocab.vocab_size
    }, 'Assignment 3/models/bilstm_checkpoint.pt')
    print("BiLSTM checkpoint saved to Assignment 3/models/bilstm_checkpoint.pt")
    
    # ============================================================
    # MODEL 3: HuggingFace mBERT Fine-tuning
    # ============================================================
    print("\n" + "="*50)
    print("TRAINING MODEL 3: Fine-tuned mBERT (Transformer-based)")
    print("="*50)
    
    print("Initializing mBERT Tokenizer...")
    tokenizer = BertTokenizer.from_pretrained('bert-base-multilingual-cased')
    
    # To run successfully on CPU while fine-tuning completely on the full dataset,
    # we use the full training, validation, and test datasets, and optimize CPU 
    # speed by setting max_len=24 (covering 95%+ of typical social media comments).
    print("Initializing datasets on the full data splits (max_len=24)...")
    train_bert_ds = BERTDataset(train_df['text'].tolist(), train_df['label_idx'].tolist(), tokenizer, max_len=24)
    val_bert_ds = BERTDataset(val_df['text'].tolist(), val_df['label_idx'].tolist(), tokenizer, max_len=24)
    test_bert_ds = BERTDataset(test_df['text'].tolist(), test_df['label_idx'].tolist(), tokenizer, max_len=24)
    
    train_bert_loader = DataLoader(train_bert_ds, batch_size=16, shuffle=True)
    val_bert_loader = DataLoader(val_bert_ds, batch_size=16, shuffle=False)
    test_bert_loader = DataLoader(test_bert_ds, batch_size=16, shuffle=False)
    
    print("Loading pretrained bert-base-multilingual-cased...")
    bert_model = BertForSequenceClassification.from_pretrained('bert-base-multilingual-cased', num_labels=3)
    bert_model = bert_model.to(device)
    
    optimizer = optim.AdamW(bert_model.parameters(), lr=2e-5)
    
    bert_epochs = 3
    print(f"Training mBERT on the full training set ({len(train_bert_ds)} samples) for {bert_epochs} epochs...")
    for epoch in range(bert_epochs):
        bert_model.train()
        epoch_loss = 0
        correct = 0
        total = 0
        step = 0
        for batch in train_bert_loader:
            optimizer.zero_grad()
            input_ids = batch['input_ids'].to(device)
            attention_mask = batch['attention_mask'].to(device)
            labels = batch['label'].to(device)
            
            outputs = bert_model(input_ids=input_ids, attention_mask=attention_mask, labels=labels)
            loss = outputs.loss
            loss.backward()
            optimizer.step()
            
            epoch_loss += loss.item()
            logits = outputs.logits
            _, preds = torch.max(logits, dim=1)
            correct += torch.sum(preds == labels).item()
            total += labels.size(0)
            step += 1
            if step % 100 == 0:
                print(f"  Epoch {epoch+1} - Step {step}/{len(train_bert_loader)} - Batch Loss: {loss.item():.4f}")
                
        train_acc = correct / total
        
        # Validation
        bert_model.eval()
        val_loss = 0
        val_correct = 0
        val_total = 0
        with torch.no_grad():
            for batch in val_bert_loader:
                input_ids = batch['input_ids'].to(device)
                attention_mask = batch['attention_mask'].to(device)
                labels = batch['label'].to(device)
                
                outputs = bert_model(input_ids=input_ids, attention_mask=attention_mask, labels=labels)
                val_loss += outputs.loss.item()
                logits = outputs.logits
                _, preds = torch.max(logits, dim=1)
                val_total += labels.size(0)
                val_correct += torch.sum(preds == labels).item()
                
        val_acc = val_correct / val_total
        print(f"mBERT Epoch {epoch+1}/{bert_epochs} - Loss: {epoch_loss/len(train_bert_loader):.4f} - Train Acc: {train_acc:.4f} - Val Loss: {val_loss/len(val_bert_loader):.4f} - Val Acc: {val_acc:.4f}")
        
    # Evaluate mBERT on Test set
    bert_model.eval()
    bert_preds = []
    with torch.no_grad():
        for batch in test_bert_loader:
            input_ids = batch['input_ids'].to(device)
            attention_mask = batch['attention_mask'].to(device)
            outputs = bert_model(input_ids=input_ids, attention_mask=attention_mask)
            logits = outputs.logits
            _, preds = torch.max(logits, dim=1)
            bert_preds.extend(preds.cpu().numpy())
            
    bert_acc = accuracy_score(test_df['label_idx'], bert_preds)
    bert_p, bert_r, bert_f1, _ = precision_recall_fscore_support(test_df['label_idx'], bert_preds, average='macro')
    
    print(f"mBERT Accuracy (Full Test Set): {bert_acc:.4f}")
    print(f"mBERT Macro F1 (Full Test Set): {bert_f1:.4f}")
    print("\nClassification Report:")
    print(classification_report(test_df['label_idx'], bert_preds, target_names=['negative', 'neutral', 'positive']))
    
    bert_cm = confusion_matrix(test_df['label_idx'], bert_preds)
    plot_confusion_matrix(bert_cm, ['negative', 'neutral', 'positive'], 
                          'mBERT Confusion Matrix', 
                          'Assignment 3/visualizations/confusion_matrix_mbert.png')
    
    results['mBERT'] = {'accuracy': bert_acc, 'precision': bert_p, 'recall': bert_r, 'f1': bert_f1}
    
    # Save bert model config/weights
    bert_model.save_pretrained('Assignment 3/models/mbert_fine_tuned')
    tokenizer.save_pretrained('Assignment 3/models/mbert_fine_tuned')
    print("mBERT fine-tuned model saved to Assignment 3/models/mbert_fine_tuned")
    
    # ============================================================
    # SAVE EXPERIMENTAL COMPARISON RESULTS
    # ============================================================
    results_df = pd.DataFrame(results).T
    print("\n" + "="*50)
    print("FINAL EXPERIMENTAL COMPARISON")
    print("="*50)
    print(results_df)
    results_df.to_csv("Assignment 3/experimental_results.csv")
    print("Results saved to Assignment 3/experimental_results.csv")
    
    # Write a clean text summary of evaluation metrics for report parsing
    with open("Assignment 3/results_summary.txt", "w") as f:
        f.write("Model,Accuracy,Precision,Recall,F1\n")
        for model_name, metrics in results.items():
            f.write(f"{model_name},{metrics['accuracy']:.4f},{metrics['precision']:.4f},{metrics['recall']:.4f},{metrics['f1']:.4f}\n")

if __name__ == "__main__":
    main()
