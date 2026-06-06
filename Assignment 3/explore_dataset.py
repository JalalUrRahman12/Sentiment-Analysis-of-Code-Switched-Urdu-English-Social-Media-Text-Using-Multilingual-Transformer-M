import os
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import re
import unicodedata
import emoji
from collections import Counter
import nltk
from nltk.tokenize import word_tokenize
from nltk.corpus import stopwords

# Ensure NLTK resources are available
nltk.download('stopwords', quiet=True)
nltk.download('punkt', quiet=True)
nltk.download('punkt_tab', quiet=True)

# Custom Urdu Stopword List
URDU_STOPWORDS = {
    'he', 'hein', 'ka', 'ki', 'ke', 'mein', 'ne', 'se', 'par',
    'aur', 'yeh', 'woh', 'ek', 'bhi', 'to', 'jo', 'ho',
    'is', 'ab', 'tha', 'thi', 'the', 'hum', 'aap', 'mujhe',
    '\u06c1\u06d2', '\u06c1\u06cc\u06ba', '\u06a9\u0627', '\u06a9\u06cc',
    '\u06a9\u06d2', '\u0645\u06cc\u06ba', '\u0646\u06d2', '\u0633\u06d2',
    '\u067e\u0631', '\u0627\u0648\u0631', '\u06cc\u06c1', '\u0648\u06c1',
    '\u0627\u06cc\u06a9', '\u0628\u06be\u06cc', '\u062a\u0648',
    '\u06a9\u06c1', '\u062c\u0648', '\u06c1\u0648', '\u0627\u0633',
    '\u0627\u0628', '\u062a\u06be\u0627', '\u062a\u06be\u06cc',
    '\u062a\u06be\u06d2', '\u06c1\u0645', '\u0622\u067e',
    '\u0645\u062c\u06be\u06d2'
}

ENGLISH_STOPWORDS = set(stopwords.words('english'))

# Preprocessing Pipeline
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

def tokenize(text):
    return word_tokenize(text)

def remove_stopwords(tokens):
    all_stopwords = ENGLISH_STOPWORDS.union(URDU_STOPWORDS)
    return [token for token in tokens if token not in all_stopwords]

def filter_numbers_and_punct(tokens):
    return [t for t in tokens if not t.isdigit() and t.strip()]

def tag_language(tokens):
    tagged = []
    for token in tokens:
        if any('\u0600' <= ch <= '\u06FF' for ch in token):
            tagged.append((token, 'UR'))
        else:
            # We also treat standard Roman Urdu words as English-like/Latin or Roman Urdu.
            # In the SentiMix dataset, standard Roman letters are evaluated.
            # To distinguish, we will assume Latin script tokens.
            # We can classify them as 'UR' (Roman Urdu) or 'EN' (English) using a dictionary 
            # or treat all non-Arabic script Latin tokens as EN/Roman Urdu.
            # For statistics, if a token is in URDU_STOPWORDS or matches Roman Urdu vocab, we tag it UR.
            roman_urdu_keywords = {'bohat', 'achi', 'bakwas', 'ghatia', 'karo', 'kuch', 'hai', 'tha', 'thi', 'par', 'bhi', 'khas', 'log', 'na', 'yaar', 'bhai', 'mujh', 'tujh', 'apna', 'apni'}
            if token in roman_urdu_keywords:
                tagged.append((token, 'UR_ROMAN'))
            else:
                tagged.append((token, 'EN'))
    return tagged

def preprocess_text(text):
    text = normalize_unicode(text)
    text = convert_emojis(text)
    text = remove_urls(text)
    text = clean_mentions_hashtags(text)
    text = lowercase_english(text)
    text = remove_noise(text)
    tokens = tokenize(text)
    tokens = remove_stopwords(tokens)
    tokens = filter_numbers_and_punct(tokens)
    return tokens

def main():
    csv_path = "Assignment 3/code_switched_sentiment_dataset_merged.csv"
    viz_dir = "Assignment 3/visualizations"
    os.makedirs(viz_dir, exist_ok=True)
    
    print("Loading dataset...")
    df = pd.read_csv(csv_path)
    
    print("Total rows:", len(df))
    print(df['sentiment'].value_counts())
    
    # 1. Visualization: Class Distribution
    plt.figure(figsize=(8, 5))
    sns.countplot(x='sentiment', data=df, hue='sentiment', palette='viridis', legend=False)
    plt.title('Sentiment Class Distribution', fontsize=14, fontweight='bold', color='#1F497D')
    plt.xlabel('Sentiment Class', fontsize=12)
    plt.ylabel('Count', fontsize=12)
    plt.grid(axis='y', linestyle='--', alpha=0.7)
    plt.tight_layout()
    plt.savefig(os.path.join(viz_dir, 'class_distribution.png'), dpi=300)
    plt.close()
    print("Generated class_distribution.png")
    
    # Preprocess all texts
    print("Preprocessing texts (this may take a minute)...")
    df['tokens'] = df['text'].apply(preprocess_text)
    df['sentence_length'] = df['tokens'].apply(len)
    
    # 2. Visualization: Sentence Length Histogram
    plt.figure(figsize=(8, 5))
    sns.histplot(df['sentence_length'], bins=30, kde=True, color='#2E74B5', edgecolor='black')
    plt.title('Distribution of Sentence Lengths (after preprocessing)', fontsize=14, fontweight='bold', color='#1F497D')
    plt.xlabel('Number of Tokens', fontsize=12)
    plt.ylabel('Frequency', fontsize=12)
    plt.grid(axis='y', linestyle='--', alpha=0.7)
    plt.tight_layout()
    plt.savefig(os.path.join(viz_dir, 'sentence_length_histogram.png'), dpi=300)
    plt.close()
    print("Generated sentence_length_histogram.png")
    
    # Collect all tokens and tag languages
    print("Analyzing tokens...")
    all_tokens = []
    for tokens in df['tokens']:
        all_tokens.extend(tokens)
        
    tagged_tokens = tag_language(all_tokens)
    
    en_tokens = [t[0] for t in tagged_tokens if t[1] == 'EN']
    ur_tokens = [t[0] for t in tagged_tokens if t[1] in ('UR', 'UR_ROMAN')]
    
    print(f"Total tokens: {len(tagged_tokens)}")
    print(f"English-like tokens: {len(en_tokens)}")
    print(f"Urdu/Roman Urdu tokens: {len(ur_tokens)}")
    
    # 3. Visualization: Language Distribution
    plt.figure(figsize=(6, 6))
    labels = ['English / Transliterated Latin', 'Urdu / Roman Urdu keywords']
    sizes = [len(en_tokens), len(ur_tokens)]
    colors = ['#2E74B5', '#EBF3FB']
    plt.pie(sizes, labels=labels, autopct='%1.1f%%', startangle=140, colors=colors, 
            textprops={'fontsize': 11, 'weight': 'bold'}, wedgeprops={'edgecolor': 'gray'})
    plt.title('Token Language Script Distribution', fontsize=14, fontweight='bold', color='#1F497D')
    plt.tight_layout()
    plt.savefig(os.path.join(viz_dir, 'language_distribution.png'), dpi=300)
    plt.close()
    print("Generated language_distribution.png")
    
    # 4. Visualization: Top 20 English Tokens
    en_counter = Counter(en_tokens)
    top_en = en_counter.most_common(20)
    top_en_df = pd.DataFrame(top_en, columns=['Token', 'Count'])
    
    plt.figure(figsize=(10, 6))
    sns.barplot(y='Token', x='Count', data=top_en_df, hue='Token', palette='Blues_r', legend=False)
    plt.title('Top 20 English-like Tokens', fontsize=14, fontweight='bold', color='#1F497D')
    plt.xlabel('Frequency', fontsize=12)
    plt.ylabel('Token', fontsize=12)
    plt.grid(axis='x', linestyle='--', alpha=0.7)
    plt.tight_layout()
    plt.savefig(os.path.join(viz_dir, 'token_frequency_english.png'), dpi=300)
    plt.close()
    print("Generated token_frequency_english.png")
    
    # 5. Visualization: Top 20 Urdu Tokens
    ur_counter = Counter(ur_tokens)
    top_ur = ur_counter.most_common(20)
    top_ur_df = pd.DataFrame(top_ur, columns=['Token', 'Count'])
    
    plt.figure(figsize=(10, 6))
    sns.barplot(y='Token', x='Count', data=top_ur_df, hue='Token', palette='Greens_r', legend=False)
    plt.title('Top 20 Urdu / Roman Urdu Tokens', fontsize=14, fontweight='bold', color='#1F497D')
    plt.xlabel('Frequency', fontsize=12)
    plt.ylabel('Token', fontsize=12)
    plt.grid(axis='x', linestyle='--', alpha=0.7)
    plt.tight_layout()
    plt.savefig(os.path.join(viz_dir, 'token_frequency_urdu.png'), dpi=300)
    plt.close()
    print("Generated token_frequency_urdu.png")
    
    # Save some key metrics for our report
    with open("Assignment 3/stats_summary.txt", "w", encoding="utf-8") as f:
        f.write(f"Total Rows: {len(df)}\n")
        f.write(f"Neutral: {len(df[df['sentiment']=='neutral'])}\n")
        f.write(f"Positive: {len(df[df['sentiment']=='positive'])}\n")
        f.write(f"Negative: {len(df[df['sentiment']=='negative'])}\n")
        f.write(f"Average Sentence Length: {df['sentence_length'].mean():.2f}\n")
        f.write(f"Max Sentence Length: {df['sentence_length'].max()}\n")
        f.write(f"Total Tokens: {len(tagged_tokens)}\n")
        f.write(f"English Tokens Count: {len(en_tokens)}\n")
        f.write(f"Urdu Tokens Count: {len(ur_tokens)}\n")
    print("Stats written to Assignment 3/stats_summary.txt")

if __name__ == "__main__":
    main()
