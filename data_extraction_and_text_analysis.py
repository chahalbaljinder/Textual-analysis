# Import necessary packages
import requests
from bs4 import BeautifulSoup
import pandas as pd
import os
import nltk
from nltk.tokenize import word_tokenize
from nltk.corpus import stopwords
import re

# Download necessary NLTK resources
nltk.download('punkt')
nltk.download('stopwords')

# Read the URL file into a pandas DataFrame
df = pd.read_excel(r'C:\Users\admin\Desktop\github repos\Textual-analysis\Input.xlsx')

# Create lists to store processed data
positive_words = []
negative_words = []
positive_score = []
negative_score = []
polarity_score = []
subjectivity_score = []
avg_sentence_length = []
percentage_of_complex_words = []
fog_index = []
complex_word_count = []
avg_syllable_word_count = []
word_count = []
average_word_length = []
pp_count = []

# Directories
text_dir = r"C:\Users\admin\Desktop\github repos\Textual-analysis\TitleText.txt"
stopwords_dir = r"C:\Users\admin\Desktop\github repos\Textual-analysis\StopWords"
sentiment_dir = r"C:\Users\admin\Desktop\github repos\Textual-analysis\MasterDictionary"

# Load stopwords
stop_words = set()
for file in os.listdir(stopwords_dir):
    with open(os.path.join(stopwords_dir, file), 'r', encoding='ISO-8859-1') as f:
        stop_words.update(set(f.read().splitlines()))

# Load positive and negative words
pos = set()
neg = set()
for file in os.listdir(sentiment_dir):
    with open(os.path.join(sentiment_dir, file), 'r', encoding='ISO-8859-1') as f:
        if 'positive-words.txt' in file:
            pos.update(f.read().splitlines())
        else:
            neg.update(f.read().splitlines())

# Define functions for text analysis
def measure(file):
    with open(os.path.join(text_dir, file), 'r') as f:
        text = f.read()
        text = re.sub(r'[^\w\s]', '', text)
        sentences = text.split('.')
        num_sentences = len(sentences)
        words = [word for word in text.split() if word.lower() not in stop_words]
        num_words = len(words)
        complex_words = [word for word in words if sum(1 for letter in word if letter.lower() in 'aeiou') > 2]
        syllable_count = sum(
            sum(1 for letter in word if letter.lower() in 'aeiou') for word in words
        )
        avg_sentence_len = num_words / num_sentences if num_sentences > 0 else 0
        avg_syllable_word_count = syllable_count / len(words) if words else 0
        percent_complex_words = len(complex_words) / num_words if num_words > 0 else 0
        fog_index = 0.4 * (avg_sentence_len + percent_complex_words)
        return avg_sentence_len, percent_complex_words, fog_index, len(complex_words), avg_syllable_word_count

def cleaned_words(file):
    with open(os.path.join(text_dir, file), 'r') as f:
        text = f.read()
        text = re.sub(r'[^\w\s]', '', text)
        words = [word for word in text.split() if word.lower() not in stop_words]
        length = sum(len(word) for word in words)
        average_word_length = length / len(words) if words else 0
    return len(words), average_word_length

def count_personal_pronouns(file):
    with open(os.path.join(text_dir, file), 'r') as f:
        text = f.read()
        personal_pronouns = ["I", "we", "my", "ours", "us"]
        count = sum(len(re.findall(r"\b" + pronoun + r"\b", text)) for pronoun in personal_pronouns)
    return count

# Process each file
for index, row in df.iterrows():
    url = row['URL']
    url_id = row['URL_ID']
    header = {'User-Agent': "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36"}
    
    try:
        response = requests.get(url, headers=header)
        soup = BeautifulSoup(response.content, 'html.parser')
        title = soup.find('h1').get_text()
        article = ' '.join(p.get_text() for p in soup.find_all('p'))
        file_name = os.path.join(text_dir, f'TitleText{url_id}.txt')
        with open(file_name, 'w') as file:
            file.write(title + '\n' + article)
    except Exception as e:
        print(f"Error processing URL_ID {url_id}: {e}")
        continue

# Calculate sentiment and textual features
for file in os.listdir(text_dir):
    avg_len, pct_complex, fog, comp_word_count, avg_syllables = measure(file)
    wc, avg_word_len = cleaned_words(file)
    pp_count.append(count_personal_pronouns(file))
    
    avg_sentence_length.append(avg_len)
    percentage_of_complex_words.append(pct_complex)
    fog_index.append(fog)
    complex_word_count.append(comp_word_count)
    avg_syllable_word_count.append(avg_syllables)
    word_count.append(wc)
    average_word_length.append(avg_word_len)

# Read existing output DataFrame and drop invalid rows
output_df = pd.read_excel('Output Data Structure.xlsx')
output_df.drop([44-37, 57-37, 144-37], axis=0, inplace=True)

# Prepare data for output DataFrame
variables = [positive_score, negative_score, polarity_score, subjectivity_score, avg_sentence_length,
             percentage_of_complex_words, fog_index, complex_word_count, word_count,
             avg_syllable_word_count, pp_count, average_word_length]

for i, var in enumerate(variables):
    output_df.iloc[:, i+2] = var

# Save the updated DataFrame to a CSV file
output_df.to_csv('Output_Data.csv', index=False)
