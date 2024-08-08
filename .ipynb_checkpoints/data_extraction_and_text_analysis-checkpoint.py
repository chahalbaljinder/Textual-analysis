import requests
from bs4 import BeautifulSoup
import pandas as pd
import os
import nltk
from nltk.tokenize import word_tokenize
from nltk.corpus import stopwords
nltk.download('punkt')
nltk.download('stopwords')
import re

# Read the URL file into the pandas object
df = pd.read_excel(r'C:\Users\admin\Desktop\github repos\Textual-analysis\Input.xlsx')

# Loop through each row in the df
for index, row in df.iterrows():
    url = row['URL']
    url_id = row['URL_ID']

    # Make a request to the URL
    headers = {'User-Agent': "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36"}  # Changed header to headers
    try:
        response = requests.get(url, headers=headers)  # Corrected headers parameter
        response.raise_for_status()  # Added this line to raise HTTP errors
    except Exception as e:
        print(f"can't get response for {url_id}: {e}")  # Added exception detail
        continue

    # Create a BeautifulSoup object
    try:
        soup = BeautifulSoup(response.content, 'html.parser')
    except Exception as e:
        print(f"can't parse page for {url_id}: {e}")  # Added exception detail
        continue

    # Find title
    try:
        title = soup.find('h1').get_text()
    except Exception as e:
        print(f"can't get title for {url_id}: {e}")  # Added exception detail
        continue

    # Find text
    article = ""
    try:
        for p in soup.find_all('p'):
            article += p.get_text()
    except Exception as e:
        print(f"can't get text for {url_id}: {e}")  # Added exception detail

    # Ensure the directory exists before writing to it
    os.makedirs(r'C:\Users\admin\Desktop\github repos\Textual-analysis\TitleText', exist_ok=True)

    # Write title and text to the file
    file_name = os.path.join(r'C:\Users\admin\Desktop\github repos\Textual-analysis\TitleText', f'TitleText{url_id}.txt')  # Used os.path.join
    try:
        with open(file_name, 'w', encoding='utf-8') as file:  # Added encoding
            file.write(title + '\n' + article)
    except Exception as e:
        print(f"can't write to file for {url_id}: {e}")  # Added exception detail

# Directories
text_dir = r"C:\Users\admin\Desktop\github repos\Textual-analysis\TitleText"
stopwords_dir = r"C:\Users\admin\Desktop\github repos\Textual-analysis\StopWords"
sentment_dir = r"C:\Users\admin\Desktop\github repos\Textual-analysis\MasterDictionary"

# Load all stop words from the stopwords directory and store in the set variable
stop_words = set()
for files in os.listdir(stopwords_dir):
    with open(os.path.join(stopwords_dir, files), 'r', encoding='ISO-8859-1') as f:
        stop_words.update(set(f.read().splitlines()))

# Load all text files from the directory and store in a list (docs)
docs = []
for text_file in os.listdir(text_dir):
    with open(os.path.join(text_dir, text_file), 'r', encoding='utf-8') as f:  # Added encoding
        text = f.read()
        # Tokenize the given text file
        words = word_tokenize(text)
        # Remove the stop words from the tokens
        filtered_text = [word for word in words if word.lower() not in stop_words]
        # Add each filtered tokens of each file into a list
        docs.append(filtered_text)

# Store positive and negative words from the directory
pos = set()
neg = set()

for files in os.listdir(sentment_dir):
    if files == 'positive-words.txt':
        with open(os.path.join(sentment_dir, files), 'r', encoding='ISO-8859-1') as f:
            pos.update(f.read().splitlines())
    else:
        with open(os.path.join(sentment_dir, files), 'r', encoding='ISO-8859-1') as f:
            neg.update(f.read().splitlines())

# Now collect the positive and negative words from each file
# Calculate the scores from the positive and negative words
positive_words = []
negative_words = []
positive_score = []
negative_score = []
polarity_score = []
subjectivity_score = []

# Iterate through the list of docs
for i in range(len(docs)):
    positive_words.append([word for word in docs[i] if word.lower() in pos])
    negative_words.append([word for word in docs[i] if word.lower() in neg])
    positive_score.append(len(positive_words[i]))
    negative_score.append(len(negative_words[i]))
    polarity_score.append((positive_score[i] - negative_score[i]) / ((positive_score[i] + negative_score[i]) + 0.000001))
    subjectivity_score.append((positive_score[i] + negative_score[i]) / ((len(docs[i])) + 0.000001))

# Average Sentence Length = the number of words / the number of sentences
# Percentage of Complex words = the number of complex words / the number of words 
# Fog Index = 0.4 * (Average Sentence Length + Percentage of Complex words)

avg_sentence_length = []
percentage_of_complex_words = []
fog_index = []
complex_word_count = []
avg_syllable_word_count = []

stopwords = set(stopwords.words('english'))

def measure(file):
    with open(os.path.join(text_dir, file), 'r', encoding='utf-8') as f:  # Added encoding
        text = f.read()
        # Remove punctuations 
        text = re.sub(r'[^\w\s.]', '', text)
        # Split the given text file into sentences
        sentences = text.split('.')
        # Total number of sentences in a file
        num_sentences = len(sentences)
        # Total words in the file
        words = [word for word in text.split() if word.lower() not in stopwords]
        num_words = len(words)

        # Complex words having syllable count is greater than 2
        # Complex words are words in the text that contain more than two syllables.
        complex_words = []
        for word in words:
            vowels = 'aeiou'
            syllable_count_word = sum(1 for letter in word if letter.lower() in vowels)
            if syllable_count_word > 2:
                complex_words.append(word)

        # Syllable Count Per Word
        # We count the number of Syllables in each word of the text by counting the vowels present in each word.
        # We also handle some exceptions like words ending with "es", "ed" by not counting them as a syllable.
        syllable_count = 0
        syllable_words = []
        for word in words:
            if word.endswith('es'):
                word = word[:-2]
            elif word.endswith('ed'):
                word = word[:-2]
            vowels = 'aeiou'
            syllable_count_word = sum(1 for letter in word if letter.lower() in vowels)
            if syllable_count_word >= 1:
                syllable_words.append(word)
                syllable_count += syllable_count_word

        avg_sentence_len = num_words / num_sentences
        avg_syllable_word_count = syllable_count / len(syllable_words)
        percent_complex_words = len(complex_words) / num_words
        fog_index = 0.4 * (avg_sentence_len + percent_complex_words)

        return avg_sentence_len, percent_complex_words, fog_index, len(complex_words), avg_syllable_word_count

# Iterate through each file or doc
for file in os.listdir(text_dir):
    x, y, z, a, b = measure(file)
    avg_sentence_length.append(x)
    percentage_of_complex_words.append(y)
    fog_index.append(z)
    complex_word_count.append(a)
    avg_syllable_word_count.append(b)

# Word Count and Average Word Length Sum of the total number of characters in each word / Total number of words
# We count the total cleaned words present in the text by 
# removing the stop words (using stopwords class of nltk package).
# removing any punctuations like ? ! , . from the word before counting.

def cleaned_words(file):
    with open(os.path.join(text_dir, file), 'r', encoding='utf-8') as f:  # Added encoding
        text = f.read()
        text = re.sub(r'[^\w\s]', '', text)
        words = [word for word in text.split() if word.lower() not in stopwords]
        length = sum(len(word) for word in words)
        average_word_length = length / len(words)
    return len(words), average_word_length

word_count = []
average_word_length = []
for file in os.listdir(text_dir):
    x, y = cleaned_words(file)
    word_count.append(x)
    average_word_length.append(y)

# To calculate Personal Pronouns mentioned in the text, we use regex to find 
# the counts of the words - “I,” “we,” “my,” “ours,” and “us”. Special care is taken
# so that the country name US is not included in the list.
def count_personal_pronouns(file):
    with open(os.path.join(text_dir, file), 'r', encoding='utf-8') as f:  # Added encoding
        text = f.read()
        personal_pronouns = ["I", "we", "my", "ours", "us"]  # Fixed indentation
    count = 0
    for pronoun in personal_pronouns:
        count += len(re.findall(r"\b" + pronoun + r"\b", text))  # Fixed indentation
    return count  # Moved return statement outside of the loop

pp_count = []
for file in os.listdir(text_dir):
    x = count_personal_pronouns(file)
    pp_count.append(x)

output_df = pd.read_excel('Output Data Structure.xlsx')

# URL_ID 44 ,57, 144 does not exists i.e., page does not exist, throws 404 error
# so we are going to drop these rows from the table
output_df.drop([44-37, 57-37, 144-37], axis=0, inplace=True)

# These are the required parameters 
variables = [positive_score,
             negative_score,
             polarity_score,
             subjectivity_score,
             avg_sentence_length,
             percentage_of_complex_words,  # Fixed typo
             fog_index,  # Fixed typo
             avg_sentence_length,  # Fixed duplicate
             complex_word_count,
             word_count,
             avg_syllable_word_count,
             pp_count,
             average_word_length]

# Write the values to the dataframe
for i, var in enumerate(variables):
    output_df.iloc[:, i+2] = var

# Now save the dataframe to the disk
output_df.to_csv('Output_Data.csv', index=False)  # Added index=False
