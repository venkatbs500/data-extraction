import requests
from bs4 import BeautifulSoup
import pandas as pd
import re
import nltk
from textblob import TextBlob
from textstat import textstat
nltk.download('punkt')
input_data = pd.read_excel("Input.xlsx")


# Define a function to calculate additional variables
def calculate_additional_variables(article_text):
    try:
        words = nltk.word_tokenize(article_text)
        sentences = nltk.sent_tokenize(article_text)

        # AVG SENTENCE LENGTH
        avg_sentence_length = len(words) / len(sentences)

        # PERCENTAGE OF COMPLEX WORDS
        complex_word_count = textstat.difficult_words(article_text)
        percentage_complex_words = (complex_word_count / len(words)) * 100

        # FOG INDEX
        fog_index = textstat.gunning_fog(article_text)

        # AVG NUMBER OF WORDS PER SENTENCE
        avg_words_per_sentence = len(words) / len(sentences)

        # COMPLEX WORD COUNT
        complex_word_count = textstat.difficult_words(article_text)

        # SYLLABLE PER WORD
        syllable_per_word = textstat.syllable_count(article_text) / len(words)

        # PERSONAL PRONOUNS
        personal_pronouns = sum(1 for word in words if
                                word.lower() in ['i', 'me', 'my', 'mine', 'myself', 'we', 'us', 'our', 'ours',
                                                 'ourselves'])

        # AVG WORD LENGTH
        avg_word_length = sum(len(word) for word in words) / len(words)

        # WORD COUNT
        word_count = len(words)

        return {
            'AVG SENTENCE LENGTH': avg_sentence_length,
            'PERCENTAGE OF COMPLEX WORDS': percentage_complex_words,
            'FOG INDEX': fog_index,
            'AVG NUMBER OF WORDS PER SENTENCE': avg_words_per_sentence,
            'COMPLEX WORD COUNT': complex_word_count,
            'SYLLABLE PER WORD': syllable_per_word,
            'PERSONAL PRONOUNS': personal_pronouns,
            'AVG WORD LENGTH': avg_word_length,
            'WORD COUNT': word_count
        }

    except Exception as e:
        print(f"Error calculating additional variables: {str(e)}")
        return None


# Define a function to extract and analyze data for a given URL
def process_url(url, url_id):
    try:
        # Send an HTTP request to the URL
        response = requests.get(url)

        if response.status_code == 200:
            # Parse the HTML content
            soup = BeautifulSoup(response.content, 'html.parser')

            # Extract the article title and text
            article_title = soup.title.string if soup.title else "No Title"
            article_text = ""
            article_text_tag = soup.find("article")  # Adjust this based on the HTML structure

            if article_text_tag:
                # Extract only the text within the article tag
                for paragraph in article_text_tag.find_all('p'):
                    article_text += paragraph.get_text()

            # Save the extracted article text to a file
            with open(f"{url_id}.txt", "w", encoding="utf-8") as file:
                file.write(f"Title: {article_title}\n")
                file.write(article_text)

            # TextBlob for sentiment analysis
            analysis = TextBlob(article_text)

            # Calculate the required variables
            positive_score = analysis.sentiment.polarity
            negative_score = -analysis.sentiment.polarity
            polarity_score = analysis.sentiment.polarity
            subjectivity_score = analysis.sentiment.subjectivity

            # Calculate additional variables
            additional_variables = calculate_additional_variables(article_text)

            if additional_variables:
                result = {
                    'URL_ID': url_id,
                    'POSITIVE SCORE': positive_score,
                    'NEGATIVE SCORE': negative_score,
                    'POLARITY SCORE': polarity_score,
                    'SUBJECTIVITY SCORE': subjectivity_score,
                    **additional_variables
                }

                return result
            else:
                return None

        else:
            print(f"Failed to retrieve data from URL_ID {url_id}")
            return None

    except Exception as e:
        print(f"Error processing URL_ID {url_id}: {str(e)}")
        return None


# Create an empty list to collect the results
output_data = []

# Process each URL and collect the results
for index, row in input_data.iterrows():
    url_id = row['URL_ID']
    url = row['URL']

    result = process_url(url, url_id)

    if result:
        output_data.append(result)

# Create the output DataFrame from the list of results
output_df = pd.DataFrame(output_data)

# Save the output to an Excel file
output_df.to_excel('Output Data Structure.xlsx', index=False)