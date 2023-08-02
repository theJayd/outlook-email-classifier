import string
import re
import nltk
from nltk.corpus import stopwords
from nltk.stem import PorterStemmer
from nltk.stem import WordNetLemmatizer
import email
from bs4 import BeautifulSoup
import joblib
from sklearn.base import BaseEstimator, TransformerMixin
from flask import Flask, request, jsonify
nltk.download('stopwords')
nltk.download('wordnet')

stemmer = PorterStemmer()
lemmatizer = WordNetLemmatizer()

class email_to_clean_text(BaseEstimator, TransformerMixin):
    def __init__(self):
        pass

    def fit(self, X, y=None):
        return self

    def transform(self, X):
        text_list = []
        for mail in X:
            b = email.message_from_string(mail)
            body = ""

            if b.is_multipart():
                for part in b.walk():
                    ctype = part.get_content_type()
                    cdispo = str(part.get('Content-Disposition'))

                    # skip any text/plain (txt) attachments
                    if ctype == 'text/plain' and 'attachment' not in cdispo:
                        body = part.get_payload(decode=True)  # get body of email
                        break
            # not multipart - i.e. plain text, no attachments, keeping fingers crossed
            else:
                body = b.get_payload(decode=True)  # get body of email

            soup = BeautifulSoup(body, "html.parser")  # get text from body (HTML/text)
            text = soup.get_text().lower()

            text = re.sub(r'(https|http)?:\/\/(\w|\.|\/|\?|\=|\&|\%)*\b', '', text, flags=re.MULTILINE)  # remove links

            text = re.sub(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b', '', text, flags=re.MULTILINE)  # remove email addresses

            text = text.translate(str.maketrans('', '', string.punctuation))  # remove punctuation

            text = ''.join([i for i in text if not i.isdigit()])  # remove digits

            stop_words = stopwords.words('english')
            words_list = [w for w in text.split() if w not in stop_words]  # remove stop words

            words_list = [lemmatizer.lemmatize(w) for w in words_list]  # lemmatization

            words_list = [stemmer.stem(w) for w in words_list]  # Stemming
            text_list.append(' '.join(words_list))
        return text_list

# Load the trained email classification pipeline
path_to_load = r'C:\Users\almos\Desktop\OutlookEmailClassifier\Email_Classification\my_pipeline.pkl'
my_loaded_pipeline = joblib.load(path_to_load)
    
def preprocess_and_classify_email(email_text):
    # Create an instance of the email_to_clean_text class
    email_cleaner = email_to_clean_text()

    # Convert email text to a list using the transform method of the email_cleaner instance
    text_list = email_cleaner.transform([email_text])

    # Now, use the loaded pipeline for classification
    prediction = my_loaded_pipeline.predict(text_list)

    return prediction



app = Flask(__name__)

# Define the route for the /classify_email endpoint with POST method
@app.route('/classify_email', methods=['POST'])
def classify_email():
    # Get the JSON data from the request
    data = request.get_json()

    # Extract the email_text from the JSON data
    email_text = data['email_text']

    # Preprocess and classify the email_text
    prediction = preprocess_and_classify_email(email_text)
    
    if(int(prediction[0]==0)):
        result="not a phishing email"
    else:
        result="Most likely a phishing email"
    # Return the prediction as JSON response
    return jsonify({'prediction': result})

if __name__ == '__main__':
    # Run the app
    app.run(debug=False)
