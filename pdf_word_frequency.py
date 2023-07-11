import os
from PyPDF2 import PdfReader
import pandas as pd
import nltk
from nltk.corpus import stopwords
from nltk.probability import FreqDist
from nltk.tokenize import word_tokenize
from collections import Counter
from tqdm import tqdm

nltk.download('punkt')
nltk.download('stopwords')

# Clear the screen
os.system('cls' if os.name == 'nt' else 'clear')

# Display the title
print('''
 __          __           _   ______                                          
 \ \        / /          | | |  ____|                                         
  \ \  /\  / /__  _ __ __| | | |__ _ __ ___  __ _ _   _  ___ _ __   ___ _   _ 
   \ \/  \/ / _ \| '__/ _` | |  __| '__/ _ \/ _` | | | |/ _ \ '_ \ / __| | | |
    \  /\  / (_) | | | (_| | | |  | | |  __/ (_| | |_| |  __/ | | | (__| |_| |
     \/  \/ \___/|_|  \__,_| |_|  |_|  \___|\__, |\__,_|\___|_| |_|\___|\__, |
                                               | |                       __/ |
                                               |_|                      |___/                                                                                  
''')

# Display the instructions
print('''
This script will count the frequencies of words in a selected PDF file,
excluding common stopwords. The word frequencies will be saved to an Excel file.

Follow the prompts to select a PDF file to process.
''')

# Get a list of PDF files in the local folder
pdf_files = [file for file in os.listdir() if file.endswith('.pdf')]

# Display numbered list of PDF files
print("PDF Files:")
for i, file in enumerate(pdf_files):
    print(f"{i+1}. {file}")

# Get user input for selecting a file
selection = input("Enter the number of the PDF file to process: ")

# Verify the input is a valid number within the range
while not selection.isdigit() or int(selection) < 1 or int(selection) > len(pdf_files):
    print("Invalid selection. Please enter a valid number.")
    selection = input("Enter the number of the PDF file to process: ")

# Get the selected PDF file based on the user's input
selected_file = pdf_files[int(selection) - 1]

# Read PDF
with open(selected_file, 'rb') as file:
    reader = PdfReader(file)
    text = ''
    for page_num in tqdm(range(len(reader.pages)), desc="Reading PDF"):
        text += reader.pages[page_num].extract_text()

print("\nTokenizing and filtering text...")
# Tokenize text (split into words)
tokens = word_tokenize(text)

# Convert to lower case
tokens = [word.lower() for word in tqdm(tokens, desc="Converting to lower case")]

# Filter out punctuation and stopwords
tokens = [word for word in tqdm(tokens, desc="Filtering out punctuation") if word.isalpha()]
tokens = [word for word in tqdm(tokens, desc="Filtering out stopwords") if word not in stopwords.words('english')]

# Count word frequencies
freq = Counter(tokens)

# Filter words with frequency > 5
filtered_words = {key: val for key, val in freq.items() if val > 5}

# Convert to DataFrame
df = pd.DataFrame.from_dict(filtered_words, orient='index', columns=['Word Count'])
df.index.name = 'Word'

# Get the current working directory
current_dir = os.getcwd()
print(f"\nCurrent working directory: {current_dir}")

# Create the output file path
output_filename = os.path.splitext(selected_file)[0] + '.xlsx'
output_path = os.path.join(current_dir, output_filename)

# Save DataFrame to Excel
df.to_excel(output_path, index=True)

# Print the name of the output file
print(f"\nOutput file created: {output_path}")

# Exit
selection = input("Press enter to Exit")