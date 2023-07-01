import openai
import subprocess
import os
from docx import Document

# Set up your OpenAI API credentials
openai.api_key = 'key'

# Check if the required packages are installed
required_packages = ['python-docx']
missing_packages = []
for package in required_packages:
    try:
        __import__(package)
    except ImportError:
        missing_packages.append(package)

# Install missing packages using pip
if missing_packages:
    subprocess.check_call(['pip3', 'install'] + missing_packages)

# Get the current directory
current_directory = os.path.dirname(os.path.abspath(__file__))

# Construct the relative file path to the Word document
document_path = os.path.join(current_directory, 'W5_Hands-on_ProjectSubmission.docx')

# Load the Word document
document = Document(document_path)

# Extract the text from the document
text = ''
for paragraph in document.paragraphs:
    text += paragraph.text + '\n'

# Set up the OpenAI ChatGPT API parameters
model = 'gpt-3.5-turbo'
context = text  # Use the entire document as the context

# Define a function to interact with the ChatGPT API
def chat_with_gpt(message):
    response = openai.Completion.create(
        engine=model,
        context=[context + message],  # Pass context as a list of strings
        max_tokens=50,
        temperature=0.7,
        n=1,
        stop=None,
    )
    return response.choices[0].text.strip()

# Main loop for user interaction
while True:
    user_input = input("Enter your question or type 'exit' to quit: ")
    if user_input.lower() == 'exit':
        break
    response = chat_with_gpt(user_input)
    print(response)
