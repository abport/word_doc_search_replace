# Importing necessary libraries
import os
import re
import pandas as pd
from docx import Document

# Function to search and replace words in a Microsoft Word document


def search_and_replace(file_path, search_replace_dict):
    # Open the Microsoft Word document
    document = Document(file_path)

    # Loop through all the paragraphs in the document
    for paragraph in document.paragraphs:

        # Loop through all the runs (text segments) in the paragraph
        for run in paragraph.runs:

            # Loop through the words to find and replace
            for search_word, replace_word in search_replace_dict.items():

                # Check if the search word is in the current run of text
                if re.search(search_word, run.text):

                    # If the search word is found, replace it with the replace word
                    run.text = re.sub(search_word, replace_word, run.text)

    # Save the changes to the Microsoft Word document
    document.save(file_path)


# Main function that runs the program
if __name__ == '__main__':
    # The path to the directory containing the Microsoft Word documents
    directory = '/path/to/directory/containing/word/documents/'

    # The path to the excel file containing the words to find and replace
    excel_file = '/path/to/excel/file/containing/words/to/find/and/replace/'

    # Read the excel file into a pandas DataFrame
    df = pd.read_excel(excel_file)

    # Create a dictionary of the words to find and replace
    search_replace_dict = dict(zip(df['find'], df['replace']))

    # Loop through all the files in the directory
    for filename in os.listdir(directory):

        # Check if the file is a Microsoft Word document (docx or doc)
        if filename.endswith('.docx') or filename.endswith('.doc'):

            # Get the full path to the file
            file_path = os.path.join(directory, filename)

            # Call the search_and_replace function for the current file
            search_and_replace(file_path, search_replace_dict)
