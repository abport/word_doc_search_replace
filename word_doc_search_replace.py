import os  # This imports the 'os' module which helps with operating system-related functionality
import re  # This imports the 're' module which provides support for regular expressions
# This imports the 'Document' class from the 'docx' module which helps with reading and writing Microsoft Word documents
from docx import Document

# This function takes a file path and a dictionary of search and replace terms and performs the search and replace operation on the Microsoft Word document


def search_and_replace(file_path, search_replace_dict):
    # This opens the Microsoft Word document at the specified file path
    document = Document(file_path)
    for paragraph in document.paragraphs:  # This loop iterates over all the paragraphs in the document
        for run in paragraph.runs:  # This loop iterates over all the text runs within each paragraph
            # This loop iterates over the items in the search and replace dictionary
            for search_word, replace_word in search_replace_dict.items():
                # This checks if the current search word is present in the current text run
                if re.search(search_word, run.text):
                    # If the current search word is present, this line replaces it with the corresponding replace word
                    run.text = re.sub(search_word, replace_word, run.text)
    document.save(file_path)  # This saves the changes made to the document


# This is the main function that is executed when the script is run
if __name__ == '__main__':
    # This is the path to the directory containing the Microsoft Word documents
    directory = '/path/to/directory/containing/word/documents/'
    # This is the dictionary of search and replace terms
    search_replace_dict = {'cat': 'dog', 'car': 'bike', 'breakfast': 'dinner'}
    # This loop iterates over all the files in the directory
    for filename in os.listdir(directory):
        # This line checks if the current file is a Microsoft Word document (either .docx or .doc format)
        if filename.endswith('.docx') or filename.endswith('.doc'):
            # This creates the full file path by joining the directory path and the filename
            file_path = os.path.join(directory, filename)
            # This calls the 'search_and_replace' function on the current Microsoft Word document
            search_and_replace(file_path, search_replace_dict)
