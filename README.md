
# Word Document Search and Replace Tool

This is a tool that helps you search for certain words in Microsoft Word documents (with either .doc or .docx format) and replace them with another word that you specify.

## How to Use

1.  Download the code from GitHub and save it to your computer.
2.  Open the code in a text editor (like Notepad or Sublime Text).
3.  Find the line that says `directory = '/path/to/directory/containing/word/documents/'`. Replace this with the path to the folder that contains the Microsoft Word documents you want to search and replace words in.
4.  Find the line that says `search_replace_dict = {'cat': 'dog', 'car': 'bike', 'breakfast': 'dinner'}`. Replace this with the words you want to search and replace in the Microsoft Word documents. The format should be `{'word_to_search': 'word_to_replace_with', ...}`.
5.  Save the changes you made to the code.
6.  Open a terminal/command prompt window and navigate to the folder where you saved the code.
7.  Type `python filename.py` (replace "filename" with the name of the code file you saved) and hit enter.
8.  Wait for the tool to run and search for the words you specified. It will then replace those words in the Microsoft Word documents.

## What the Code Does

The code has two parts: the `search_and_replace` function and the `if __name__ == '__main__'` section.

The `search_and_replace` function takes two arguments: the file path of a Microsoft Word document and a dictionary that tells the function which words to search for and what to replace them with. It then opens the document using the `Document` module from the `docx` library and loops through every paragraph and every run in each paragraph. For each run, it checks if any of the words in the search/replace dictionary are in the run text. If they are, it replaces them. Finally, it saves the changes to the document.

The `if __name__ == '__main__'` section is a special section of code that only runs when the code is executed as the main program (not when it's imported as a module into another program). In this section, it sets the directory to the folder containing the Microsoft Word documents, sets the search/replace dictionary, and then loops through every file in the directory. If the file is a Microsoft Word document (.doc or .docx), it calls the `search_and_replace` function with the file path and the search/replace dictionary.

## Notes

-   Make sure the path to the folder with the Microsoft Word documents is correct and that the folder contains only the documents you want to search and replace in.
-   Make sure the search/replace dictionary is formatted correctly (as a dictionary with keys being the words to search for and values being the words to replace with).
-   This code is tested to work with Python 3.x. It may not work with earlier versions of Python.
-   This code uses the `docx` library to read and write Microsoft Word documents. You will need to install this library if you don't have it already. You can install it by running `pip install python-docx` in a terminal/command prompt window.
