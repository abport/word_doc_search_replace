# Microsoft Word Document Word Replace

This is a simple Python script that searches for certain words (strings) in Microsoft Word documents (doc and docx format) and replaces them with certain words (strings) specified in an Excel file. The script uses the `pandas` library to read the words to find and replace from an Excel file, and the `docx` library to manipulate the Microsoft Word documents.

## How to use

1.  Install the required libraries: `pandas`, `docx`, and `re` using `pip`:

```python
pip install pandas
pip install python-docx
pip install re
```

2.  Save the script to your computer and modify the following two lines to specify the path to the directory containing your Microsoft Word documents and the path to the Excel file containing the words to find and replace:

```python
directory = '/path/to/directory/containing/word/documents/'
excel_file = '/path/to/excel/file/containing/words/to/find/and/replace/'
```

3.  Make sure the Excel file contains two columns named `find` and `replace` that contain the words to find and the words to replace respectively.
4.  Run the script using Python:

```python
python word_replace.py
```

The script will loop through all the Microsoft Word documents in the specified directory, find and replace the specified words, and save the updated documents.

## License

This code is released under the MIT License.
