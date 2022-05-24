Download and use the docx_to_md.py script to convert MS Word (.docx) files to Markdown Files (.md).
Run the script in an IDE like Pycharm or VS Code. 

When you run the script you'll need to update the following in the main method:

FULL_FILE_PATH = the full filepath including the .docx ending of the word file you want to convert

FILE_NAME = the name you want to give the output file, do not included paths or endings (i.e., "output" not "./output.md")

PATH = the folder path where you want to save your output file. 

To be safe, use absolute paths. DO NOT REMOVE the r infront of the variable strings (i.e., r"PATH" not "PATH"). 

Sources used include:

https://stackoverflow.com/questions/8024248/telling-python-to-save-a-txt-file-to-a-certain-directory-on-windows-and-mac

https://stackoverflow.com/questions/1347791/unicode-error-unicodeescape-codec-cant-decode-bytes-cannot-open-text-file

https://www.geeksforgeeks.org/how-to-convert-html-to-markdown-in-python/

You will need the following packages installed for everything to run:

https://pypi.org/project/mammoth/

https://pypi.org/project/markdownify/

