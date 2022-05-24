# This is a simple series of scripts to convert word documents to markdown
"""
Sources used include:
https://stackoverflow.com/questions/8024248/telling-python-to-save-a-txt-file-to-a-certain-directory-on-windows-and-mac
https://stackoverflow.com/questions/1347791/unicode-error-unicodeescape-codec-cant-decode-bytes-cannot-open-text-file
https://www.geeksforgeeks.org/how-to-convert-html-to-markdown-in-python/

You will need the following packages installed for everything to run:
https://pypi.org/project/mammoth/
https://pypi.org/project/markdownify/
"""


def convert_html(file):
    import mammoth
    with open(file, "rb") as docx_file:
        result = mammoth.convert_to_html(docx_file).value
        docx_file.close()

    return result


def convert_markdown(html_file):
    import markdownify
    result = markdownify.markdownify(html_file, heading_style="ATX")
    return result


def save_markdown(file, file_name, path):
    import os.path
    completename = os.path.join(path, file_name + ".md")
    file_saved = open(completename, "w")
    file_saved.write(file)
    file_saved.close();


# Update FULL_FILE_PATH, FILE_NAME, and PATH before running
if __name__ == '__main__':
    res_html = convert_html(r"FULL_FILE_PATH")
    print(res_html)
    res_markdown = convert_markdown(res_html)
    print(res_markdown)
    save_markdown(res_markdown, r"FILE_NAME", r"PATH")
