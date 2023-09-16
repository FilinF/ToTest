import os
import re
import docx
import requests
from bs4 import BeautifulSoup
from transliterate import translit
import win32com.client as win32
from win32com.client import constants
from glob import glob


def save_as_docx(path):
    # Opening MS Word
    word = win32.gencache.EnsureDispatch('Word.Application')
    doc = word.Documents.Open(path)
    doc.Activate ()

    # Rename path with .docx
    new_file_abs = os.path.abspath(path)
    new_file_abs = re.sub(r'\.\w+$', '.docx', new_file_abs)

    # Save and Close
    word.ActiveDocument.SaveAs(
        new_file_abs, FileFormat=constants.wdFormatXMLDocument
    )
    doc.Close(False)
def transliterate_file(filename):
    transliterated_name = translit(filename, 'ru', reversed=True)
    os.rename(filename, transliterated_name)


# Create a directory called "downloaded_files" if it doesn't exist already
if not os.path.exists('downloaded_files'):
    os.makedirs('downloaded_files')


def save_from_rsreu(link):
    filename = os.path.join('downloaded_files', link.split('/')[-1])
    print(filename)
    r = requests.get(link, allow_redirects=True)
    open(filename, "wb").write(r.content)

    return filename

def URLDocx():
    with open("index.html", encoding="utf-8") as file:
        src = file.read()

    soup = BeautifulSoup(src, "html.parser")

    sociallinks = soup.find_all('a', href=re.compile("-kurs"), )

    return sociallinks

sosual_link = URLDocx()
for item in sosual_link:
    item_text = item.text
    item_url = item.get("href")
    url = "http://www.rsreu.ru/" + item_url + ".doc"

# Save the file to the "downloaded_files" directory
    filename = save_from_rsreu(url)

# Transliterate the filename and move it to the "downloaded_files" directory
    fNewname = os.path.join('downloaded_files', translit(item_text, 'ru', reversed=True) + '.doc')
    if os.path.exists(fNewname):
        os.remove(fNewname)
    os.rename(filename, fNewname)

# Remove unwanted files from the "downloaded_files" directory
    unwanted_filenames = [" 1 kurs.doc", " 2 kurs.doc", " 3 kurs.doc", " 4 kurs.doc", " 5 kurs.doc", " 6 kurs.doc", ".doc", "Dovuzovskaja podgotovka.doc" ]

# Remove unwanted files from the "downloaded_files" directory
    for filename in unwanted_filenames:
        unwanted_filename = os.path.join('downloaded_files', filename)
        if os.path.exists(unwanted_filename):
            os.remove(unwanted_filename)

paths = glob('\\*.doc', recursive=True)
for path in paths:
    save_as_docx(path)
    os.remove(path)
