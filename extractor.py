#!/usr/bin/env python
# coding: utf-8

# In[111]:


import pandas as pd
from itertools import zip_longest
import textract
from tika import parser
import os
import re

import io

from pdfminer.pdfinterp import PDFResourceManager
from pdfminer.layout import LAParams
from pdfminer.converter import TextConverter
from pdfminer.pdfinterp import PDFPageInterpreter
from pdfminer.pdfpage import PDFPage

def pdf_to_text(path):
    with open(path, 'rb') as fp:
        rsrcmgr = PDFResourceManager()
        outfp = io.StringIO()
        laparams = LAParams()
        device = TextConverter(rsrcmgr, outfp, laparams=laparams)
        interpreter = PDFPageInterpreter(rsrcmgr, device)
        for page in PDFPage.get_pages(fp):
            interpreter.process_page(page)
    text = outfp.getvalue()
    return text

phoneRegex = r"^(\+\d{1,}\s?)?1?\-?\.?\s?\(?\d{3}\)?[\s.-]?\d{3}[\s.-]?\d{4}$"

emailRegex = re.compile(r'''(
    [a-zA-Z0-9._%+-] + #username
    @                   # @symbole
    [a-zA-Z0-9.-] +     # domain
    (\.[a-zA-Z]{2,4})   # dot-something
    )''', re.VERBOSE)

print("Enter path:")
path = input()
print("Enter name for output file:")
filenm = input()

files = os.listdir(path)

try:
    os.mkdir("Output")
except:
    pass

cmp = []
for currentpath, folders, files in os.walk('AutoDetect'):
    for file in files:
        item = os.path.join(currentpath, file)
        if ".pdf" in str(item):
            print(item)
            try:
                text = pdf_to_text(item)
            except:
                continue
        elif ".doc" in str(item):
            print(item)
            try:
                text = textract.process(item)
            except:
                continue
        else:
            continue
        for email, phone in zip_longest(emailRegex.findall(str(text)), re.findall(r'[\+\(]?[1-9][0-9 .\-\(\)]{8,}[0-9]', str(text)), fillvalue=''):
            fnl = {}
            try:
                fnl["Email"] = email[0]
            except:
                fnl["Email"] = ""
            fnl["Phone"] = phone
            cmp.append(fnl)
        
pd.DataFrame(cmp).to_excel(f"Output/{filenm}.xlsx", index=False)
print("Completed........")

