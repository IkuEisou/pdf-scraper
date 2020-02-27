import requests
import os
import camelot
import pdfplumber
import re
import sys
import getopt

from openpyxl import Workbook
from PyPDF2 import PdfFileReader


def downloadFile(url, filename):
    r = requests.get(url, stream=True)
    with open(filename, 'wb') as f:
        f.write(r.content)


def decryptPDF(filename):
   # open the pdf file
    fp = open(filename, mode="rb")
    pdfFile = PdfFileReader(fp)
    if pdfFile.isEncrypted:
        try:
            pdfFile.decrypt('')
            print('File Decrypted (PyPDF2)')
        except:
            command = ("cp " + filename +
                       " temp.pdf; qpdf --password='' --decrypt temp.pdf " + filename
                       + "; rm temp.pdf")
            os.system(command)
            print('File Decrypted (qpdf)')
            fp = open(filename, mode="rb")
            pdfFile = PdfFileReader(fp)


def searchPDF(key, pdfFile):
    # get number of pages
    NumPages = pdfFile.getNumPages()

    # define keyterms
    String = "環境"

    # extract text and do the search
    for i in range(0, NumPages):
        PageObj = pdfFile.getPage(i)
        print("this is page " + str(i))
        Text = PageObj.extractText().encode("shift_jisx0213").decode("shift_jisx0213")
        print(Text)
        ResSearch = re.search(String, Text)
        print(ResSearch)


def extractPDF2Text(filename):
    with pdfplumber.open(filename) as pdf:
        for page in pdf.pages:
            print(page.extract_text())


def extractTables(filename, page):
    tables = camelot.read_pdf(filename, pages=page,
                              # flavor='stream',
                              #   strip_text=' .\n'
                              )
    for table in tables:
        df = table.df
        headers = df[0][0].split('\n')
        header_len = len(headers)
        col_len = len(df.columns)
        if header_len > col_len:
            for i in range(col_len, header_len):
                headers[i % col_len] += headers[i]
            headers = headers[0:col_len]
        if header_len > 0:
            for col in range(0, col_len):
                if col > 0 and df[col][0] != '':
                    df[col][1] = df[col][0]
                df[col][0] = headers[col]

    output = filename.split('.')[0]
    # export all tables at once to CSV files
    tables.export(output+".csv", f="csv")

    # export all tables at once to CSV files in a single zip
    # tables.export("camelot_tables.csv", f="csv", compress=True)

    # export each table to a separate worksheet in an Excel file
    # tables.export(output+'.xlsx', f="excel")


def main(argv):
    helpTips = ' -i <url> -p<page num> -o <outputfile>'
    inputfile = ''
    page = '1'
    outputfile = ''
    try:
        opts, args = getopt.getopt(
            argv, "hi:p:o:", ["ifile=", "page=", "ofile="])
    except getopt.GetoptError:
        print(helpTips)
        sys.exit(2)
    for opt, arg in opts:
        if opt == '-h':
            print(helpTips)
            sys.exit()
        elif opt in ("-i", "--ifile"):
            inputfile = arg.strip()
        elif opt in ("-p", "--page"):
            page = arg.strip()
        elif opt in ("-o", "--ofile"):
            outputfile = arg.strip()

    if outputfile == '':
        filename = inputfile.split('/')[-1]
    else:
        filename = outputfile
    downloadFile(inputfile, filename)
    decryptPDF(filename)
    extractTables(filename, page)


if __name__ == "__main__":
    main(sys.argv[1:])
