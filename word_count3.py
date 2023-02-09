#!/usr/bin/env python3
import os
import sys
import argparse
import zipfile
import shutil
import tempfile
import xml.etree.ElementTree as ET
import time
import docx2txt

from progress.bar import Bar


def count_fodt(filename):
    root = ET.parse(filename).getroot()
    return sum(len(text.split()) for text in root.itertext())


def count_docx(filename):
    doc = docx2txt.process(filename)
    #return sum(len(text.split()) for text in doc.itertext())
    return len(doc.split())

def count_odt(filename):
    with tempfile.NamedTemporaryFile() as tmp_file:
        with zipfile.ZipFile(filename) as odt_file:
            with odt_file.open("content.xml") as content_file:
                shutil.copyfileobj(content_file, tmp_file)
                tmp_file.seek(0)
        word_count = count_fodt(tmp_file.name)
    return word_count


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description='Count words in Open Document text file'
    )
    parser.add_argument("filename", help="Name of odt or fodt file")
    parser.add_argument("goal", help="")
    parser.add_argument('indent', help='add indent')
    parser.add_argument('relative', help='T or F')


    if len(sys.argv) == 1:
        parser.print_help()
        sys.exit(1)
    args = parser.parse_args()
    if os.path.splitext(args.filename)[1].lower() == ".fodt":
        word_count = count_fodt(args.filename)
    else:
        #word_count = count_odt(args.filename)
        word_count = count_docx(args.filename)
    #print(word_count)
    print('\n')
    #while True:
    goal = int(args.goal)
    start = word_count-23
    #print(start)
    bar = Bar('Loading', fill='#', suffix='%(percent)d%%')

    if args.indent:
        pre_text = '\t\t\t\t\t\t\t\t'
    else:
        pre_text = 'count \t'

    if args.relative == None or args.relative == 'F':
        range_ = start+goal
    else:
        range_ = goal

    with Bar(pre_text, max=range_) as bar:
        for i in range(start):
            time.sleep(0.0001)
            bar.next()
        
        word_count_old = word_count
        #Check file
        while True:
            time.sleep(10)
            word_count_new = count_docx(args.filename)
            
            if word_count_new > word_count_old:
                for i in range(word_count_new-word_count_old):
                    bar.next()

                word_count_old = word_count_new


