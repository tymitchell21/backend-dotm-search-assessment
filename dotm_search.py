#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Given a directory path, search all files in the path for a given text string
within the 'word/document.xml' section of a MSWord .dotm file.
"""
__author__ = "Tyler Ward"

from zipfile import ZipFile
import os
import argparse
import time

def decode(fileName):
    """reads document.xml from dotm zip file"""

    with ZipFile(fileName, 'r') as zip:
        document_xml = zip.read('word/document.xml', pwd=None)
        return document_xml.decode()

def find_text(dotm_str, path, text):
    """searches through dotm string to find a substring"""

    text_location = dotm_str.find(text)
    if text_location != -1:
        substring = dotm_str[text_location-40:text_location+40+len(text)]
        print('Match found in file ' + path)
        print('...' + substring + '...\n')
        return True
    return False

def directory_loop(directory, text):
    """loops through directories to find text in dotm files"""
    count_searched = 0
    count_matched = 0
    for root, dirs, files in os.walk(directory):
        for filename in files:
            if filename[-4:] == 'dotm':
                count_searched += 1
                full_path = os.path.join(directory, filename)
                if find_text(decode(full_path), full_path, text): 
                    count_matched += 1
    print(f'Total dotm files searched: {count_searched}')
    print(f'Total dotm files matched: {count_matched}')

def main():
    """Add your code here"""

    parser = argparse.ArgumentParser()

    parser.add_argument('--dir', help='enter a directory')
    parser.add_argument('text')
    args = parser.parse_args()
    
    if args.dir:
        print(f'Searching directory {args.dir} for text "$" ...\n\n')
        directory_loop(args.dir, args.text)
    else:
        print('Searching current directory for text "$" ...\n\n')
        directory_loop('./', args.text)

if __name__ == '__main__':
    main()