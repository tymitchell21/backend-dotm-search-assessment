#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Given a directory path, search all files in the path for a given text string
within the 'word/document.xml' section of a MSWord .dotm file.
"""
__author__ = "Tyler Ward"

from zipfile import ZipFile
from xml.etree.cElementTree import XML
from xml.etree import ElementTree
import os
import argparse
import time

def decode(fileName):
    with ZipFile(fileName, 'r') as zip:
        document_xml = zip.read('word/document.xml', pwd=None)
        tree = XML(document_xml)
        xmlstr = ElementTree.tostring(tree).decode()
        return xmlstr

def find_money(dotm_str, path, text):
    text_location = dotm_str.find(text)
    if text_location != -1:
        substring = dotm_str[text_location-40:text_location+41]
        print('Match found in file ' + path)
        print('...' + substring + '...\n')
        return 1
    return 0

def directory_loop(directory, text):
    count_searched = 0
    count_matched = 0
    for root, dirs, files in os.walk(directory):
        for filename in files:
            if filename[-4:] == 'dotm':
                count_searched += 1
                count_matched += find_money(decode('./dotm_files/' + filename), directory + filename, text)
    print('Total dotm files searched: ' + str(count_searched))
    print('Total dotm files matched: ' + str(count_matched))

def main():
    """Add your code here"""
    parser = argparse.ArgumentParser()

    parser.add_argument('--dir', help='enter a directory')
    parser.add_argument('text')
    args = parser.parse_args()
    print(args)
    
    if args.dir:
        print('Searching directory %s for text "$" ...\n\n' % args.dir)
        directory_loop(args.dir, args.text)
    else:
        print('Searching current directory for text "$" ...\n\n')
        directory_loop('./', args.text)

if __name__ == '__main__':
    main()