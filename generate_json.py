#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
from pathlib import Path
from codecs import open
from unicodedata import name as uni_name
import csv
import json
import pyexcel.ext.xls


class GenerateJson(object):

    """
    Generate json from csv
    Convert the xls(x) to csv and this script will convert it to json
    """

    def __init__(self, xls_file):
        sheet = pyexcel.get_sheet(file_name=xls_file)
        self.header = sheet.row[0]
        self.rows = sheet.row[1:]

        self.categories = dict()
        self.requirements = dict()
        self.levels = dict()

        self.levels.update({1: {'en': "Oppertunistic"}})
        self.levels.update({2: {'en': "Standard"}})
        self.levels.update({3: {'en': "Advanced"}})

        self.create_categories()

    def create_categories(self):
        for row in self.rows:
            cat = row[2].split(': ')
            version = int(cat[0][1:])
            title = cat[1]
            self.categories.update({version: title})

    def create_requirements(self):
        for row in self.rows:
            pass

if __name__ == "__main__":
    try:
        asvs_file = "ASVS-excel.xlsx"
        fc = GenerateJson(asvs_file)
    except KeyError:
        print("Boom!")
