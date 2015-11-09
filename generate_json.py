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

        self.categories = list()
        self.requirements = list()
        self.levels = list()
        self.levels.append({'en': "Oppertunistic"})
        self.levels.append({'en': "Standard"})
        self.levels.append({'en': "Advanced"})

        self.create_categories()
        print(self.categories)

    def create_categories(self):
        for row in self.rows:
            cat = row[2].split(': ')
            version = cat[0]
            title = cat[1]
            self.categories[version] = title

    def create_requirements(self):
        for row in self.rows:
            pass

if __name__ == "__main__":
    try:
        asvs_file = "/home/tonko/Downloads/ASVS-excel.xlsx"
        fc = GenerateJson(asvs_file)
    except KeyError:
        print("Boom!")
