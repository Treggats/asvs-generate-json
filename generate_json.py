#!/usr/bin/env python
# -*- coding: utf-8 -*-
import getopt
import sys
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

        self.fields = list()
        self.categories = dict()
        self.requirements = dict()
        self.levels = dict()

        self.levels.update({1: {'en': "Oppertunistic"}})
        self.levels.update({2: {'en': "Standard"}})
        self.levels.update({3: {'en': "Advanced"}})

        self.create_categories()
        self.create_requirements()

        self.fields.append(dict(level=self.levels))
        self.fields.append(dict(category=self.categories))
        self.fields.append(dict(requirement=self.requirements))

    def get_json(self, dictionary=''):
        if dictionary == 'level':
            dictionary = self.levels
        elif dictionary == 'category':
            dictionary = self.categories
        elif dictionary == 'requirement':
            dictionary = self.requirements
        elif dictionary == 'all':
            dictionary = self.fields
        return json.dumps(dictionary, indent=4)

    def create_categories(self):
        for row in self.rows:
            cat = row[2].split(': ')
            version = int(cat[0][1:])
            title = cat[1]
            self.categories.update({version: title})

    def create_requirements(self):
        for row in self.rows:
            levels = list()
            number = row[1].split('.')
            cat_id = number[0]
            req_id = number[1]

            if row[4]:
                levels.append(1)
            if row[5]:
                levels.append(2)
            if row[6]:
                levels.append(3)

            self.requirements.update({int(row[0]): {
                "requirement_nr": req_id,
                "category_nr": cat_id,
                "title": {"en": row[3]},
                "levels": levels}})


def usage():
    print("\nUsage:")
    print("-h [--help]")
    print("-i= [--instance=] level, category, requirement or all")
    print("-d (debug)\n")


def main(argv):
    global _debug
    _debug = False
    try:
        opts, args = getopt.getopt(argv, "hdf:i:", [
            "help",
            "instance=",
            "file="])
    except getopt.GetoptError as err:
        print(err)
        sys.exit(2)
    if not opts:
        usage()
        sys.exit(2)
    file_name = ""
    argument = ""
    for opt, arg in opts:
        if opt in ("-h", "--help"):
            usage()
            sys.exit()
        elif opt == "-d":
            _debug = True
        elif opt in ("-f", "--file"):
            file_name = arg
        elif opt in ("-i", "--instance"):
            if arg in ('level', 'category', 'requirement', 'all'):
                argument = arg
            else:
                usage()
        else:
            usage()
        if len(file_name) > 0 and len(argument) > 0:
            generator = GenerateJson(file_name)
            print(generator.get_json(argument))
        if _debug:
            print("Debug is set")

if __name__ == '__main__':
    main(sys.argv[1:])
