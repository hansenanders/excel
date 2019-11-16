#!/usr/bin/env python3
import argparse
import json
import locale
import pandas
import xlrd
import xlwt

import pprint

import warnings
warnings.simplefilter(action='ignore', category=FutureWarning)

def sanitize(df, col_name):
    for idx in df.index:
        v = df.at[idx, col_name]
        v = v.replace(".", "")
        v = v.replace(",", ".")
        df.set_value(idx, col_name, v)

def main():
    parser = argparse.ArgumentParser(
        description='Sanitize Nordeas Excel Sheets')
    parser.add_argument('infile', nargs='?')
    parser.add_argument('outfile', nargs='?')
    args = parser.parse_args()


    df = pandas.read_excel(args.infile)

#    pprint.pprint(df)
    sanitize(df, 'Belopp')
    with open('kategorier.json') as fn:
        categories = json.load(fn)
#    pprint.pprint(kategorier)

    def get_category(value):
        for k in categories.keys():
            if any(s in value for s in categories[k]):
                return k
        return None

    result = dict()
    for idx in df.index:
        v = df.at[idx, 'Transaktion']
        cat = get_category(v)
        if cat:
            try:
                result[cat] += float(df.at[idx, 'Belopp'])
            except KeyError: # key not found
                result[cat] = 0
        else:
            print("Failed to match: (%s), %r" % (v, df.at[idx, 'Belopp']))

    pprint.pprint(result)
    out = pandas.ExcelWriter(args.outfile)
    df.to_excel(out, 'Transactions')
    out.save()
    print("Wrote file: ", args.outfile)



if __name__ == "__main__":
    main()
