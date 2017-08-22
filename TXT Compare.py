# -*- coding: utf-8 -*-
"""
Created on Sat Aug 19 23:21:27 2017

@author: Lloyd Jackman
"""

import pandas as pd
import numpy as np
from fuzzywuzzy import fuzz
import easygui as eg
from os import startfile

title = "TXT Compare"
eg.msgbox(msg="Welcome to TXT Compare", title=title)

first_file = eg.fileopenbox(msg="Select your first file", title=title,default="*.txt")
try:
    first_file_df = pd.read_csv(first_file, sep="\t")
except UnicodeDecodeError:
    first_file_df = pd.read_csv(first_file, sep="\t", encoding='cp1252')
except:
    eg.msgbox(msg="Sorry, we're having trouble parsing %s" % first_file, title=title)

second_file = eg.fileopenbox (msg="Select your second file", title=title,default="*.txt")
try:
    second_file_df = pd.read_csv(second_file, sep="\t")
except UnicodeDecodeError:
    second_file_df = pd.read_csv(second_file, sep="\t", encoding='cp1252')
except:
    eg.msgbox(msg="Sorry, we're having trouble parsing %s" % second_file, title=title)

cols = first_file_df.columns
text_cols = eg.multchoicebox(msg="Select the text fields", title=title, choices=cols)
number_cols = eg.multchoicebox(msg="Select the number fields", title=title, choices=cols)
on = eg.multchoicebox(msg="Select the unique field(s)", title=title, choices=cols)
xref = on[0]

missing_in_first_df = second_file_df[~second_file_df[xref].isin(first_file_df[xref])]
missing_in_first = second_file_df[~second_file_df[xref].isin(first_file_df[xref])][xref]

missing_in_second_df = first_file_df[~first_file_df[xref].isin(second_file_df[xref])]
missing_in_second = first_file_df[~first_file_df[xref].isin(second_file_df[xref])][xref]

universe_diffs = pd.concat([missing_in_first, missing_in_second])

first_file_root = first_file.split('\\')[-1][:-4]
second_file_root = second_file.split('\\')[-1][:-4]

if first_file_root == second_file_root:
    first_file_root = ("-").join(first_file.split('\\')[-2:])[:-4]
    second_file_root = ("-").join(second_file.split('\\')[-2:])[:-4]

df = first_file_df.merge(second_file_df, how='outer', on=on, suffixes=("_%s" % first_file_root, "_%s" % second_file_root))
df = df[~df[xref].isin(universe_diffs)]
df = df.fillna("-")

writer = pd.ExcelWriter("%s vs %s - TXT compare.xlsx" % (first_file_root, second_file_root), engine='xlsxwriter')
df_dict = {}
missing_in_first_df.to_excel(writer, "Missing in %s" % first_file_root, index=False)
missing_in_second_df.to_excel(writer, "Missing in %s" % second_file_root, index=False)
for col in cols:
    if col in on:
        continue
    check = []
    ratio = []
    pc_diff = []
    columns = []
    columns = on + [col+"_"+first_file_root, col+"_"+second_file_root]
    print(columns)
    check_df = df[columns]
    for n in range(df.shape[0]):
        if text_cols is not None:
            if col in text_cols:
                ratio.append(fuzz.ratio(str(check_df.iloc[n, -2]),str(check_df.iloc[n, -1])))
        if number_cols is not None:
            if col in number_cols:
                try:
                    diff = float(round((((check_df.iloc[n, -2])/(check_df.iloc[n, -1])-1)*100),2))
                except:
                    diff = "-"
                pc_diff.append(diff)
        if str(check_df.iloc[n, -2]) == str(check_df.iloc[n, -1]):
            check.append("Match")
        elif str(check_df.iloc[n, -2]) == "-":
            check.append("Data not in CFG")
        elif str(check_df.iloc[n, -1]) == "-":
            check.append("Data not in GFD-1")
        else:
            check.append("Difference")
    check_df["Check"] = np.array(check)
    if text_cols is not None:
        if col in text_cols:
            check_df["Ratio"] = np.array(ratio)
    if number_cols is not None:
        if col in number_cols:
            check_df["% Diff"] = np.array(pc_diff)
    check_df.drop_duplicates(inplace=True)
    check_df.to_excel(writer, '%s' % col, index=False)
    check_df[check_df.Check != "Match"].to_excel(writer, '%s Diff' % col, index=False)
writer.save()

startfile("%s vs %s - TXT compare.xlsx" % (first_file_root, second_file_root))
eg.msgbox(msg="Your file is now ready/nThank you for using TXT Compare!", title=title)
