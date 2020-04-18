#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Fri Apr 10 23:52:06 2020

@author: alfiantjandra
"""
import pandas as pd
import numpy as np
from openpyxl import load_workbook
import sys

argc = len(sys.argv)
argv = sys.argv

if not (argc < 4 and argc > 1):
    print("Usage: python3 [name_of_file]")
    sys.exit()

excel_filename = argv[1]
workbook=load_workbook(filename=excel_filename)
sheet = workbook.active



total_courses = 0
min_courses_line = 2

for value in sheet.iter_rows(min_row=1,max_row=10000,min_col=1,max_col=14,values_only=True):
    if value[0] == "Course ID":
        break
    else: 
        min_courses_line+=1
        
        
for value in sheet.iter_rows(min_row=5,max_row=10000,min_col=1,max_col=14,values_only=True):
    if value[0] == "Grand Total":
        break
    else: 
        total_courses += 1
        
""" Total Enrollment """
total_dict = {"General Education":0}
name_tracker_1 = "General Education"
for value in sheet.iter_rows(min_row=min_courses_line,max_row=total_courses+min_courses_line-1,min_col=1,max_col=14,values_only=True):
    if value[4] == name_tracker_1:
        total_dict[value[4]] += value[13]
    else:
       name_tracker_1 = value[4]
       if value[4] in total_dict:
           total_dict[value[4]]+=value[13]
       else:
           total_dict[value[4]] = value[13]
               
""" Undergrad """
total_undergrad = {"General Education":0}
name_tracker_2 = "General Education"
for value in sheet.iter_rows(min_row=min_courses_line,max_row=total_courses+min_courses_line-1,min_col=1,max_col=14,values_only=True):
    if value[4] == name_tracker_2:
        total_undergrad[value[4]] += value[6]
    else:
       name_tracker_2 = value[4]
       if value[4] in total_undergrad:
           total_undergrad[value[4]]+=value[6]
       else:
           total_undergrad[value[4]] = value[6]

""" Cross-Registration """
total_cross = {"General Education":0}
name_tracker_3 = "General Education"
for value in sheet.iter_rows(min_row=min_courses_line,max_row=total_courses+min_courses_line-1,min_col=1,max_col=14,values_only=True):
    if value[4] == name_tracker_3:
        total_cross[value[4]] += value[9]
    else:
       name_tracker_3 = value[4]
       if value[4] in total_cross:
           total_cross[value[4]]+=value[9]
       else:
           total_cross[value[4]] = value[9]

""" Graduate students """
total_grad = {"General Education": 0}
name_tracker_4 = "General Education"
for value in sheet.iter_rows(min_row=min_courses_line, max_row=total_courses + min_courses_line - 1, min_col=1,
                             max_col=14, values_only=True):
    if value[4] == name_tracker_4:
        total_grad[value[4]] += value[7]
    else:
        name_tracker_4 = value[4]
        if value[4] in total_grad:
            total_grad[value[4]] += value[7]
        else:
            total_grad[value[4]] = value[7]

total_frame = pd.DataFrame(total_dict.items(), columns=['Department', 'Total Enrollment'])
undergrad_frame = pd.DataFrame(total_undergrad.items(), columns=['Department', 'Undergrad Enrollment'])
crossreg_frame = pd.DataFrame(total_cross.items(), columns=['Department', 'Cross-Registration'])
grad_frame = pd.DataFrame(total_grad.items(), columns=['Department', 'Grad Enrollment'])

crossreg_sorted = crossreg_frame.sort_values(by="Cross-Registration",ascending=False)

# print(total_frame)
# print(undergrad_frame)
# print(crossreg_frame)
# print(crossreg_sorted)

""" Top 5 """
dep_names = total_frame['Department'].values

# Crossreg
top5_crossreg = crossreg_sorted.head(5)
crrg_dep_titles = top5_crossreg["Department"].values.tolist()
crrg_dep_values = top5_crossreg["Cross-Registration"].values.tolist()

# Non-undergrad
non_undergrad_series = (total_frame["Total Enrollment"]
                         - undergrad_frame["Undergrad Enrollment"])
non_undergrad_count = non_undergrad_series.to_numpy()
non_undergrad = pd.DataFrame({'Department': dep_names,
                              'Non-undergrad': non_undergrad_count})
top5_non_undergrad = non_undergrad.sort_values(by="Non-undergrad", ascending=False).head(5)

# Graduate students
top5_grad = grad_frame.sort_values(by="Grad Enrollment", ascending=False).head(5)
grad_dep_titles = top5_grad["Department"].values.tolist()
grad_dep_values = top5_grad["Grad Enrollment"].values.tolist()

def writeOutput():
    """ File output below """

    output_file_title = excel_filename[:-5] + "_top5.txt"
    print(output_file_title)
    with open(output_file_title, 'w') as outfile:
        outfile.write("Crossreg top 5:\n")
        for i, (x, y) in enumerate(zip(crrg_dep_titles, crrg_dep_values)):
            outfile.write("{}: {:32s}  {}\n".format(i, x, y))
        outfile.write('\n')
        outfile.write("Grad top 5:\n")
        for i, (x, y) in enumerate(zip(grad_dep_titles, grad_dep_values)):
            outfile.write("{}: {:32s}  {}\n".format(i+1, x, y))

writeOutput()