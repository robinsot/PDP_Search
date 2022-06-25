# PDP_Search
# A powershell script to search a file list in an excel file for a set of strings listed in another file.  The found output is a .csv file
#
# Author T. Robinson
# 6/18/2022
#
# 1. Read the string table to get a list of strings to search
# 2. Search each file listed in the source file for presence of any search string
# 3. Output matches as a CSV to the output file
#
# args[0] - the search string file (full path)
# args[1] - the search string sheet name 
# args[2] - the source file (This list of files - all should be full path)
# args[3] - the source file sheet
# args[4] - the output file name (full path)
