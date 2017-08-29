################################################################################
## ABOUT															          ##
################################################################################

# ---- AUTHOR ----
# Hannah Young
# Development Associate | ODC
# August 28, 2017

# ---- OVERVIEW ----
# This script takes in the exported JSON from a Trello board and converts
# it into an excel document. It only has information from the list titles
# and the card titles, but could easily be extended to have more.

# ---- INSTRUCTIONS FOR USE ----
# Please note that these instructions are intended to be accessible to those
# who may not have a high familiarity with Python.
#
# 0. If you are not familiar with programming, use a Mac--this is doable on a
#    PC but it involves extra work not documented here.
# 1. Update python to 2.7.12 (needs to be 2.7 for library to work)
# 2. To install the excel libraries, use "pip install pandas" and
#    "pip install xlrd" (on Mac)
# 3. Export JSON from trello board you wish to convert.
# 4. To run, navigate to folder where this script is stored
#    in terminal and then type "python ./trelloConverter.py". Make sure that
#    your JSON output from Trello exists in the same folder.
# 5. You will be prompted to input the name of the JSON export, the name of the
#    excel WB you wish to create, and the name of the sheet you wish to add this
#    to.
# 6. Tadah--data is exported to excel! Manipulate, format, arrange as desired.

################################################################################
## LOGIC															          ##
################################################################################
# Libraries for working with data
import json
import numpy as np
import pandas as pd
from pandas import ExcelWriter

# Library for testing work while editing this script
from pprint import pprint

# Gather user inputs
print "Welcome!"
print "Enter name of json exported from Trello:",
filename = raw_input()

print "Enter a workbook name, either existing or new:"
newWBName = raw_input()

print "Enter a sheet name, either exisiting or new:"
newSheetName = raw_input()
print "Your file is now generated! Go ODC!"

# imports the trello json export as an object we can
# work with and manipulate
with open(filename) as data_file:
    data = json.load(data_file)

# Format: {listID: [card1, card2, ... , cardN]}
cardObj = {}

# Format: {listID1:Name1, ... , listIDN: NameN}
listObj = {}

# Format : {Name1: [card1, card2,...], ..., NameN: [cardn-2, ..., cardN]}
combinedObj = {}

# Note that these both generate indexed lists, not OBjs,
# at the highest level
justCards = data["cards"]
justTrelloLists = data["lists"]

# Associate all card names with the list ID
for x in range(len(justCards)):
    name = justCards[x]["name"]
    trelloList = justCards[x]["idList"]

    if trelloList in cardObj:
        cardObj[trelloList] += [name]
    else:
        cardObj[trelloList] = [name]

# Associate all List names with the list ID
for y in range(len(justTrelloLists)):
    listName = justTrelloLists[y]["name"]
    trelloList = justTrelloLists[y]["id"]
    listObj[trelloList] = listName

# Combine listObj and cardObj into combinedObj
for key in listObj:
    newKey = listObj[key]

    # This check shouldn't be necessary, but does ensure safety
    # in the off chance something goes wrong between the different
    # ID's--leaving it in for now since I have not had a chance for
    # thorough testing
    if key in cardObj:
        for z in range(len(cardObj[key])):
            if newKey in combinedObj:
                combinedObj[newKey] += [cardObj[key][z]]
            else:
                combinedObj[newKey] = [cardObj[key][z]]

# Take the data and load into excel to generate report
df = pd.DataFrame.from_dict(combinedObj, orient='index')
df.transpose()
writer = pd.ExcelWriter(newWBName + '.xlsx')
df.to_excel(writer, sheet_name=newSheetName)
writer.save()
