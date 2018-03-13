# @Author: Hannah Young
# ODC - Development
# Choosing a random raffle winner for our March 2018 Gala


import sys
import openpyxl
import random

wbName = raw_input("Workbook name (note, must live in same folder as this script!): ")
print "you entered", wbName
sheetName = raw_input("Sheet name: ")
print "you entered", sheetName

# open up the work book
from openpyxl import load_workbook
wb = load_workbook(filename = wbName)
sheet = wb.get_sheet_by_name(sheetName)

# where we store the donation id's and their weights (and name)
choices = []

#Initialize starting values
count = 2
columnA = "A"
columnB = "B"
currentACell = "A2"
currentBCell = "B2"


# load donation id's and weights into the list
while (sheet[currentACell].value > 0 ):
	choices += [[sheet[currentACell].value, sheet[currentBCell].value]]
	count += 1
	currentACell = columnA + str(count)
	currentBCell = columnB + str(count)

print choices

# pulled algorithm from here:
# https://stackoverflow.com/questions/3679694/a-weighted-version-of-random-choice
def weighted_choice(choices):
	# sums up all of the weighting
   total = sum(w for c, w in choices)
   # picks a random number evenly distributed over our weighting
   r = random.uniform(0, total)
   upto = 0
   # Loops through the weights and compares to the random number choice
   # if the current weight + all the weights we've see so far are greater 
   # than our random number, it wins
   # otherwise keeps looking
   for c, w in choices:
      if upto + w >= r:
         return [c, w]
      upto += w
   assert False, "Shouldn't get here"

# Final raffle results
print weighted_choice(choices)

