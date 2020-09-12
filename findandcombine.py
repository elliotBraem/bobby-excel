import os
from fnmatch import fnmatch
import xlrd
import xlwt

root = './data' # root directory of files
pattern = "*.xlsx" # regex for desired file extension

# let's go for a walk into all of the paths, sub directorys, and files in the root directory
for path, subdirs, files in os.walk(root):
  # when we get into a directory, lets get all of the files
    for name in files:
      # when we have a file, let's make sure it has our desired file extension
        if fnmatch(name, pattern):
          # print the name of the file we found
            print(name)


# now let's do it again, but save the file names to a list
target_files = []

for path, subdirs, files in os.walk(root):
    for name in files:
        if fnmatch(name, pattern):
            target_files.append(name)

print(target_files)

# ok, now we have the files, let's work on extracting data
# we need to download some new libraries (https://www.sitepoint.com/using-python-parse-spreadsheet-data/),
# so let's create a virtual environment

# run this in the terminal (https://code.visualstudio.com/docs/python/python-tutorial):
# python3 -m venv .venv
# source .venv/bin/activate

# now install the packages 

# let's open a workbook
# import xlrd
workbook = xlrd.open_workbook('./data/1/sample1.xlsx')

# now we have a workbook open, to open a specific sheet inside the workbook, we can do it two ways:

# by sheet name (im going to do this one)
worksheet = workbook.sheet_by_name('SalesOrders')

# or by sheet index
# worksheet = workbook.sheet_by_index(0)

# now we can get data from the sheet as well
print(worksheet.cell(2, 3).value) # should be OrderDate

# how about we want to grab all of the names of the orders? 
#for cell in worksheet.col(2):
  # print(cell.value)
# that's simple, we just need to know the column

# what if we want to grab all of the names and their order?
for i in range(worksheet.nrows):
  print("Name: " + worksheet.cell(i, 2).value + ", Order: " + worksheet.cell(i, 3).value)
# it's a little more complicated, it's smarter to traverse by row and grab data as we need it
# btw, it goes cell(row, col)

# ok, how about we filter only the names that bought Pencils
for i in range(worksheet.nrows):
  item = worksheet.cell(i, 3).value
  if (item == "Pencil"):
    print(worksheet.cell(i, 2).value)

# now I want to write all these names in a new excel
# let's create a new work book (notice xlwt vs xlrd, one is xlWRITE, one is xlREAD)
new_workbook = xlwt.Workbook()
# create a sheet
new_sheet = new_workbook.add_sheet("Sheet")
# and save
new_workbook.save('target_workbook.xls')

# now it should show up in your project

# let's write all the names that bought Pencil to the workbook
# first let's do it the long way, let's find all the names, add them to a list, then write them
names = []

for i in range(worksheet.nrows):
  item = worksheet.cell(i, 3).value
  if (item == "Pencil"):
    name = worksheet.cell(i, 2).value
    names.append(name)

for i, name in enumerate(names):
  new_sheet.row(i).write(0, name)

new_workbook.save('target_workbook.xls')

# save it, finished