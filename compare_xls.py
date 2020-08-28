from itertools import zip_longest # Python3 = zip_longest || Python2 = izip_longest
import xlrd # pip install xlrd - for reading xls documents
import xlwt # pip install xlwt - for writing xls documents

# rb1 = pre-changes : rb2 = post-changes
rb1 = xlrd.open_workbook(	"Compare1.xls")
rb2 = xlrd.open_workbook(	"Compare2.xls")
filename = 					"Result.xls"

# Create a new xls workbook and 'Result' sheet tab
workbook = xlwt.Workbook()
sheet = workbook.add_sheet('Result')

# Use "Sheet 1" tab on each xls document
sheet1 = rb1.sheet_by_index(0)
sheet2 = rb2.sheet_by_index(0)

# Function to test whether a string is a number
def CheckIfInt(num):
	try:
		int(num)
		return True
	except ValueError:
		return False

# Compare two excel (.xls) documents
def Compare():
	# Loop through the columns and rows in xls document
	for rownum in range(max(sheet1.nrows, sheet2.nrows)):
		if rownum < sheet1.nrows:
			row_rb1 = sheet1.row_values(rownum)
			row_rb2 = sheet2.row_values(rownum)

			for column, (c1, c2) in enumerate(zip_longest(row_rb1, row_rb2)):
				# Check if a number and if so, minus pre from post
				if CheckIfInt(c1):
					sheet.write(rownum, column, round(c2 - c1, 2))
				# Otherwise, just write out the string
				else:
					sheet.write(rownum, column, c1)

# Run the compare function
Compare()

# Save the workbook into an xls file
workbook.save(filename)
