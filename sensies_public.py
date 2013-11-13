######################################
#   dependencies
######################################
import gspread

######################################
#   vars and definitions 
######################################
targetBook = "IRR Table"    # define the GOOG workbook
targetWorksheet = 1     # define the worksheet to work in 
xRows = [2,2]   # define x-axis rows for selection
xCols = [10,14] # define x-axis cols for selection
yRows = [3,9]   # define y-axis rows for selection
yCols = [9,9]   # define y-axis cols for selection
fillRows = [3,9]    # define sensie table rows
fillCols = [10,14]  # define sensie table cols
var_xLocation = "B2"    # define the location of the x-variable you are varying 
var_yLocation = "B5"    # define the location of the y-variable you are varying 
outputVar = [6,2]   # define the location of the output variable you are running the sensitivity analysis against and filling into the table (e.g., IRR, MOIC, DPI), by [row,column]

######################################
#   instantiations   
######################################
gc = gspread.login("email@gmail_domain.com", "password")  # define the gspread object with login/pw info
targetWorksheet = 1     # define the worksheet to work in 
book = gc.open(targetBook)  # enter the workbook
wks = book.get_worksheet(targetWorksheet)   # enter the target worksheet

######################################
#   methods   
######################################
def getSelection(rows, cols):
    contents = []
    currRow = rows[0]
    currCol = cols[0]
    while currRow <= rows[1]:
        currCol = cols[0]
        while currCol <= cols[1]:
            curCell = wks.cell(currRow,currCol).value
            contents.append(curCell)
            currCol += 1
        currRow += 1 
    return contents

def fillSelection(rows, cols):
    currRow = rows[0]
    currCol = cols[0]
    xidx = 0
    while currCol <= cols[1]:   # iterate through the columns
        currRow = rows[0]
        wks.update_acell(var_xLocation, x[xidx])
        yidx = 0
        while currRow <= rows[1]:   # iterate through the rows
            wks.update_acell(var_yLocation, y[yidx])
            output = wks.cell(outputVar[0], outputVar[1]).value
            wks.update_cell(currRow, currCol, output)
            print "[row, col]: [" + str(currRow) + ", " + str(currCol) + "] output -> " + output 
            currRow += 1
            yidx += 1
        currCol += 1
        xidx += 1

######################################
#   actions   
######################################
x = getSelection(xRows, xCols)  # get x-axis sensitivities 
y = getSelection(yRows, yCols)  # get y-axis sensitivities 
print "x-axis: " + str(x) 
print "y-axis: " + str(y)
fillSelection(fillRows, fillCols)   # fill the sensitivity table with your output at these coords
