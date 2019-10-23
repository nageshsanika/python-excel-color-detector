import xlrd
import webcolors

# max color to find for now its just 4
MAX_UNIQUE_COLORS = 4

# open Excel file and get current workbook sheet
file_loc = ("TestExcel.xls")
workbook = xlrd.open_workbook(file_loc, formatting_info=True) 
active_sheet = workbook.sheet_by_index(0)

# get colors == MAX_UNIQUE_COLORS from sheet
def getUnique_colors_from_sheet(book, sheet):
    rows, cols = sheet.nrows, sheet.ncols
    unique_colors = set()
    row_break = False

    # traverse rows * cols
    for row in range(rows):
        for col in range(cols):
            cell_index = sheet.cell_xf_index(row, col)
            cell = book.xf_list[cell_index]
            bgCell = cell.background.pattern_colour_index
            pattern_colour = book.colour_map[bgCell]
            unique_colors.add(pattern_colour)

            if len(unique_colors) == 4 :
                row_break = True
                break

        if row_break :
            break

    return unique_colors

colors = getUnique_colors_from_sheet(workbook, active_sheet)

print(colors)

for color in colors :
    if color:
        r, g, b = color
        if b > g and b > r :
            print(color)
