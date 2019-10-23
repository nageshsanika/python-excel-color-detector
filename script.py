import xlrd

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

            if len(unique_colors) == MAX_UNIQUE_COLORS :
                row_break = True
                break

        if row_break :
            break

    return unique_colors

def get_blue_shade_colors(colors):
    blues = []
    for color in colors :
        if color:
            r, g, b = color

            # Only Blue Shade Colors
            if b > g and b > r :
                blues.append(color)

    return blues

def sort_list_of_tuples(tup):  
    # reverse = None (Sorts in Ascending order)  
    # key is set to sort using third element ( blue value )
    return(sorted(tup, key = lambda x: x[2]))  

sheet_colors = getUnique_colors_from_sheet(workbook, active_sheet)
blue_colors = get_blue_shade_colors(sheet_colors)

# Sort based on blue value
# lower the blue value darker the color
blue_colors =  sort_list_of_tuples(blue_colors)

result = []
for i, blue in enumerate(blue_colors):
    r, g, b = blue
    # Dark shade of pixels
    if b < 128 :
        result.append({ 'DARK BLUE ' + str(i): blue})
    else :
        result.append({ 'LIGHT BLUE ' + str(i) : blue})

print("Colors in Sheet", sheet_colors, '\n')
print("Only Blue Colors in Sheet", blue_colors,  '\n')
print("Shades of blue", result,  '\n')
