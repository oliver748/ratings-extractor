from openpyxl.styles import Alignment
from openpyxl.styles import PatternFill, Font, Border
from openpyxl.worksheet.dimensions import ColumnDimension, DimensionHolder
from openpyxl.utils import get_column_letter
from openpyxl.styles.borders import Border, Side
from openpyxl import Workbook

FILE_NAME = "episode_ratings.xlsx"
BACKGROUND_COLOR = "262626"
LOW_RATING = "B04A4A"
MID_RATING = "F5862B"
HIGH_RATING = "1E9A47"
PERFECT_RATING = "31869B"
SEA_EPS_CELL_COLOR = "262626" # color for cells in rows and columns of what episode and season it is
SEA_EPS_FONT_COLOR = "FFFFFF" # color for text in rows and columns of what episode and season it is
BORDER = True
THIN_BORDER = Border(left=Side(style='thin'), 
                    right=Side(style='thin'), 
                    top=Side(style='thin'), 
                    bottom=Side(style='thin'))

def fill_cell(cell, value):
    episode = float(value)
    if episode < 7:
        cell.fill = PatternFill("solid", start_color=LOW_RATING)
    elif episode < 8:
        cell.fill = PatternFill("solid", start_color=MID_RATING)
    elif episode < 10:
        cell.fill = PatternFill("solid", start_color=HIGH_RATING)
    elif episode == 10:
        cell.fill = PatternFill("solid", start_color=PERFECT_RATING)

def draw_spreadsheet(ratings, amount):
    wb = Workbook()
    ws = wb.active

    j = 2
    for season in ratings:
        i = 2 
        cell = ws.cell(row=1, column=season+1, value=season)
        cell.alignment = Alignment(horizontal='center')
        cell.fill = PatternFill("solid", start_color=SEA_EPS_CELL_COLOR)
        cell.font = Font(color=SEA_EPS_FONT_COLOR)
        for episode in ratings[season]:
            cell = ws.cell(row=i, column=j, value=episode)
            cell.alignment = Alignment(horizontal='center')
            fill_cell(cell, episode)
            i += 1
        j += 1

    #################################### find season with most episodes to put the correct number of episodes on the y axis
    all_seasons = []
    for season in ratings:
        episode_num = len(ratings[season])
        all_seasons.append(episode_num)
    max_episodes = max(all_seasons)

    for i in range(max_episodes+1):
        cell = ws.cell(row=i+1, column=1, value=i)
        cell.alignment = Alignment(horizontal='center')
        cell.fill = PatternFill("solid", start_color=SEA_EPS_CELL_COLOR)
        cell.font = Font(color=SEA_EPS_FONT_COLOR)
    ws.cell(row=1, column=1, value="") # clears first cell (A1)

    # centers every cell
    dim_holder = DimensionHolder(worksheet=ws)
    for col in range(ws.min_column, ws.max_column + 1):
        dim_holder[get_column_letter(col)] = ColumnDimension(ws, min=col, max=col, width=5)
    ws.column_dimensions = dim_holder

    # formats the rest: background color and border
    for col in range(amount+1):
        for row in range(max_episodes+1):
            cell = ws.cell(row=row+1, column=col+1)
            fill = str(cell.fill)[136:144] # extracts background color from specific cell
            if fill == "00000000": # fills specific background color if the episode doesnt exist
                cell.fill = PatternFill("solid", start_color=BACKGROUND_COLOR)
            if BORDER:
                cell.border = THIN_BORDER

    wb.save(filename = FILE_NAME)
    print(F"\nEverything is done — you can now close this window and open up '{FILE_NAME}'")
