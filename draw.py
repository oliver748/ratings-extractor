from openpyxl.styles import Alignment
from openpyxl.styles import PatternFill, Font, Border
from openpyxl.worksheet.dimensions import ColumnDimension, DimensionHolder
from openpyxl.utils import get_column_letter
from openpyxl.styles.borders import Border, Side
from openpyxl import Workbook


def fill_cell(cell, episode):
    if episode < 7:
        cell.fill = PatternFill("solid", start_color="B04A4A")
    elif episode < 8:
        cell.fill = PatternFill("solid", start_color="F5862B")
    elif episode < 10:
        cell.fill = PatternFill("solid", start_color="1E9A47")
    elif episode == 10:
        cell.fill = PatternFill("solid", start_color="31869B")

def draw_spreadsheet(ratings, amount):
    wb = Workbook()
    ws = wb.active

    j = 1
    for season in ratings:
        i = 1
        j += 1
        cell = ws.cell(row=1, column=season+1, value=season)
        cell.alignment = Alignment(horizontal='center')
        cell.fill = PatternFill("solid", start_color="262626")
        cell.font = Font(color='ffffff')
        for episode in ratings[season]:
            i += 1
            cell = ws.cell(row=i, column=j, value=episode)
            cell.font = Font(color='000000')
            cell.alignment = Alignment(horizontal='center')
            fill_cell(cell, episode)

    #################################### find season with most episodes to put the correct number of episodes on the y axis
    all_seasons = []
    for season in ratings:
        episode_num = len(ratings[season])
        all_seasons.append(episode_num)
    max_episodes = max(all_seasons)

    for i in range(max_episodes+1):
        cell = ws.cell(row=i+1, column=1, value=i)
        cell.alignment = Alignment(horizontal='center')
        cell.fill = PatternFill("solid", start_color="262626")
        cell.font = Font(color='ffffff')
    ws.cell(row=1, column=1, value="") # clears first cell (A1)
    ###################################


    dim_holder = DimensionHolder(worksheet=ws)

    for col in range(ws.min_column, ws.max_column + 1):
        dim_holder[get_column_letter(col)] = ColumnDimension(ws, min=col, max=col, width=5)
    ws.column_dimensions = dim_holder

    thin_border = Border(left=Side(style='thin'), 
                        right=Side(style='thin'), 
                        top=Side(style='thin'), 
                        bottom=Side(style='thin'))

    for col in range(amount+1):
        for row in range(max_episodes+1):
            cell = ws.cell(row=row+1, column=col+1)
            fill = str(cell.fill)[136:144] # extracts background color from specific cell
            if fill == "00000000": # fills specific background color if the episode doesnt exist
                # now i look at this again i couldve just checked if there's text info on the cell instead of all of this :(
                cell.fill = PatternFill("solid", start_color=BACKGROUND_COLOR)
            if BORDER:
                cell.border = thin_border


    wb.save(filename = FILE_NAME)
    print(F"\nEverything is done — you can now close this window and open up '{FILE_NAME}'")