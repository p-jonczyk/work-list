import openpyxl as xl
import calendar
import sys

# TODO: time format for cells
# TODO: int(None) ? why ?


def create_month_table(base_ws, year, month):
    """Creats table with rows number same as days in month
    and insert it into base file

    Parameter:

    base_ws : activated loaded base file

    year (int): current year 

    month (int): order of current month """

    # insert proper quantity of rows based on days of current month and days off (color)
    cal = calendar.Calendar()
    for counter, day in cal.itermonthdays2(year, month):
        if counter != 0:
            # use 10 row + counter becuase we want to start with 11 row - counter starts from 1
            current_row_number = 10 + counter
            base_ws.insert_rows(current_row_number)
            # number each day of month from 1 to number of days
            current_cell = f'A{current_row_number}'
            base_ws[current_cell] = counter
            # marking  Saturdays and Sundays
            # first week day is 0 so 5th and 6th are Saturday and Sunday respectively
            if day == 5 or day == 6:
                for cells in base_ws[f"{current_row_number}:{current_row_number}"]:
                    cells.style = "days_off"
            else:
                for cells in base_ws[f"{current_row_number}:{current_row_number}"]:
                    cells.style = "base"


def fill_rows(base_ws, row, max_col, work_start, pause):
    """Fills each row with data obtained from source file

    Parameters:

    base_ws: activated loaded base file

    row (list): listo of values from current row

    max_col (int): number of columns in row (length of row list)

    work_start (int/float): time at which work/shift starts

    pause (int/float): length of workers' break"""

    for j, value in enumerate(row[2:max_col]):
        work_start_cell = f'B{11+j}'
        pause_cell = f'C{11+j}'
        work_end_cell = f'D{11+j}'
        work_hours_cell = f'E{11+j}'

        try:
            value = int(value)
            base_ws[work_start_cell] = work_start
            base_ws[pause_cell] = pause
            base_ws[work_end_cell] = work_start + pause + value
            base_ws[work_hours_cell] = value
        except ValueError:
            base_ws[work_start_cell] = value
            base_ws[pause_cell] = value
            base_ws[work_end_cell] = value
            base_ws[work_hours_cell] = value


def fill_sign_date(base_ws, year, month, days_in_month):
    """Fills date of sign field in lower part of table

    Parameters: 

    base_ws: activated loaded base file

    year (int): current year

    month (int): order of current month

    days_in_month (int): number of days in month """

    sign_cell = f'B{13+days_in_month}'
    if month < 10:
        base_ws[sign_cell] = (f'{days_in_month}.0{month}.{year}')
    else:
        base_ws[sign_cell] = (f'{days_in_month}.{month}.{year}')


def fill_worksheet(source_ws, base_ws, month_name, load_base, month, year, work_start, pause):
    """Filling worksheet with data from source file

    Parameters:

    source_ws: loaded source worksheet

    base_ws: activated loaded base file

    month_name (str): name of current month

    load_base: loaded base file (before activation)

    month (int): order of current month

    year (int): current year

    work_start (int/float): time at which work/shift starts

    pause (int/float): length of workers' break """

    # get max row/col
    max_row = source_ws.max_row + 1
    max_col = source_ws.max_column - 1
    # getting number of days in current month
    days_in_month = calendar.monthrange(year, month)[1]
    # filling month name in base file
    base_ws['G7'] = month_name

    # fill table with source file's data
    for row in source_ws.iter_rows(6, max_row, 2, max_col, values_only=True):
        base_ws['D5'] = row[0]
        base_ws['D3'] = row[1]
        # row filler
        fill_rows(base_ws, row, max_col, work_start, pause)
        # fills date for sign place - last day of month
        fill_sign_date(base_ws, year, month, days_in_month)
        # setting sum formula for workhours of month
        base_ws[f'E{11+days_in_month}'] = (f"=SUM(E11:E{10+days_in_month})")
        # saving data into new file with "Name Surname" filename
        load_base.save(f"{row[0]}.xlsx")


def exit():
    """Exit program by pressing a key"""
    input("\nData was converted successfully.\nPress key to exit...")
    sys.exit()
