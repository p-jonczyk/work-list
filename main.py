import openpyxl as xl
import calendar
import work_list
import row_style


def main():

    welcome_msg = """
            Hello!\n           
Your work record workbook will be converted into separet files each for individual worker.
Please enter the following data:\n"""

    # loading the source excel file to get data from
    source_file_path = ".\List of hours.xlsx"
    source_ws = xl.load_workbook(source_file_path).worksheets[0]

    # loading the base excel file to save data to
    base_file_path = ".\\base.xlsx"
    base_file = xl.load_workbook(base_file_path)
    base_ws = base_file.active

    # loading style
    base_file.add_named_style(row_style.base)
    base_file.add_named_style(row_style.days_off)

    # get data from source file
    month_name = source_ws['R1'].value

    # get user's input
    while True:
        print(welcome_msg)
        try:
            work_start = float(input('Work starts at: '))
            pause = float(input('Break length: '))
            year = int(input('Year: '))
            month = int(input('Month (1 - 12): '))
            break
        except ValueError:
            print("You have to enter numbers.\n")

    # table handeling
    work_list.create_month_table(base_ws, year, month)
    work_list.fill_worksheet(source_ws, base_ws, month_name, base_file,
                             month, year, work_start, pause)
    work_list.exit()


if __name__ == "__main__":
    main()
