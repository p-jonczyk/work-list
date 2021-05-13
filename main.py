import openpyxl as xl
import calendar
import work_list
import row_style
import os
import files_handeling
import constants as const
import input_handeling


def main():

    # loading the source excel file to get data from
    source_file_path = ".\List of hours.xlsx"
    files_handeling.check_required_files_existence(source_file_path)
    source_ws = xl.load_workbook(source_file_path).worksheets[0]

    # get data from source file
    month_name = source_ws['R1'].value
    construction_site = source_ws['D3'].value

    # loading the base excel file to save data to
    base_file_path = ".\\base.xlsx"
    files_handeling.check_required_files_existence(base_file_path)
    base_file = xl.load_workbook(base_file_path)
    base_ws = base_file.active

    # check output folder existance
    output_folder_path = f".\{construction_site} employees worklists"
    files_handeling.check_required_files_existence(output_folder_path, exist=False,
                                                   construction_site=construction_site)

    # different work hour start in month dict
    work_hour_start_days = {}

    # welcomme message
    print(const.welcome_msg)
    # count for help msg to be shown
    cnt = 0
    # get user's input
    while True:
        try:
            # help msg if problems
            input_handeling.show_help(cnt)

            pause = float(input('Break length: '))
            year = int(input('Year: '))
            month = int(input('Month (1 - 12): '))
            # get days in month
            days_in_month = calendar.monthrange(year, month)[1]
            if input_handeling.check_input(pause, year, month):
                print(
                    f"Chosen month has {days_in_month} days.")
                # input handeling
                input_handeling.handeling_work_hour_days(
                    work_hour_start_days, days_in_month)
                break
            else:
                cnt += 1
                print(const.incorrect_values_msg)
        except ValueError:
            cnt += 1
            print(const.incorrect_values_msg)

    # create new directory after chcecking all conditions
    os.mkdir(f"{output_folder_path}")

    # loading style
    base_file.add_named_style(row_style.base)
    base_file.add_named_style(row_style.days_off)

    # table creation
    work_list.fill_worksheet(source_ws, base_ws, month_name, base_file,
                             month, year, work_hour_start_days, pause, construction_site)

    files_handeling.exit()


if __name__ == "__main__":
    main()
