import sys
import os
import row_style
import openpyxl as xl
import constants as cosnt


def check_required_files_existence(file_path, exist=True, construction_site=None):
    """Checks if required files existing in directory

    Parameters: 

    file_path (str): path to be checked"""

    if exist == True:
        check = os.path.exists(file_path)
        if check == False:
            print(cosnt.msg_exist)
            exit(fail=True)
    else:
        check = os.path.exists(file_path)
        if check == True:
            print(cosnt.output_folder_check_fail_msg(construction_site))
            exit(fail=True)


def exit(fail=False):
    """Exit program by pressing a key

    Parameters:

    fail (boolen): flag for successfullnes of program"""

    if fail == False:
        input("\nData was converted successfully.\nPress key to exit...")
    else:
        input("\nSomething went wrong.\nCheck shown information and try again.\nPress key to exit...")
    sys.exit()
