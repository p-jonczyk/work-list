defined_cell_occurence = ["X", "SL", "V", "UV", "H"]

# main
welcome_msg = ("\n\nHello!\n"
               "Your workhours record file will be converted into separet files "
               "each for individual worker.\nPlease enter the following data...\n")

incorrect_values_msg = ("\nOne or multiple entered values were incorrect\n"
                        "Please try again...\n")


# files handeling
def required_files_chack_fail_msg(source_file_name):

    msg_exist = ("\nRequired files do not exist in current location...\n"
                 f"Please check if 'base.xlsx' and '{source_file_name}' "
                 "are within current location and confirm their proper names.")
    return msg_exist


def output_folder_check_fail_msg(construction_site):
    msg_not_exist = (f"\nThe following directory '{construction_site} employees worklists' "
                     "already exists.")
    return msg_not_exist


source_file_name_base_msg = "Source file cannot have name: 'base.xlsx'"

help_msg = """\nCreated by Pawel Jonczyk

DATA ENTERING
> 'Break' sholud be greater or equal 0 but less than 24
> 'Month' be number greater than 0 and less or equal than 12
> 'Year' can be set from 1990 to 2100
> 'Work starts' should be number greater than 0 and less than 24
> 'Day (included)' should not be greater than 'Days in the month' 
  but should be greater than 0 and previously entered value
  Entered day (value) will also have start hour as given before entering 'Day'

REQUIRED FILES
For program to work properly you need to have two files in the same directory
as '.exe' file:
> 'base.xlsx' - which is template for individual hourslist
> 'your_source_file_name.xlsx' - which is source file from which all data will be taken

SOURCE FILE
> Hours entered in source file ('your_source_file_name.xlsx') have to be in number format
Any different format will be not considerd in sum of worked hours\n\n"""
