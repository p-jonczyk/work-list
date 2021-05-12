import constants as const


def check_input(pause, year, month):
    """ Return True when all conditions are met

    Parameters:

    pause (int/float): length of workers' break

    year (int): current year

    month (int): order of current month """

    if (0 <= pause <= 24) and (0 < month <= 12) and (1990 <= year <= 2100):
        return True
    else:
        return False


def show_help(cnt):
    """Ask if user want to see 'help' message if wrong data were entered

    Parameter:

    cnt (int): counter of wrong data entrance"""

    if cnt > 0:
        help_ask = input("Do you need help ? (yes/no)\n> ")
        if help_ask.lower() == 'yes':
            print(const.help_msg)
        else:
            cnt = 0


def check_day_range_wrong(work_hour_start_days, work_start, day_range_selection, days_in_month, previous_day_entered=0):
    """Checks if proper value of upper boundary of day to be considered
    for given 'work start' range was entered:

    Parameters:

    work_hour_start_days (dict): dictionary of pairs {day_range_selection : work_start} 

    work_start (float): start hour of work

    day_range_selection (int): upper boundary of range

    days_in_month (int): number of days in month

    previous_day_entered (int) -> 0 : previous entered boundary"""

    if day_range_selection <= previous_day_entered or day_range_selection > days_in_month:
        print("\nSomething went wrong...\nRe-enter day.\n")
    else:
        return work_hour_start_days.update(
            {day_range_selection: work_start})


def handeling_work_hour_days(work_hour_start_days, days_in_month):
    """Handles input of different work start for different days in month

    Parameter:

    work_hour_start_days (dict): dictionary of pairs {day_range_selection : work_start}

    days_in_month (int): number of days in month"""

    while work_hour_start_days == {} or list(work_hour_start_days.keys())[-1] != days_in_month:
        try:
            work_start = float(input('Work starts at: '))
            if 0 < work_start <= 24:
                if len(work_hour_start_days) > 0:
                    previous_day_entered = list(
                        work_hour_start_days.keys())[-1]
                    day_range_selection = int(input(
                        f"Day (included) until which entered 'Work start' is valid ({previous_day_entered+1} - {days_in_month}): "))
                    check_day_range_wrong(
                        work_hour_start_days, work_start, day_range_selection, days_in_month, previous_day_entered)
                else:
                    day_range_selection = int(input(
                        f"Day (included) until which entered 'Work start' is valid (1 - {days_in_month}): "))
                    check_day_range_wrong(
                        work_hour_start_days, work_start, day_range_selection, days_in_month)
            else:
                print(const.incorrect_values_msg)
        except ValueError:
            print(const.incorrect_values_msg)
