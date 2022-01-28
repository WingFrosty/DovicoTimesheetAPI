import sys
import os
import logging
import datetime
import time
#Other modules to import
from modules.utils_retry import retry
import modules.utils_logger as utils_logger
from timesheet import Timesheet

#Program
program_path = sys.argv[0]
program_name = os.path.basename(program_path)
program_name_no_ext = os.path.splitext(program_name)[0]

#Logging
logger = logging.getLogger(__name__)
logging_level = logging.INFO #DEBUG > INFO > WARNING > ERROR > CRITICAL
log_dir_path = "./logs"
log_file_path = log_dir_path + "/" + program_name_no_ext + ".log"
log_file_absolute_path = os.path.abspath(log_file_path)

#Retries
max_tries = 1 #5 #Maximum number os tries a function can be executed when an exception occurs
wait_interval_seconds = 300 #Waiting interval between each retry in seconds (currently: 5 minutes)

@retry(max_tries=max_tries, wait_interval_seconds=wait_interval_seconds)
def start_logger():
    utils_logger.init_file_logger(log_file_path, logging_level)
    logger = logging.getLogger(program_name_no_ext)

    return logger

def fill_workday_cell(date, hours=8):
    timesheet.reset_day(date)
    timesheet.fill_cell(
        date=date,
        client="Luz Saude", #Change Me
        project="GSTEP_2022", #Change Me
        task="Implementation", #Change Me
        hours=hours
    )

def fill_holiday_cell(date, hours=8):
    timesheet.reset_day(date)
    timesheet.fill_cell(
        date=date,
        client="", #Change Me
        project="Leave/Absences", #Change Me
        task="Bank Holliday", #Change Me
        hours=hours
    )

def fill_vacation_cell(date, hours=8):
    timesheet.reset_day(date)
    timesheet.fill_cell(
        date=date,
        client="", #Change Me
        project="Leave/Absences", #Change Me
        task="Vacation", #Change Me
        hours=hours
    )

def fill_training_cell(date, hours=8):
    timesheet.reset_day(date)
    timesheet.fill_cell(
        date=date,
        client="Formacao", #Change Me
        project="Presencial", #Change Me
        task="Training", #Change Me
        hours=hours
    )

def fill_days(start_date, end_date, holidays, vacations):
    current_date = start_date
    for i in range((end_date - start_date).days + 1):

        if (current_date.weekday() in (0, 1, 2, 3, 4)):
            if (current_date in holidays):
                fill_holiday_cell(current_date)
            elif (current_date in vacations):
                fill_vacation_cell(current_date)
            else:
                fill_workday_cell(current_date)

        current_date += datetime.timedelta(days=1)

@retry(max_tries=max_tries, wait_interval_seconds=wait_interval_seconds)
def main():
    global logger
    global timesheet

    # Start Logger
    logger = start_logger()
    logger.info("Starting...")

    url = ""
    username = ""
    password = ""
    headless_browser = True

    holidays = (
        datetime.date(2022, 1, 1),
        datetime.date(2022, 4, 15),
        datetime.date(2022, 4, 17),
        datetime.date(2022, 4, 25),
        datetime.date(2022, 5, 1),
        datetime.date(2022, 6, 10),
        datetime.date(2022, 6, 13),
        datetime.date(2022, 6, 16),
        datetime.date(2022, 8, 15),
        datetime.date(2022, 10, 5),
        datetime.date(2022, 11, 1),
        datetime.date(2022, 12, 1),
        datetime.date(2022, 12, 8),
        datetime.date(2022, 12, 25)
    )

    vacations = (
        datetime.date(2022, 6, 22),
        datetime.date(2022, 6, 23),
        datetime.date(2022, 6, 24),
        datetime.date(2022, 7, 1),
        datetime.date(2022, 8, 8),
        datetime.date(2022, 8, 9),
        datetime.date(2022, 8, 10),
        datetime.date(2022, 8, 11),
        datetime.date(2022, 8, 12),
        datetime.date(2022, 8, 16),
        datetime.date(2022, 8, 17),
        datetime.date(2022, 8, 18),
        datetime.date(2022, 8, 19),
        datetime.date(2022, 10, 3),
        datetime.date(2022, 10, 4),
        datetime.date(2022, 11, 28),
        datetime.date(2022, 11, 29),
        datetime.date(2022, 11, 30),
        datetime.date(2022, 12, 2)
    )

    fill_start_date = datetime.date(2022, 1, 1)
    fill_end_date = datetime.date(2022, 12, 31)

    submit_end_date = datetime.date.today()
    submit_start_date = datetime.date.today() - datetime.timedelta(days=6)

    try:
        timesheet = Timesheet(url, headless_browser)
        timesheet.login(username, password)

        #fill_days(fill_start_date, fill_end_date, holidays, vacations)

        timesheet.submit(submit_start_date, submit_end_date)
    finally:
        if (headless_browser):
            time.sleep(10)
        timesheet.quit()

if __name__ == '__main__':
    main()