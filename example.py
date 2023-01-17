import datetime
import modules.dovico_timesheet as dovico_timesheet
from modules.dovico_timesheet import DovicoTimesheet

#Timesheet
browser_path = "C:\\Program Files\\Mozilla Firefox\\firefox.exe"  # Change Me
webdriver_path = "webdrivers\\geckodriver.exe"  # Change Me
webdriver_log = "logs\\geckodriver.log"  # Change Me
url = "http://timesheet/DovTimesheet/login.aspx"  # Change Me
username = "username"  # Change Me
password = "password"  # Change Me
headless_browser = False  # Change Me

def main():
    with DovicoTimesheet(url=url,
                         browser_type=dovico_timesheet.FIREFOX,
                         browser_path=browser_path,
                         webdriver_path=webdriver_path,
                         webdriver_log=webdriver_log,
                         headless_browser=headless_browser
                         ) as timesheet:
        timesheet.login(username, password)

        day = datetime.date(2023, 4, 3)

        timesheet.fill_cell(
            date=day,
            client="",  # Change Me
            project="Leave/Absences",  # Change Me
            task="Bank Holliday",  # Change Me
            hours=8
        )

        timesheet.reset_day(day)

        timesheet.fill_cell(
            date=day,
            client="Microsoft",  # Change Me
            project="2023",  # Change Me
            task="Implementation",  # Change Me
            hours=8
        )

        timesheet.submit(day, day)

if __name__ == '__main__':
    main()