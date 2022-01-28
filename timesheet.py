import logging
import time
from selenium.webdriver import Firefox
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.common.keys import Keys
#Other modules to import
from modules.utils_retry import retry

logger = logging.getLogger(__name__)

#Retries
max_tries = 10 #10 #Maximum number os tries a function can be executed when an exception occurs
wait_interval_seconds = 1 #Waiting interval between each retry in seconds (currently: 1 second)
default_timeout = 2 #timeout for elements
sleep_time = 0.1 #time.sleep

class Timesheet:
    def __init__(self, url, headless_browser=False):
        self.__current_frame = ""

        self.url = url
        self.opts = Options()

        if (headless_browser):
            self.opts.headless = True

        self.__driver = Firefox(
            executable_path='./webdrivers/geckodriver.exe',
            options=self.opts,
            service_log_path='./logs/geckodriver.log'
        )

        self.__navigate_to_website()

    def __month_number_to_text(self, month_number):
        result = ""
        if (month_number == 1):
            result = "January"
        elif (month_number == 2):
            result = "February"
        elif (month_number == 3):
            result = "March"
        elif (month_number == 4):
            result = "April"
        elif (month_number == 5):
            result = "May"
        elif (month_number == 6):
            result = "June"
        elif (month_number == 7):
            result = "July"
        elif (month_number == 8):
            result = "August"
        elif (month_number == 9):
            result = "September"
        elif (month_number == 10):
            result = "October"
        elif (month_number == 11):
            result = "November"
        elif (month_number == 12):
            result = "December"
        return result

    def __month_text_to_number(self, month_text):
        result = 0
        if (month_text == "January"):
            result = 1
        elif (month_text == "February"):
            result = 2
        elif (month_text == "March"):
            result = 3
        elif (month_text == "April"):
            result = 4
        elif (month_text == "May"):
            result = 5
        elif (month_text == "June"):
            result = 6
        elif (month_text == "July"):
            result = 7
        elif (month_text == "August"):
            result = 8
        elif (month_text == "September"):
            result = 9
        elif (month_text == "October"):
            result = 10
        elif (month_text == "November"):
            result = 11
        elif (month_text == "December"):
            result = 12
        return result

    @retry(max_tries=max_tries, wait_interval_seconds=wait_interval_seconds)
    def __navigate_to_website(self):
        self.__driver.get(self.url)
        logger.info("Loaded Timesheet's Webpage.")

    @retry(max_tries=max_tries, wait_interval_seconds=wait_interval_seconds)
    def __switch_to_frame(self, frame):
        if (self.__current_frame != frame):
            self.__driver.switch_to.default_content()
            if (frame == "tetTsFrame_FramesetGrid"):
                self.__driver.switch_to.frame("tetFrame_tet_ts_frame")
            elif (frame == "tetTsFrame_tet_ts_grid_columns"):
                self.__driver.switch_to.frame("tetFrame_tet_ts_frame")
                self.__driver.switch_to.frame("tetTsFrame_FramesetGrid")
            elif (frame == "tetTsFrame_tet_ts_grid_timesheet"):
                self.__driver.switch_to.frame("tetFrame_tet_ts_frame")
                self.__driver.switch_to.frame("tetTsFrame_FramesetGrid")
            elif (frame == "tetTsFrame_tet_ts_frame_assignmenttree"):
                self.__driver.switch_to.frame("tetFrame_tet_ts_frame")
            elif (frame == "tetTsFrameAssignmentTree_tet_ts_form_assignment_tree"):
                self.__driver.switch_to.frame("tetFrame_tet_ts_frame")
                self.__driver.switch_to.frame("tetTsFrame_tet_ts_frame_assignmenttree")
            elif (frame == "tetTsFrameAssignmentTree_tet_ts_tree_assignmnettree"):
                self.__driver.switch_to.frame("tetFrame_tet_ts_frame")
                self.__driver.switch_to.frame("tetTsFrame_tet_ts_frame_assignmenttree")
            self.__driver.switch_to.frame(frame)
            self.__current_frame = frame

    @retry(max_tries=max_tries, wait_interval_seconds=wait_interval_seconds)
    def __leave_frames(self):
        self.__driver.switch_to.default_content()
        self.__current_frame = ""

    @retry(max_tries=max_tries, wait_interval_seconds=wait_interval_seconds)
    def __exists_element_by_css(self, css_selector):
        try:
            self.__driver.find_element(By.CSS_SELECTOR, css_selector)
        except NoSuchElementException:
            return False
        return True

    @retry(max_tries=max_tries, wait_interval_seconds=wait_interval_seconds)
    def __click_calendar(self):
        self.__switch_to_frame("tetTsFrame_tet_ts_grid_columns")
        timesheet_calendar = WebDriverWait(self.__driver, timeout=default_timeout).until(
            EC.element_to_be_clickable((By.ID, "imgTetGridTimesheetCalendar")))
        ActionChains(self.__driver).move_to_element(timesheet_calendar).click().perform()
        time.sleep(sleep_time)

    @retry(max_tries=max_tries, wait_interval_seconds=wait_interval_seconds)
    def __get_calendar_year_dropdown_select(self):
        self.__switch_to_frame("tetTsFrame_tet_ts_grid_timesheet")
        active_year = WebDriverWait(self.__driver, timeout=default_timeout).until(
            lambda d: d.find_element(By.CSS_SELECTOR, "#divDropDownCalendar table tr.CALENDAR_HEADER td:nth-child(2) select")
        )
        active_year_select = Select(active_year)
        return active_year_select

    @retry(max_tries=max_tries, wait_interval_seconds=wait_interval_seconds)
    def __get_calendar_month_dropdown_select(self):
        self.__switch_to_frame("tetTsFrame_tet_ts_grid_timesheet")
        active_month = WebDriverWait(self.__driver, timeout=default_timeout).until(
            lambda d: d.find_element(By.CSS_SELECTOR, "#divDropDownCalendar table tr.CALENDAR_HEADER td:nth-child(1) select")
        )
        active_month_select = Select(active_month)
        return active_month_select

    @retry(max_tries=max_tries, wait_interval_seconds=wait_interval_seconds)
    def __set_calendar_year_month_dropdown(self, date):
        self.__switch_to_frame("tetTsFrame_tet_ts_grid_timesheet")
        year = str(date.year)
        month = self.__month_number_to_text(date.month)

        year_dropdown_select = self.__get_calendar_year_dropdown_select()
        current_year = year_dropdown_select.first_selected_option.text
        logger.info("Currently selected year is " + current_year)
        if (current_year != year):
            year_dropdown_select.select_by_visible_text(year)
            logger.info("Year changed to " + year)

        month_dropdown_select = self.__get_calendar_month_dropdown_select()
        current_month = month_dropdown_select.first_selected_option.text
        logger.info("Currently selected month is " + current_month)
        if (current_month != month):
            month_dropdown_select.select_by_visible_text(month)
            logger.info("Month changed to " + month)

    @retry(max_tries=max_tries, wait_interval_seconds=wait_interval_seconds)
    def __select_calendar_week(self, date):
        self.__switch_to_frame("tetTsFrame_tet_ts_grid_timesheet")
        day = str(date.day)
        day_element = ""

        calendar_table = WebDriverWait(self.__driver, timeout=default_timeout).until(
            lambda d: d.find_element(By.CSS_SELECTOR, "#divDropDownCalendar table.CALENDAR_TABLE tbody")
        )
        for row in calendar_table.find_elements(By.CSS_SELECTOR, "tr"):
            i = 0
            for cell in row.find_elements(By.CSS_SELECTOR, "td"):
                if (cell.text == day and cell.get_attribute("class") == "CALENDAR_DAYS"):
                    day_element = cell
                    day_week_index = i
                i = i + 1
        ActionChains(self.__driver).move_to_element(day_element).click().perform()
        time.sleep(sleep_time)
        logger.info("Week has been changed.")

        return day_week_index

    @retry(max_tries=max_tries, wait_interval_seconds=wait_interval_seconds)
    def __change_day(self, date):
        self.__click_calendar()
        self.__set_calendar_year_month_dropdown(date)
        week_index = self.__select_calendar_week(date)
        return week_index

    @retry(max_tries=max_tries, wait_interval_seconds=wait_interval_seconds)
    def __get_timesheet_grid(self):
        self.__switch_to_frame("tetTsFrame_tet_ts_grid_timesheet")
        grid = WebDriverWait(self.__driver, timeout=default_timeout).until(
            lambda d: d.find_element(By.CSS_SELECTOR, "#csgTimesheetGrid table.tableGrid tbody")
        )
        return grid

    @retry(max_tries=max_tries, wait_interval_seconds=wait_interval_seconds)
    def __reset_grid_day(self, grid, grid_day_index):
        self.__switch_to_frame("tetTsFrame_tet_ts_grid_timesheet")

        for row in grid.find_elements(By.CSS_SELECTOR, "tr"):
            if (row.get_attribute("class") == "ROW_NORMAL" or row.get_attribute("class") == "ROW_ALTERNATE"):
                i = 0
                for cell in row.find_elements(By.CSS_SELECTOR, "td"):
                    for el in cell.find_elements(By.CSS_SELECTOR, "div > div"):  # Deverá conter apenas 1 elemento no máximo
                        if (el.get_attribute("class") == "NOT_READONLY" and i == grid_day_index):
                            data_cell = el.find_element(By.CSS_SELECTOR, "table.tblDataCell tbody tr td.tdDataCell div.divDataCell")
                            data_cell.click()
                            time.sleep(sleep_time)
                            ActionChains(self.__driver).send_keys(Keys.BACKSPACE).perform()
                            time.sleep(sleep_time)
                        i = i + 1

    @retry(max_tries=max_tries, wait_interval_seconds=wait_interval_seconds)
    def __add_task(self, client, project, task):
        self.__switch_to_frame("tetTsFrameAssignmentTree_tet_ts_tree_assignmnettree")

        level0_css_selector = "div#AdivNode--1 div#AdivNode--2 > div > div"
        # Get Task Tree Element
        task_tree = WebDriverWait(self.__driver, timeout=default_timeout).until(
            lambda d: d.find_element(By.CSS_SELECTOR, "td#tdTreeUserTimeTreeGoesHere"))
        # Get all the first nodes of task tree
        # nodes = task_tree.find_elements(By.CSS_SELECTOR, level0_css_selector)
        nodes = WebDriverWait(self.__driver, timeout=default_timeout).until(
            lambda d: d.find_elements(By.CSS_SELECTOR, "td#tdTreeUserTimeTreeGoesHere " + level0_css_selector))
        for i in range(len(nodes)):
            # Get name of client
            name = task_tree.find_element(By.CSS_SELECTOR, level0_css_selector + ":nth-child(" + str(
                i + 1) + ") > table.RowLines table > tbody > tr > td:nth-child(3)")
            name_text = name.text.strip()
            # If client is the same (or project if there is no client)
            if ((len(client) > 0 and client == name_text) or (len(client) == 0 and project == name_text)):
                # Expand node
                # expand_btn = task_tree.find_element(By.CSS_SELECTOR, level0_css_selector + ":nth-child(" + str(i+1) + ") > table.RowLines tr > td:nth-child(2) > div")
                expand_btn = WebDriverWait(self.__driver, timeout=default_timeout).until(
                    lambda d: d.find_element(By.CSS_SELECTOR, "td#tdTreeUserTimeTreeGoesHere " + level0_css_selector + ":nth-child(" + str(i + 1) + ") > table.RowLines tr > td:nth-child(2) > div")
                )
                # expand_btn = WebDriverWait(self.__driver, timeout = wait_interval_seconds).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "td#tdTreeUserTimeTreeGoesHere " + level0_css_selector + ":nth-child(" + str(i+1) + ") > table.RowLines tr > td:nth-child(2) > div")))
                expand_btn.click()
                time.sleep(sleep_time)
                # Get all child nodes
                level1_css_selector = level0_css_selector + ":nth-child(" + str(i + 1) + ") > div > div"
                # nodes2 = task_tree.find_elements(By.CSS_SELECTOR, level1_css_selector)
                nodes2 = WebDriverWait(self.__driver, timeout=default_timeout).until(
                    lambda d: d.find_elements(By.CSS_SELECTOR, "td#tdTreeUserTimeTreeGoesHere " + level1_css_selector))
                for j in range(len(nodes2)):
                    # Get name of project
                    name = task_tree.find_element(By.CSS_SELECTOR, level1_css_selector + ":nth-child(" + str(
                        j + 1) + ") > table.RowLines table > tbody > tr > td:nth-child(3)")
                    name_text = name.text.strip()
                    # If project is the same (or task if there is no client)
                    if ((len(client) > 0 and project == name_text) or (len(client) == 0 and task == name_text)):
                        # Expand node if there is a client
                        if (len(client) == 0):
                            # checkbox = task_tree.find_element(By.CSS_SELECTOR, level1_css_selector + ":nth-child(" + str(j+1) + ") > table.RowLines table > tbody > tr > td:nth-child(1) > input")
                            checkbox = WebDriverWait(self.__driver, timeout=default_timeout).until(
                                EC.element_to_be_clickable((By.CSS_SELECTOR,
                                                            "td#tdTreeUserTimeTreeGoesHere " + level1_css_selector + ":nth-child(" + str(
                                                                j + 1) + ") > table.RowLines table > tbody > tr > td:nth-child(1) > input")))
                            checkbox.click()
                            time.sleep(sleep_time)
                        else:
                            # expand_btn = task_tree.find_element(By.CSS_SELECTOR, level1_css_selector + ":nth-child(" + str(j+1) + ") > table.RowLines tr > td:nth-child(2) > div")
                            expand_btn = WebDriverWait(self.__driver, timeout=default_timeout).until(
                                EC.element_to_be_clickable((By.CSS_SELECTOR,
                                                            "td#tdTreeUserTimeTreeGoesHere " + level1_css_selector + ":nth-child(" + str(
                                                                j + 1) + ") > table.RowLines tr > td:nth-child(2) > div")))
                            expand_btn.click()
                            time.sleep(sleep_time)
                            # Get all child nodes
                            level2_css_selector = level1_css_selector + ":nth-child(" + str(j + 1) + ") > div > div"
                            # nodes3 = task_tree.find_elements(By.CSS_SELECTOR, level2_css_selector)
                            nodes3 = WebDriverWait(self.__driver, timeout=default_timeout).until(
                                lambda d: d.find_elements(By.CSS_SELECTOR, "td#tdTreeUserTimeTreeGoesHere " + level2_css_selector)
                            )
                            for k in range(len(nodes3)):
                                # Get name of task
                                name = task_tree.find_element(By.CSS_SELECTOR, level2_css_selector + ":nth-child(" + str(k + 1) + ") > table.RowLines > tbody > tr > td:nth-child(3) >div > table > tbody > tr > td:nth-child(3)")
                                name_text = name.text.strip()
                                # If task is the same
                                if (task == name_text):
                                    # checkbox = task_tree.find_element(By.CSS_SELECTOR, level2_css_selector + ":nth-child(" + str(k+1) + ") > table.RowLines > tbody > tr > td:nth-child(3) >div > table > tbody > tr > td:nth-child(1) > input")
                                    checkbox = WebDriverWait(self.__driver, timeout=default_timeout).until(
                                        EC.element_to_be_clickable((By.CSS_SELECTOR,
                                                                    "td#tdTreeUserTimeTreeGoesHere " + level2_css_selector + ":nth-child(" + str(
                                                                        k + 1) + ") > table.RowLines > tbody > tr > td:nth-child(3) >div > table > tbody > tr > td:nth-child(1) > input")))
                                    checkbox.click()
                                    time.sleep(sleep_time)
        # Add to timesheet
        self.__switch_to_frame("tetTsFrameAssignmentTree_tet_ts_form_assignment_tree")
        btn = WebDriverWait(self.__driver, timeout=default_timeout).until(
            lambda d: d.find_element(By.CSS_SELECTOR, "#divFrmTreeUserTimeAddGrid"))
        btn.click()
        time.sleep(sleep_time)

    @retry(max_tries=max_tries, wait_interval_seconds=wait_interval_seconds)
    def __add_hours(self, grid, grid_day_index, client, project, task, hours):
        self.__switch_to_frame("tetTsFrame_tet_ts_grid_timesheet")

        if (len(client) == 0):
            group = project
        else:
            group = client + " - " + project
        current_group = ""
        current_task = ""

        for row in grid.find_elements(By.CSS_SELECTOR, "tr"):
            row_class = row.get_attribute("class")
            if (row_class == "ROW_GROUP"):
                group_name = row.find_element(By.CSS_SELECTOR, "td:nth-child(1) > div:nth-child(1) > table:nth-child(1) > tbody:nth-child(1) > tr:nth-child(1) > td:nth-child(1) > div:nth-child(1) > b:nth-child(1)")
                current_group = group_name.text
                logger.info("GROUP = " + current_group)
                continue

            if ((row_class == "ROW_NORMAL" or row_class == "ROW_ALTERNATE") and current_group == group):
                i = 0
                for cell in row.find_elements(By.CSS_SELECTOR, "td"):
                    if (cell.get_attribute("class") == "PROJECT_NAMES"):
                        task_name = cell.find_element(By.CSS_SELECTOR, "div:nth-child(1) > table:nth-child(1) > tbody:nth-child(1) > tr:nth-child(1) > td:nth-child(1) > div:nth-child(2)")
                        current_task = task_name.text
                        continue

                    if (current_task == task):
                        for el in cell.find_elements(By.CSS_SELECTOR, "div > div"):  # Deverá conter apenas 1 elemento no máximo
                            if (el.get_attribute("class") == "NOT_READONLY" and i == grid_day_index):
                                data_cell = el.find_element(By.CSS_SELECTOR, "table.tblDataCell tbody tr td.tdDataCell div.divDataCell")
                                data_cell.click()
                                time.sleep(sleep_time)
                                ActionChains(self.__driver).send_keys(hours, Keys.ENTER).perform()
                                time.sleep(sleep_time)
                            i = i + 1

    @retry(max_tries=max_tries, wait_interval_seconds=wait_interval_seconds)
    def __launch_submit_window(self):
        self.__switch_to_frame("tetToolbar")
        btn = WebDriverWait(self.__driver, timeout=default_timeout).until(
            EC.element_to_be_clickable((By.ID, "tbiSubmit")))
        btn.click()
        time.sleep(sleep_time)

    @retry(max_tries=max_tries, wait_interval_seconds=wait_interval_seconds)
    def __change_to_submit_window(self, original_window):
        try:
            wait = WebDriverWait(self.__driver, timeout=default_timeout)
            wait.until(EC.number_of_windows_to_be(2))

            for window_handle in self.__driver.window_handles:
                if window_handle != original_window:
                    self.__driver.switch_to.window(window_handle)
                    break

            wait.until(EC.title_is("Submit"))

        except TimeoutException:
            self.__launch_submit_window()
            raise

    @retry(max_tries=max_tries, wait_interval_seconds=wait_interval_seconds)
    def __change_submit_dates(self, start_date, end_date):
        start_date_formatted = str(start_date.month) + "/" + str(start_date.day) + "/" + str(start_date.year)
        end_date_formatted = str(end_date.month) + "/" + str(end_date.day) + "/" + str(end_date.year)

        wait = WebDriverWait(self.__driver, timeout=default_timeout)
        from_date = wait.until(EC.element_to_be_clickable((By.ID, "dpTetSubmitFromDate_dpTetSubmitFromDate_Date")))
        to_date = wait.until(EC.element_to_be_clickable((By.ID, "dpTetSubmitToDate_dpTetSubmitToDate_Date")))

        from_date.clear()
        from_date.send_keys(start_date_formatted)
        to_date.clear()
        to_date.send_keys(end_date_formatted)

        refresh = wait.until(EC.element_to_be_clickable((By.ID, "cmdTetSubmitRefresh")))
        refresh.click()
        time.sleep(sleep_time)

    @retry(max_tries=max_tries, wait_interval_seconds=wait_interval_seconds)
    def __click_submit(self):
        wait = WebDriverWait(self.__driver, timeout=default_timeout)

        submit = wait.until(EC.element_to_be_clickable((By.ID, "cmdTetSubmitSubmit")))
        submit.click()
        time.sleep(sleep_time)

    @retry(max_tries=max_tries, wait_interval_seconds=wait_interval_seconds)
    def login(self, username, password):
        username_box = self.__driver.find_element(By.ID, "txtLogin")
        username_box.clear()
        username_box.send_keys(username)

        password_box = self.__driver.find_element(By.ID, "txtPassword")
        password_box.clear()
        password_box.send_keys(password)

        login_button = self.__driver.find_element(By.ID, "cmdLogin")
        login_button.click()
        time.sleep(sleep_time)

        logger.info("Log in attempt sent.")

    @retry(max_tries=max_tries, wait_interval_seconds=wait_interval_seconds)
    def reset_day(self, date):
        #time.sleep(sleep_time)
        grid_day_index = self.__change_day(date)
        #time.sleep(sleep_time)
        grid = self.__get_timesheet_grid()
        self.__reset_grid_day(grid, grid_day_index)

        logger.info("Day reset.")

    @retry(max_tries=max_tries, wait_interval_seconds=wait_interval_seconds)
    def fill_cell(self, date, client, project, task, hours):
        #time.sleep(sleep_time)
        grid_day_index = self.__change_day(date)
        #time.sleep(sleep_time)
        self.__add_task(client, project, task)
        #time.sleep(sleep_time)
        grid = self.__get_timesheet_grid()
        self.__add_hours(grid, grid_day_index, client, project, task, hours)

        logger.info("Hours added.")

    @retry(max_tries=max_tries, wait_interval_seconds=wait_interval_seconds)
    def submit(self, start_date, end_date):
        #time.sleep(sleep_time)
        original_window = self.__driver.current_window_handle
        self.__launch_submit_window()
        #time.sleep(sleep_time)
        self.__change_to_submit_window(original_window)
        self.__change_submit_dates(start_date, end_date)
        #time.sleep(sleep_time)
        self.__click_submit()

        logger.info("Hours submited")

    @retry(max_tries=max_tries, wait_interval_seconds=wait_interval_seconds)
    def quit(self):
        self.__driver.quit()
        logger.info("Browser closed.")