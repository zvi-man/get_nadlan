import os
from time import sleep
from selenium import webdriver

# Constants
CHROME_DRIVER = "chromedriver.exe"
CHROME_PATH = os.path.join(os.getcwd(), CHROME_DRIVER)


class GovChromeHandlerError(Exception):
    pass


class ChromeHandler(object):
    def __init__(self):
        self.driver = webdriver.Chrome(CHROME_PATH)
        self.driver.maximize_window()

    def go_to_url(self, url):
        self.driver.get(url)

    def wait_till_in_url(self, url_str):
        while url_str not in self.driver.current_url:
            sleep(1)

    def get_value_of_id(self, element_id):
        self.driver.find_element_by_id(element_id)

    def close(self):
        self.driver.close()

    def get_elem_by_id(self, field_id):
        elements = self.driver.find_elements_by_id(field_id)
        if len(elements) != 1:
            raise GovChromeHandlerError("Found more or less then one element with the same id")
        return elements[0]

    def get_text_by_filed_id(self, field_id):
        elements = self.driver.find_elements_by_id(field_id)
        if len(elements) > 1:
            raise GovChromeHandlerError("Found more then one element with the same id: {0}".format(field_id))
        if len(elements) == 0:
            return None
        return elements[0].text

    def click_element_by_id(self, field_id):
        self.get_elem_by_id(field_id).click()
        self.driver.refresh()

    def click_element_by_text(self, field_text):
        self.driver.find_element_by_link_text(field_text).click()
        self.driver.refresh()

    def get_data_from_id_list(self, field_id_list):
        data_list = []
        for field_id in field_id_list:
            data_list.append(self.get_text_by_filed_id(field_id))
        return data_list

    def save_screenshot_to_dir(self, dir_path, pic_name):
        pic_name_full_path = os.path.join(dir_path, pic_name)
        self.driver.save_screenshot(pic_name_full_path)

    def get_table_data(self, table_id, start_row, start_cell, num_rows, num_cells):
        element = self.driver.find_element_by_id(table_id)
        rows = element.find_elements_by_tag_name("tr")
        table_data = []
        for row in rows[start_row:start_row + num_rows]:
            row_data = []
            cols = row.find_elements_by_tag_name("td")
            for col in cols[start_cell: start_cell + num_cells]:
                row_data.append(col.text)
            table_data.append(row_data)
        return table_data

    def go_to_prev_page(self):
        self.driver.back()
        self.driver.refresh()
