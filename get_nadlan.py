import os
import re
from datetime import datetime
from time import sleep

from chrome_handler import ChromeHandler
from gov_excel_handler import ExcelHandler

# Constants
NADLAN_START_URL = r"https://nadlan.taxes.gov.il/svinfonadlan2010/startpageNadlanNewDesign.aspx?ProcessKey=06b4de83-6094-4401-9bc8-e980a0577cc6"
RESULT_PAGE_STR = "InfoNadlanPerut"
EXCEL_RESULT_FOLDER_NAME = "Excel_Result_"
DATE_TIME_STR = "%d-%m-%y_%H-%M-%S"
DEFAULT_EXCEL_FILE_NAME = "results.xlsx"
TABLE_START_ROW = 1
TABLE_START_COL = 1
NUM_OF_CELLS_TO_EXTRACT = 20
STARS_STR = "*" * 30 + "\n"
RUN_AGAIN_STR = STARS_STR + "Run Again - Enter\n" + "Close - D+Enter\n"
END_MSG = "Press enter to close\n"
START_RUN_MSG = "Going to MEIDA NADLAN\nPlease enter your search"
MAX_NUM_OF_DEALS_IN_PAGE = 12
MAIN_PAGE_PIC_NAME = "PAGE_{0}.png"
DETAIL_PAGE_PIC_NAME = "{0}.png"
MAIN_PAGE_TABLE_DATA_COLS = 10
MAIN_PAGE_START_ROW = 1
MAIN_PAGE_START_CELL = 1


# Field_IDs
NUM_OF_DEALS_ID = "lblresh"
NUM_OF_PAGES_ID = "lblPage"
TABLE_ENTRY_LINK_ID = "ContentUsersPage_GridMultiD1_LogShow_{0}"
TABLE_ENTRY_NO_LINK_ID = "ContentUsersPage_GridMultiD1_LogNotExist_{0}"
TABLE_ID = "ContentUsersPage_GridMultiD1"

NAMES_TO_TAKE_FROM_MANE_TABLE = []

NAME_TO_ID = {
    "גוש חלקה": "ContentUsersPage_lblGush",
    "יום מכירה": "ContentUsersPage_lblTarIska",
    "תמורה מוצהרת בש\"ח": "ContentUsersPage_lblMcirMozhar",
    "שווי מכירה בש\"ח": "ContentUsersPage_lblMcirMorach",
    "מהות": "ContentUsersPage_lblTifkudYchida",
    "חלק נמכר": "ContentUsersPage_lblShumaHalakim",
    "ישוב": "ContentUsersPage_lblYeshuv",
    "שנת בניה": "ContentUsersPage_lblShnatBniya",
    "שטח": "ContentUsersPage_lblShetachBruto",
    "חדרים": "ContentUsersPage_lblMisHadarim",
    "שטח רשום": "ContentUsersPage_lblShetachNeto",
    "רחוב": "ContentUsersPage_lblRechov",
    "מספר בית": "ContentUsersPage_lblBayit",
    "מספר דירה": "ContentUsersPage_lblDira",
    "קומה": "ContentUsersPage_lblKoma",
    "מספר קומות": "ContentUsersPage_lblMisKomot",
    "דירות בבניין": "ContentUsersPage_lblDirotBnyn",
    "גינה": "ContentUsersPage_lblHzer",
    "גג": "ContentUsersPage_lblGag",
    "מחסן": "ContentUsersPage_lblMachsan"
}

NAME_TO_MAIN_TABLE_COL = {
    "גוש חלקה": 0,
    "יום מכירה": 1,
    "תמורה מוצהרת בש\"ח": 2,
    "שווי מכירה בש\"ח": 3,
    "מהות": 4,
    "חלק נמכר": 5,
    "ישוב": 6,
    "שנת בניה": 7,
    "שטח": 8,
    "חדרים": 9
}

# Patterns
GET_DIGITS_PAT = r"\d+"


class GovGet(object):
    def __init__(self):
        self.download_dir = None
        self.excel_file_path = None
        self.excel_handler = ExcelHandler(self.excel_file_path, table_row=TABLE_START_ROW, table_col=TABLE_START_COL)
        self.detail_page_field_names = self.excel_handler.get_table_head(NUM_OF_CELLS_TO_EXTRACT)
        self.field_ids = [NAME_TO_ID[field_name] for field_name in self.detail_page_field_names]
        self.chrome = ChromeHandler()

    def init_dir_and_excel(self):
        self.download_dir = self.create_new_dir()
        self.excel_file_path = os.path.join(self.download_dir, DEFAULT_EXCEL_FILE_NAME)
        self.excel_handler = ExcelHandler(self.excel_file_path, table_row=TABLE_START_ROW, table_col=TABLE_START_COL)

    def get_excel_from_gov(self):
        self.chrome.go_to_url(NADLAN_START_URL)
        self.chrome.wait_till_in_url(RESULT_PAGE_STR)
        current_page, num_of_pages = self.get_page_info()
        num_of_deals = self.get_total_num_of_deals()
        for page_num in range(1, num_of_pages + 1):
            if page_num != 1:
                self.go_to_page(page_num)
            num_table_rows = self.get_num_table_rows_in_page(num_of_pages, num_of_deals, page_num)
            main_page_table_data = self.chrome.get_table_data(TABLE_ID, MAIN_PAGE_START_ROW, MAIN_PAGE_START_CELL,
                                                              num_table_rows, MAX_NUM_OF_DEALS_IN_PAGE)
            for table_row in range(num_table_rows):
                deal_num, has_detailed_page = self.get_deal_info(table_row)
                if has_detailed_page:
                    self.get_data_from_detailed_page_and_get_screenshot(table_row, main_page_table_data)
                else:
                    self.get_data_from_general_page(table_row, main_page_table_data)
            self.chrome.save_screenshot_to_dir(self.download_dir, MAIN_PAGE_PIC_NAME.format(page_num))
        self.excel_handler.save()

    @staticmethod
    def create_new_dir():
        new_dir_name = EXCEL_RESULT_FOLDER_NAME + datetime.now().strftime(DATE_TIME_STR)
        os.mkdir(new_dir_name)
        return os.path.join(os.getcwd(), new_dir_name)

    def get_total_num_of_deals(self):
        total_num_of_deals_info = self.chrome.get_text_by_filed_id(NUM_OF_DEALS_ID)
        ret = re.findall(GET_DIGITS_PAT, total_num_of_deals_info)
        if len(ret) != 1:
            raise EnvironmentError("Could not parse page info string:{0}".format(total_num_of_deals_info))
        total_num_of_deals_str = ret[0]
        return int(total_num_of_deals_str)

    def get_data_from_general_page(self, table_row, table_data):
        self.excel_handler.insert_list_to_row(data_list=table_data[table_row])

    def get_data_from_detailed_page_and_get_screenshot(self, table_row, main_page_table_data):
        self.chrome.click_element_by_id(TABLE_ENTRY_LINK_ID.format(table_row))
        detailed_page_data_list = self.chrome.get_data_from_id_list(self.field_ids)
        data_list = main_page_table_data[table_row] + detailed_page_data_list[len(main_page_table_data[table_row]):]
        self.excel_handler.insert_list_to_row(data_list=data_list)
        self.chrome.save_screenshot_to_dir(self.download_dir, DETAIL_PAGE_PIC_NAME.format(data_list[0]))
        self.chrome.go_to_prev_page()
        sleep(0.5)

    def get_page_info(self):
        page_info = self.chrome.get_text_by_filed_id(NUM_OF_PAGES_ID)
        ret = re.findall(GET_DIGITS_PAT, page_info)
        if len(ret) != 2:
            raise EnvironmentError("Could not parse page info string:{0}".format(page_info))
        current_page_str, num_of_pages_str = ret[0], ret[1]
        return int(current_page_str), int(num_of_pages_str)

    def go_to_page(self, page_num):
        self.chrome.click_element_by_text(str(page_num))

    def close_chrome(self):
        self.chrome.close()

    @staticmethod
    def get_num_table_rows_in_page(num_of_pages, num_of_deals, page_num):
        if (MAX_NUM_OF_DEALS_IN_PAGE * num_of_pages) < num_of_deals:
            raise EnvironmentError("Num of pages does not match expected. deals:{0}, pages:{1}".format(num_of_deals,
                                                                                                       num_of_pages))
        if page_num < num_of_pages:
            return MAX_NUM_OF_DEALS_IN_PAGE
        else:
            return num_of_deals - (MAX_NUM_OF_DEALS_IN_PAGE * (num_of_pages - 1))

    def get_deal_info(self, table_row):
        deal_num = self.chrome.get_text_by_filed_id(TABLE_ENTRY_LINK_ID.format(table_row))
        if deal_num is not None:
            return deal_num, True
        deal_num = self.chrome.get_text_by_filed_id(TABLE_ENTRY_NO_LINK_ID.format(table_row))
        if deal_num is not None:
            return deal_num, False
        raise EnvironmentError("No such table row entry exist, table_row: {0}".format(table_row))


if __name__ == '__main__':
    gov_get = GovGet()
    keep_running = True
    while keep_running:
        print(START_RUN_MSG)
        gov_get.init_dir_and_excel()
        gov_get.get_excel_from_gov()
        user_input = input(RUN_AGAIN_STR)
        keep_running = user_input == ''
    gov_get.close_chrome()
