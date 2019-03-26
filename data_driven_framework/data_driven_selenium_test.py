import time

import unittest

from selenium import webdriver

import xlrd


class EfffactorLogin(unittest.TestCase):
    def setUp(self):
        self.driver = webdriver.Firefox()

    def test_efffactor_login(self):
        driver = self.driver
        driver.get("http://localhost:8000/login/?next=/")
        driver.maximize_window()
        time.sleep(5)
        workbook = xlrd.open_workbook("DataFile.xls")
        sheet = workbook.sheet_by_name("UserCredentials")

        rowcount = sheet.nrows  # Get number of rows with data in excel sheet
        # Get number of columns with data in each row. Returns highest number
        colcount = sheet.ncols

        result_data = []
        for curr_row in range(1, rowcount):
            row_data = []

            for curr_col in range(1, colcount):
                # Read the data in the current cell
                data = sheet.cell_value(curr_row, curr_col)
                row_data.append(data)
            print "row data", row_data

            result_data.append(row_data)
        print "result_data", result_data

        for i in range(0, rowcount - 1):
            driver.find_element_by_id("id_username").click()
            driver.find_element_by_id("id_username").clear()
            driver.find_element_by_id("id_username").send_keys(
                str(result_data[i][0]))
            driver.find_element_by_id("id_password").click()
            driver.find_element_by_id("id_password").clear()
            driver.find_element_by_id("id_password").send_keys(
                int(result_data[i][1]))
            driver.find_element_by_xpath(
                "/html/body/div[3]/div/div/form/input[2]").click()
            self.assertEqual(":: mpower app", driver.title)
            self.assertEqual(str(result_data[i][2]),
                             driver.find_element_by_xpath(
                "/html/body/div[1]/div/div[2]/ul[2]/li[1]/a/strong").text)
            time.sleep(2)
            driver.find_element_by_link_text("Logout").click()
            time.sleep(2)

    def tearDown(self):
        self.driver.quit()


if __name__ == "__main__":
    unittest.main()
