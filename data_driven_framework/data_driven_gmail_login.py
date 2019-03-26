import time

import unittest

from selenium import webdriver

import xlrd


class EfffactorLogin(unittest.TestCase):
    def setUp(self):
        self.driver = webdriver.Chrome()

    def test_efffactor_login(self):
        driver = self.driver
        driver.get("https://accounts.google.com")
        driver.maximize_window()
        time.sleep(5)
        self.assertIn("Sign in", driver.title)
        workbook = xlrd.open_workbook("GmailData.xls")
        sheet = workbook.sheet_by_name("GmailCredentials")

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
            driver.find_element_by_id("identifierId").click()
            driver.find_element_by_id("identifierId").clear()
            driver.find_element_by_id("identifierId").send_keys(
                str(result_data[i][0]))
            driver.find_element_by_id("identifierNext").click()
            time.sleep(2)

            driver.find_element_by_name("password").click()            
            driver.find_element_by_name("password").clear()
            driver.find_element_by_name("password").send_keys(
                result_data[i][1])
            driver.find_element_by_id("passwordNext").click()
            time.sleep(2)
            self.assertIn("Account", driver.title)
            time.sleep(2)
            driver.find_element_by_xpath("//*[@id='gb']/div[2]/div[3]/div/div[2]/div/a").click()
            driver.find_element_by_link_text("Sign out").click()
            time.sleep(2)
            driver.find_elements_by_xpath(".//div[@role='link']")[1].click()
            time.sleep(2)

    def tearDown(self):
        self.driver.quit()


if __name__ == "__main__":
    unittest.main()
