import json

import time

import unittest

from selenium import webdriver


class EfffactorLogin(unittest.TestCase):
    def setUp(self):
        self.driver = webdriver.Firefox()

    def test_efffactor_login(self):
        driver = self.driver
        driver.get("http://localhost:8000/login/?next=/")
        driver.maximize_window()
        time.sleep(5)
        with open('user_info.json', 'r') as f:
            user_dict = json.load(f)
            print len(user_dict)

            for i in range(0, len(user_dict)):
                print(user_dict[i]['username'])
                print(user_dict[i]['password'])
                print(user_dict[i]['name'])
                driver.find_element_by_id("id_username").click()
                driver.find_element_by_id("id_username").clear()
                driver.find_element_by_id("id_username").send_keys(
                    user_dict[i]['username'])
                driver.find_element_by_id("id_password").click()
                driver.find_element_by_id("id_password").clear()
                driver.find_element_by_id("id_password").send_keys(
                    user_dict[i]['password'])
                driver.find_element_by_xpath(
                    "/html/body/div[3]/div/div/form/input[2]").click()
                self.assertEqual(":: mpower app", driver.title)
                self.assertEqual(user_dict[i]['name'],
                                 driver.find_element_by_xpath(
                    "/html/body/div[1]/div/div[2]/ul[2]/li[1]/a/strong").text)
                time.sleep(2)
                driver.find_element_by_link_text("Logout").click()
                time.sleep(2)

    def tearDown(self):
        self.driver.quit()


if __name__ == "__main__":
    unittest.main()
