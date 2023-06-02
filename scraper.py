from genericpath import exists
from types import NoneType
from selenium import webdriver
from selenium.webdriver.common.by import By
import time
import sys
import os

# TO DO: Fix selenium destruct

class Scraper():
    def __init__(self):
        self.driver = None      # Web Driver
        self.query = None       # Account's Name in this case
        self._setup()
    
    # Destructor
    def __del__(self):
        try:
            self.log_file_success.close()
            self.driver.quit()
        except:
            self.log_file.close()
        else:
            self._log("[SETUP] Web Driver destructed successfully.")
            self.log_file.close()

    # returns a dictionary with website, phone and address if found
    def scrape(self, query):
        self.query = query
        ret_dict = {'shipping-street': "", 
                    'shipping-city': "", 
                    'shipping-state': "", 
                    'shipping-zip': "",
                    'shipping-country': "", 
                    'website': "", 
                    'phone': ""}

        url = self._generate_url(self.query)      # generate query complete url
        self.driver.get(url)                      # navigate to webpage

        # find address
        address_dict = self._scrape_address()
        if type(address_dict) != NoneType:
            ret_dict['shipping-street'] = address_dict.get('street')
            ret_dict['shipping-city'] = address_dict.get('city')
            ret_dict['shipping-state'] = address_dict.get('state')
            ret_dict['shipping-zip'] = address_dict.get('zip')
            ret_dict['shipping-country'] = address_dict.get('country')
        
        # find phone
        ret_dict['phone'] = self._scrape_phone()
        
        # find website
        ret_dict['website'] = self._scrape_website()

        # close window
        #self.driver.close()

        # return
        return ret_dict

    def _setup(self):
        # Create the log file
        log_id = str(time.time_ns())
        
        try:
            os.mkdir("/log")
        except FileExistsError:
            pass

        self.log_file = open("/log/log_file " + log_id + '.txt',mode="w")
        self.log_file_success = open("/log/log_file_success " + log_id + '.txt',mode="w")

        # Setup web driver options
        driver_options = webdriver.ChromeOptions()
        driver_options.add_argument("headless")     # no gui
        #driver_options.add_experimental_option('excludeSwitches', ['enable-logging'])   # idk
        
        # Create driver with selected options
        self.driver = webdriver.Chrome(options=driver_options)
        self._log("[SETUP] Web Driver created successfully.")

    # gets address text and returns a dict divided by shipping street, city, state and zip 
    def _scrape_address(self):
        try:
            address = self.driver.find_element(By.XPATH, '//span[@class="LrzXr"]')
        except:
            self._log("[NOT FOUND] Query (" + self.query + ")'s address not found") 
            return None
        else:
            # address example: 100 N Union St, Montgomery, AL 36104, United States
            address = address.text  # convert address object to text
            self._log_success("[SUCCESS] Query (" + self.query + ")'s address found")
            try:
                # split address by commas
                temp_list = address.split(', ')         # now we have: ['100 N Union St', 'Montgomery', 'AL 36104', 'United States']
                temp_list[2] = temp_list[2].split(" ")  # now we have: ['100 N Union St', 'Montgomery', ['AL', 36104'], 'United States']

                # generate dictionary
                address_dict = {'street': temp_list[0],
                                'city': temp_list[1],
                                'state': temp_list[2][0],
                                'zip': temp_list[2][1],
                                'country': temp_list[3]}
            except:
                self._log("[ERROR] _scrape_address(). Something happened in the address split or dictionary generation")
                return None
            else:
                return address_dict

    # gets address text and returns phone's text
    def _scrape_phone(self):
        try:
            phone = self.driver.find_element(By.XPATH, '//span[@class="LrzXr zdqRlf kno-fv"]/a/span/span')
        except:
            self._log("[NOT FOUND] Query (" + self.query + ")'s phone not found") 
            return ""
        else:
            self._log_success("[SUCCESS] Query (" + self.query + ")'s phone found")
            return phone.text   # convert phone object to text and return

    def _scrape_website(self):
        try:
            website = self.driver.find_element(By.XPATH, '//div[@class="IzNS7c duf-h"]/div/a')
        except:
            self._log("[NOT FOUND] Query (" + self.query + ")'s website not found") 
            return ""
        else:
            self._log_success("[SUCCESS] Query (" + self.query + ")'s website found")
            try:
                return website.get_attribute('href')
            except:
                self._log("[ERROR] _scrape_website(). Something happened in the address split or dictionary generation")

    # prints writes to a file all the operations
    def _log(self, log_msg):
        print(log_msg)
        self.log_file.write(log_msg + "\n")

    def _log_success(self, log_msg):
        print(log_msg)
        self.log_file_success.write(log_msg + "\n")

    # generates a google search url based in the input
    def _generate_url(self, search):
        url = "https://www.google.com/search?hl=en&q="
        try:
            parameters = search.replace(" ", "+")       # replace spaces with "+" signs
        except:
            parameters = " "
        return url + parameters                     # concatenate and return
