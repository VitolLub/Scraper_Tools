import requests
from playwright.sync_api import sync_playwright
from dataclasses import dataclass, asdict, field
import pandas as pd
import argparse
from openpyxl import Workbook,load_workbook
from openpyxl.styles import PatternFill
from bs4 import BeautifulSoup as bs
import json
import csv


class Rieltors:
    def __init__(self):
        self.domain = ''
        self.link_array = []

        self.phone = []
        self.email = []
        self.address = []
        self.city = []
        self.license_r = []
        self.listed_by = []
        self.rieltor_name_array = []

        self.cities_array = []

        self.email_array  = []
        self.phone_array  = []
        self.name_array  = []

        self.state_array = []
        self.city_array = []
        self.zip_array = []

        self.addreses_what_need_check = []
        self.spend_addreses = []
        self.company_array = []
        self.index_for_save = 0

    def make_request(self):
        pass
    def create_file(self):
        wb = Workbook()
        ws = wb.active
        ws['A1'] = "Name"
        ws['B1'] = "Phones"
        ws['C1'] = "Emails"
        ws['D1'] = "State"
        ws['E1'] = "City"
        ws['F1'] = "ZIP"
        ws['G1'] = "Company"

        wb.save("rieltors.xlsx")
    def save_data(self):

        xlsx_file_path = "rieltors.xlsx"
        wb = load_workbook(xlsx_file_path)

        # Step 2: Select the worksheet where you want to append data
        sheet_name = "Sheet"  # Change this to the name of your target sheet
        sheet = wb[sheet_name]

        # Calculate the next row to append data to (assuming data starts from row 2)
        next_row = sheet.max_row + 1
        for i in range(len(self.name_array)):
            try:
                sheet.cell(row=next_row, column=1, value=str(self.name_array[i]))
                sheet.cell(row=next_row, column=2, value=str(self.phone_array[i]))
                sheet.cell(row=next_row, column=3, value=str(self.email_array[i]))
                sheet.cell(row=next_row, column=4, value=str(self.state_array[i]))
                sheet.cell(row=next_row, column=5, value=str(self.city_array[i]))
                sheet.cell(row=next_row, column=6, value=str(self.zip_array[i]))
                sheet.cell(row=next_row, column=7, value=str(self.company_array[i]))
            except Exception as e:
                print(e)
            next_row += 1
        wb.save(xlsx_file_path)

        # Optional: Close the workbook
        wb.close()

        # clear all variables
        for name, value in vars(self).items():
            # clear all variables
            if type(value) is list and name != 'cities_array':
                value.clear()
        print("Data saved")



    def get_summary_field(self,page):
        return page.inner_text('div.SearchList__summary', timeout=500000)

    def check_page(self,page,index,zip_code):
        # check if element exists div.KWPropertyCard__courtesy
        page.wait_for_selector('div.KWPropertyCard__courtesy', timeout=20000)

        properties = page.query_selector_all('div.KWPropertyCard__courtesy')
        properties_adreses = page.query_selector_all('div.KWPropertyCardInfo__address')
        print(f"properties_adreses")
        print(len(properties_adreses))
        for property_adres in properties_adreses:
                if property_adres.inner_text() not in self.addreses_what_need_check and property_adres.inner_text() not in self.spend_addreses:
                    self.addreses_what_need_check.append(property_adres.inner_text())


        # loop through each property
        for property_p in properties:
            property_rieltor_name = property_p.inner_text()
            # print(property_rieltor_name)
            # if Keller, EXP or Century 21 in property name
            if "Keller" in property_rieltor_name or "EXP" in property_rieltor_name or "Century 21" in property_rieltor_name:
                # print(property_rieltor_name)
                # get parent element
                parent = property_p.query_selector("xpath=..")
                # get property link
                link = parent.get_attribute('href')
                # print(link)
                if link not in self.link_array:
                    self.link_array.append(link)

        print(f"self.link_array {len(self.link_array)}")
        # get div class SearchList__summary text
        summary_field = self.get_summary_field(page)
        of_p = summary_field.find("of")
        properties_pos = summary_field.find(" Properties")
        # cut summary_field to get number of properties
        properties_count = summary_field[of_p+3:properties_pos]
        print('properties_count')
        # if index == 0:
        try:
            propersties_count_int = int(properties_count.strip())
        except Exception as e:
            print(e)
            propersties_count_int = 1000
        print(f"propersties_count_int {propersties_count_int}")
        if propersties_count_int > 40:
            # click button with button class KWZoomControl__button
            page.click('button.KWZoomControl__button', timeout=500000)
            # just wait for 5 seconds
            page.wait_for_timeout(20000)
            # index += 1
            self.check_page(page, index,zip_code)
        elif propersties_count_int <= 40:
            for addrs in self.addreses_what_need_check:
                # remove from self.addreses_what_need_check
                self.addreses_what_need_check.remove(addrs)
                # add to self.spend_addreses
                self.spend_addreses.append(addrs)

                self.page_per_page(page, str(addrs), zip_code)




        # if page.inner_text('text="See More"', timeout=500000):
        #     # click Show more button
        #     page.click('text="See More"')
        index += 1
        print(f"self.link_array {len(self.link_array)} and index {index}")
        # if index < 2:
        #     self.check_page(page,index)

    def generate_link(self):
        # link sample https://www.kw.com/search/location/ChIJX9vMLpaC54gRoeGj7jz2wek-0.8984511517167186,Florida%20Ocoee%2034761,Ocoee%2C%20FL%2034761%2C%20USA?fallBackCityAndState=Ocoee%2C%20FL&fallBackPosition=28.5830551%2C%20-81.52672820000001&fallBackStreet=&isFallback=true&viewport=28.61479654954222%2C-81.48971430441892%2C28.52615080507333%2C-81.58086649558103&zoom=13
        state = "Florida"
        city = "Aventura"
        zip = "33160"
        state_code = "FL"
        link = f"https://www.kw.com/search/location/ChIJX9vMLpaC54gRoeGj7jz2wek-0.8984511517167186,{state}%20{city}%20{zip},{city}%2C%20FL%20{zip}%2C%20USA?fallBackCityAndState={city}%2C%20{state_code}&fallBackPosition=28.5830551%2C%20-81.52672820000001&fallBackStreet=&isFallback=true&viewport=28.61479654954222%2C-81.48971430441892%2C28.52615080507333%2C-81.58086649558103&zoom=13"
        print(link)

    def page_per_page(self,page,go_to_link_str,zip_code):
        # generate link
        page.goto("https://www.kw.com/", timeout=100000)

        # set data into input with class KWSearchInput__input
        page.fill('input.KWSearchInput__input', go_to_link_str, timeout=100000)

        # click search enter
        page.press('input.KWSearchInput__input', 'Enter')

        # check if text exist See More
        # wait for 10 seconds
        page.wait_for_timeout(10000)
        # find all div elements with class name 'KWPropertyCard__light'
        # this is the container for each property
        try:
            click_index = 0
            self.search_new_links(page,click_index)
            # self.check_page(page, 0, zip_code)
        except Exception as e:
            print(e)

    def search_new_links(self,page,click_index=None):
        page.wait_for_selector('div.KWPropertyCard__courtesy', timeout=20000)

        properties = page.query_selector_all('div.KWPropertyCard__courtesy')
        # loop through each property
        for property_p in properties:
            property_rieltor_name = property_p.inner_text()
            # print(property_rieltor_name)
            # if Keller, EXP or Century 21 in property name
            if "KELLER" in property_rieltor_name or "CENTURY" in property_rieltor_name or "Keller" in property_rieltor_name or "EXP" in property_rieltor_name or "Century 21" in property_rieltor_name:
                # print(property_rieltor_name)
                # get parent element
                parent = property_p.query_selector("xpath=..")
                # get property link
                link = parent.get_attribute('href')
                # print(link)
                if link not in self.link_array:
                    self.link_array.append(link)

        print(len(self.link_array))

        # mouse move to the span with text See More
        page.hover('text="Save Search"', timeout=10000)
        # scroll down to several pixels
        index = 0
        index2 = 100
        for i in range(0, 10):

            # scroll down to several pixels
            # using JS script
            page.evaluate("window.scrollBy(0, " + str(index) + ")")
            index += 80
            index2 += 100
        # check if bootom of the page
        # check if \"See More\ button exist

        if page.query_selector('button.KWButton.KWButton--secondary.KWButton--medium.KWButton--block'):
            print(f"click_index {click_index}")
            # click Show more button
            if click_index == 5:
                # click button with class KWButton KWButton--secondary KWButton--medium KWButton--block
                page.click('button.KWButton.KWButton--secondary.KWButton--medium.KWButton--block', timeout=20000)
                # page.click('text="See More"', timeout=10000)
                click_index = 0
            # just wait for 10 seconds
            page.wait_for_timeout(10000)
            click_index += 1
            self.search_new_links(page,click_index)

    def save_step(self):
        for link_param in self.link_array:
            link = "https://www.kw.com/" + str(link_param)
            print(link)
            # meake request to link
            lisk_res = requests.get(link)
            print(lisk_res.status_code)
            # print(lisk_res.text)
            if lisk_res.status_code == 200:
                bs_result = bs(lisk_res.text, 'html.parser')

                # find script by id '__NEXT_DATA__'
                script = bs_result.find('script', id='__NEXT_DATA__').text

                json_res = json.loads(script)
                # print(json_res['props']['pageProps']['propertyData']['listingAgentData']['courtesyOfBrokerage'])
                # quit()

                json_data = json_res['props']['pageProps']['propertyData']['listingAgentData']
                json_state = json_res['props']['pageProps']['propertyData']['locator']['address']['state']
                json_city = json_res['props']['pageProps']['propertyData']['locator']['address']['city']
                json_zip = json_res['props']['pageProps']['propertyData']['locator']['address']['zipcode']
                self.state_array.append(json_state)
                self.city_array.append(json_city)
                self.zip_array.append(json_zip)
                self.company_array.append(
                    json_res['props']['pageProps']['propertyData']['listingAgentData']['courtesyOfBrokerage'])

                print(json_data)
                try:
                    full_name = json_data['fullName']
                    if full_name == None:
                        full_name = json_data['brokerLicense']
                except:
                    full_name = ''
                print(full_name)
                # if full_name == None:
                #     print(json_res)
                phones_arr = []

                type = ''
                phone = ''
                try:
                    for phones in json_data['phones']:
                        phoneNumber = phones['phoneNumber']
                        phoneNumberType = phones['phoneNumberType']
                        phones_arr.append(str(phoneNumberType) + " " + str(phoneNumber))
                    print(phones_arr)
                except:
                    phones_arr.append('')
                email = ''
                try:
                    for cont_E in json_data['contactMethods']:
                        if "@" in cont_E['value']:
                            email = cont_E['value']

                except:
                    email = ''
                print(email)

                self.email_array.append(email)
                self.phone_array.append(" ".join(phones_arr))
                self.name_array.append(full_name)
        self.save_data()

    def goto(self):
        # start selenium browser
        print(f"self.cities_array {len(self.cities_array)}")
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=False)
            page = browser.new_page()

            # wait until page is loaded
            index_step = 0
            print(f"self.cities_array {len(self.cities_array)}")
            for go_to_link in self.cities_array:
                zip_code = go_to_link[2]
                go_to_link_str = " ".join(go_to_link)
                # print(type(go_to_link_str))
                print(go_to_link_str)

                self.page_per_page(page,go_to_link_str,zip_code)

                index_step += 1
                self.index_for_save += 1
                print(f"index_step {index_step}")
                print(f"self.link_array {len(self.link_array)}")

                # save
                if len(self.link_array) > 200:
                    print("len(self.link_array) % 10")
                    print(len(self.link_array))
                    self.save_step()


                # if index_step == 4:
                #     break

            # print(len(self.addreses_what_need_check))




            # for link in self.link_array:
            #     # go ti property link and wait until page is loaded
            #
            #     page.goto(str("https://www.kw.com")+str(link), timeout=100000)
            #
            #
            #     # check if page has text 'Log In'
            #     if page.title():
            #         # dont wait for page to load
            #         # get property details
            #         try:
            #             address = page.inner_text('span.PropertyGeneralInfo__addressStreet', timeout=5000)
            #         except:
            #             address = ''
            #         try:
            #             city = page.inner_text('span.PropertyGeneralInfo__addressCity', timeout=5000)
            #         except:
            #             city = ''
            #         license_r = ''
            #         phone = ''
            #         email = ''
            #         rieltor_name = ''
            #         listed_by = ''
            #         try:
            #             listed_by = page.inner_text('span.PropertyGeneralInfo__brokerageInfo', timeout=5000)
            #             listed_by = listed_by.replace('Listing Courtesy Of:','')
            #             # split string by ,
            #             listed_by = listed_by.split(',')[0]
            #             rieltor_name = listed_by[1]
            #         except:
            #             listed_by = ''
            #
            #         print(address,city,listed_by,phone,email,license_r) #
            #         # find div class name 'AgentDetails__info--details', if exist then get data
            #         try:
            #             if page.inner_text('div.AgentDetails__info--details', timeout=5000):
            #                 try:
            #                     phone  = page.inner_text('div.AgentDetails__contact--phones', timeout=1000)
            #                     print(phone)
            #                 except:
            #                     phone = ''
            #                 try:
            #                     email = page.inner_text('a.AgentDetails__contact--email', timeout=1000)
            #                     print(email)
            #                 except:
            #                     email = ''
            #                 try:
            #                     license_r = page.inner_text('span.AgentDetails__info--license', timeout=1000)
            #                     print(license_r)
            #                 except:
            #                     license_r = ''
            #         except:
            #             print('Bad')
            #         print('-------------------')
            #         print(f"index: {len(self.phone)}")
            #         self.phone.append(phone)
            #         self.email.append(email)
            #         self.address.append(address)
            #         self.city.append(city)
            #         self.license_r.append(license_r)
            #         self.listed_by.append(listed_by)
            #         self.rieltor_name_array.append(rieltor_name)
            #
            # self.save_data()





            # just wait for 10 minutes before closing the browser
            # page.wait_for_timeout(600000)





    def read_csv_file(self):
        self.cities_array = []
        # read csv file uscities.csv line by line
        with open('uscities.csv', newline='') as csvfile:
            # get state_name
            reader = csv.DictReader(csvfile)
            for row in reader:
                if row['state_name'] == 'Florida':
                    zip_arr = row['zips'].split(' ')
                    for zip in zip_arr:
                        # print(row['state_name'],row['city_ascii'],zip)
                        self.cities_array.append([str(row['state_name'])+",",str(row['city_ascii'])+",",str(zip)])

        return self.cities_array


if __name__ == "__main__":
    rieltors = Rieltors()
    rieltors.read_csv_file()
    rieltors.create_file()
    rieltors.domain = "https://www.kw.com/search/location/ChIJY10Hv_i02YgRjdzvoWOVM6w-0.7420868142967443,Florida%2C%20Miami%2C%2033109,Miami%20Beach%2C%20FL%2033109%2C%20USA?fallBackCityAndState=Miami%20Beach%2C%20FL&fallBackPosition=25.7560139%2C%20-80.1344842&fallBackStreet=&isFallback=true&viewport=25.872362435965854%2C-80.1049454695791%2C25.826943802010476%2C-80.15103654929102&zoom=14"
    rieltors.goto()
    # rieltors.generate_link()



    # rieltors.read_csv_file()

