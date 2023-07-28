import time
from bs4 import BeautifulSoup as bs
import requests
from openpyxl import Workbook,load_workbook
from openpyxl.styles import PatternFill
import json
import sys
import random

class ShopifyScrapper:

    def __init__(self):
        self.url = ''
        self.id_by_id_arr = list()
        self.id_arr = []
        self.product_name_arr = []
        self.price_arr = []
        self.full_link_arr = []
        self.data_value_list_arr = []
        self.handle_arr = []
        self.related_collections_handle_arr = []
        self.full_description_arr = []
        self.variants_arr = []
        self.full_description_html_arr = []
        self.title_arr = []
        self.title_html_arr = []
        self.ceo_title_arr = []
        self.ceo_description_arr = []
        self.images_arr = []
        self.imge_primary_arr = []
        self.domain = 'https://univers-chinois.com'
        self.https = 'https:'
        self.variants_arr_primary = []
        self.bullet_points_arr = []
        self.h2_html_arr = []
        self.product_id_arr = []
        self.tags_arr = []
        self.vendor_arr = []
        self.type_arr = []
        self.images_arr_variant = []
        self.secure_url_arr = []
        self.published_at = list()
        self.created_at = list()
        self.available_arr = []
        self.compare_at_price_varies_arr = []
        self.price_varies_arr = []
        self.product_counter = []
        self.compare_at_price_arr = []
        self.compare_at_price_max_arr = []

    def request_link_by_link(self,link,proxy_index,s):
        # get id after ?variant=
        id_lisitng = link.split('?variant=')[1]
        # print(link)
        # time.sleep(1)
        response = self.make_request(link,proxy_index,s)
        soup = bs(response.text, 'html.parser')
        # get h1 with class product_name
        try:
            product_name = soup.find('h1', class_='product_name').text
        except:
            product_name = ''

        # get all div with data-value
        data_value_list = []
        # if div has data-value, append to data_value_list
        for div in soup.find_all('div', {'data-value': True}):
            data_value_list.append(div['data-value'])

        try:
            # get meta with property product:price:amount
            price = soup.find('meta', {'property': 'product:price:amount'})['content']
        except:
            price = ''



        try:
            title = soup.find('h1',class_="product_name").text
            title_html = soup.find('h1',class_="product_name")
        except:
            title = ''
            title_html = ''

        try:
            # get title from head
            ceo_title = soup.find('title').text
            ceo_title = ceo_title.strip()
            ceo_description = soup.find('meta', {'name': 'description'})['content']
        except:
            ceo_title = ''
            ceo_description = ''

        images_arr = []
        try:
            images = soup.find('div', class_='product_gallery_nav').find_all('img')
            if len(images) > 0:
                for image in images:
                    if image['src'].find('_300x.') > -1:
                        # replace all _300x. to _800x.
                        images_arr.append(self.https+image['src'].replace('_300x.', '_800x.'))
                    else:
                        images_arr.append(self.https+image['src'])
            else:
                images = soup.find('div', class_='image__container').find_all('img')
                for image in images:
                    if image['data-src'].find('_300x.') > -1:
                        # replace all _300x. to _800x.
                        images_arr.append(self.https+image['data-src'].replace('_300x.', '_800x.'))
                    else:
                        images_arr.append(self.https+image['data-src'])
        except:
            images_arr = []



        try:
            bullet_points_arr = []
            bullet_points = soup.find_all('dl', class_='accordion')
            for bullet_point in bullet_points:
                # get html
                bullet_points_arr.append(str(bullet_point))
        except Exception as e:
            bullet_points_arr = []

        try:
            h2_html = soup.find('div', class_='description').find('h2')
        except:
            h2_html = ''

        try:
            vendor = ''
            type = ''
            # find all script
            scripts = soup.find_all('script')
            for script in scripts:
                # find var meta
                if script.text.find('var meta') > -1:
                    # cut string
                    # find { from start and }}; from end
                    script_text = script.text[script.text.find('var meta = {')+10:script.text.rfind('}};') + 2]
                    print("script_text")
                    # convert string to json
                    script_json = json.loads(script_text)
                    id_by_id = script_json['product']['id']
                    vendor = script_json['product']['vendor']
                    type  =  script_json['product']['type']

        except:
            id_by_id = ''

        try:
            related_collections_arr = []
            try:
                related_collections = soup.find('div', class_='product-links').find_all('a')
            except:
                related_collections = soup.find('div', class_='product_links').find_all('a')

            print('related_collections_href')
            # get href
            for related_collections_a in related_collections:
                related_collections_href = related_collections_a['href']

                hendle_pos = related_collections_href.find('/collections/')
                related_collections_handle = related_collections_href[hendle_pos + 13:]
                related_collections_arr.append(related_collections_handle)
        except Exception as e:
            print(e)
            related_collections_arr = []

        try:
            product_data = soup.find('div', class_='product_form')
            # get attributes data-product
            product_data = product_data['data-product']
            print(product_data)

            # convert string to json
            product_data = json.loads(product_data)
            tags_arr = product_data['tags']
            self.published_at.append(product_data['published_at'])
            self.created_at.append(product_data['created_at'])
            self.available_arr.append(product_data['available'])
            self.compare_at_price_varies_arr.append(product_data['compare_at_price_varies'])
            self.price_varies_arr.append(product_data['price_varies'])
            self.compare_at_price_arr.append(self.cut_compare_price(product_data['compare_at_price']))
            self.compare_at_price_max_arr.append(self.cut_compare_price(product_data['compare_at_price_max']))


            full_description_html = product_data['content']
            # remove h2 from full_description_html
            soup = bs(full_description_html, 'html.parser')
            for h2 in soup.find_all('h2'):
                h2.decompose()
            full_description_html = str(soup)

            full_description = product_data['description']
            # full_description convert to text
            full_description = bs(full_description, 'html.parser').text

            secure_url = ''
            for img_variants in product_data['variants']:

                if str(img_variants['id']) == str(id_lisitng):
                    secure_url = 'https:'+ str(img_variants['featured_image']['src'])
                    variants = []
                    if img_variants['option1'] == None:
                        variants.append('')
                    else:
                        variants.append(img_variants['option1'])

                    if img_variants['option2'] == None:
                        variants.append('')
                    else:
                        variants.append(img_variants['option2'])

                    if img_variants['option3'] == None:
                        variants.append('')
                    else:
                        variants.append(img_variants['option3'])


            if len(secure_url) == 0:
                secure_url = ''

        except Exception as e:
            print(e)
            # show error line
            exc_type, exc_obj, exc_tb = sys.exc_info()
            print(exc_tb.tb_lineno)

            full_description = ''
            full_description_html = ''
            secure_url = ''

        return product_name, price,link,",".join(data_value_list), full_description, full_description_html,title,title_html,ceo_title,ceo_description,images_arr,variants,bullet_points_arr,h2_html,id_by_id,tags_arr, vendor, type,related_collections_arr,secure_url

    def cut_compare_price(self,compare_at_price):
        a = str(compare_at_price)
        pos2 = a[len(a) - 2:]
        pos1 = a[:len(a) - 2]
        compare_at_price = str(pos1) + ',' + (pos2)

        return compare_at_price

    def save_to_xlsx_product_count(self):
            xlsx_file_path = "shopify.xlsx"
            wb = load_workbook(xlsx_file_path)

            # Step 2: Select the worksheet where you want to append data
            sheet_name = "Sheet"  # Change this to the name of your target sheet
            sheet = wb[sheet_name]

            # Calculate the next row to append data to (assuming data starts from row 2)
            next_row = sheet.max_row + 1

            sheet.cell(row=next_row, column=1, value="Product Quantity:"+str(len(self.product_counter)))

            wb.save(xlsx_file_path)
            wb.close()


    def save_to_xlsx(self,id_by_id_arr,product_name_arr, price_arr,full_link_arr,
                    data_value_list_arr,variants_arr,related_collections_handle_arr,handle_arr,
                    full_description_arr,full_description_html_arr,title_arr,title_html_arr,ceo_title_arr,
                    ceo_description_arr,images_arr,imge_primary_arr,variants_arr_primary,bullet_points_arr,h2_html_arr,product_id_arr,tags_arr,
                     vendor_arr,type_arr):
        xlsx_file_path = "shopify.xlsx"
        wb = load_workbook(xlsx_file_path)

        # Step 2: Select the worksheet where you want to append data
        sheet_name = "Sheet"  # Change this to the name of your target sheet
        sheet = wb[sheet_name]

        # Calculate the next row to append data to (assuming data starts from row 2)
        next_row = sheet.max_row + 1

        for i in range(len(id_by_id_arr)):
            try:
                option1 = variants_arr_primary[i][0]
            except:
                option1 = ''
            try:
                option2 = variants_arr_primary[i][1]
            except:
                option2 = ''
            try:
                option3 = variants_arr_primary[i][2]
            except:
                option3 = ''

            try:
                bullet_points_variant1 = bullet_points_arr[i][0]
            except:
                bullet_points_variant1 = ''

            try:
                bullet_points_variant2 = bullet_points_arr[i][1]
            except:
                bullet_points_variant2 = ''

            try:
                bullet_points_variant3 = bullet_points_arr[i][2]
            except:
                bullet_points_variant3 = ''

            try:
                total_description_html = str(h2_html_arr[0]) + str(full_description_html_arr[0]) + str(bullet_points_variant1) + str(bullet_points_variant2) + str(bullet_points_variant3)

            except:
                total_description_html = ''

            # Step 4: Append the data to the selected worksheet
            # ws.append([str(self.secure_url_arr[i])," ",str(price_arr[i])])
            sheet.cell(row=next_row, column=1, value=str(id_by_id_arr[i]))
            sheet.cell(row=next_row, column=2, value=str(product_id_arr[i]))
            sheet.cell(row=next_row, column=3, value=str(full_link_arr[i]))
            sheet.cell(row=next_row, column=4, value=str(handle_arr[i]))
            sheet.cell(row=next_row, column=5, value=str(related_collections_handle_arr[i]))
            sheet.cell(row=next_row, column=6, value=str(related_collections_handle_arr[i]))
            sheet.cell(row=next_row, column=7, value=str(title_arr[i]))
            sheet.cell(row=next_row, column=8, value=str(title_html_arr[i]))
            sheet.cell(row=next_row, column=9, value=str(ceo_title_arr[i]))
            sheet.cell(row=next_row, column=10, value=str(ceo_description_arr[i]))
            sheet.cell(row=next_row, column=11, value=str(product_name_arr[i]))
            sheet.cell(row=next_row, column=12, value=str(full_description_arr[i]))
            sheet.cell(row=next_row, column=13, value=str(total_description_html)).fill = PatternFill(start_color='ADFF2F', end_color='ADFF2F', fill_type='solid')
            sheet.cell(row=next_row, column=14, value=str(h2_html_arr[i])).fill = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
            sheet.cell(row=next_row, column=15, value=str(full_description_html_arr[i])).fill = PatternFill(start_color='B0E0E6', end_color='B0E0E6', fill_type='solid')
            sheet.cell(row=next_row, column=16, value=str(bullet_points_variant1)).fill = PatternFill(start_color='CD853F', end_color='CD853F', fill_type='solid')
            sheet.cell(row=next_row, column=17, value=str(bullet_points_variant2)).fill = PatternFill(start_color='FFEBCD', end_color='FFEBCD', fill_type='solid')
            sheet.cell(row=next_row, column=18, value=str(bullet_points_variant3)).fill = PatternFill(start_color='FF7F50', end_color='FF7F50', fill_type='solid')
            sheet.cell(row=next_row, column=19, value=str(self.published_at[i]))
            sheet.cell(row=next_row, column=20, value=str(self.created_at[i]))
            sheet.cell(row=next_row, column=21, value=str(vendor_arr[i]))
            sheet.cell(row=next_row, column=22, value=str(type_arr[i]))
            sheet.cell(row=next_row, column=23, value=str(tags_arr[i]))
            sheet.cell(row=next_row, column=24, value=str(price_arr[i]))
            sheet.cell(row=next_row, column=25, value=str(price_arr[i]))
            sheet.cell(row=next_row, column=26, value=str(price_arr[i]))
            sheet.cell(row=next_row, column=27, value=str(self.available_arr[i]))
            sheet.cell(row=next_row, column=28, value=str(self.price_varies_arr[i]))
            sheet.cell(row=next_row, column=29, value=str(self.compare_at_price_arr[i]))
            sheet.cell(row=next_row, column=30, value=str(self.compare_at_price_max_arr[i]))
            sheet.cell(row=next_row, column=31, value=str(self.compare_at_price_varies_arr[i]))
            sheet.cell(row=next_row, column=32, value=str(''))
            sheet.cell(row=next_row, column=33, value=str(''))
            sheet.cell(row=next_row, column=34, value=str(images_arr[i]))
            sheet.cell(row=next_row, column=35, value=str(imge_primary_arr[i]))
            sheet.cell(row=next_row, column=36, value=str(variants_arr[i]))
            sheet.cell(row=next_row, column=37, value=str(option1))
            sheet.cell(row=next_row, column=38, value=str(option2))
            sheet.cell(row=next_row, column=39, value=str(option3))
            sheet.cell(row=next_row, column=40, value=str(self.secure_url_arr[i]))
            sheet.cell(row=next_row, column=41, value=str(''))
            sheet.cell(row=next_row, column=42, value=str(price_arr[i]))
            next_row += 1

        # Step 5: Save the changes back to the XLSX file
        wb.save(xlsx_file_path)

        # Optional: Close the workbook
        wb.close()

        # clear
        # get all variables from __init__
        for name, value in vars(self).items():
            # clear all variables
            if type(value) is list and name != 'product_counter':
                value.clear()


    def save_to_csv(self,id_by_id_arr,product_name_arr, price_arr,full_link_arr,
                    data_value_list_arr,variants_arr,related_collections_handle_arr,handle_arr,
                    full_description_arr,full_description_html_arr,title_arr,title_html_arr,ceo_title_arr,
                    ceo_description_arr,images_arr,imge_primary_arr):
        import csv
        with open('shopify.csv', 'w', newline='') as file:
            writer = csv.writer(file)
            writer.writerow(["id","full_link","handle","collection_handele","related_collections_handle","title","title_html","ceo_title","ceo_description","product_name","full_description","full_description_html","h2_html","bullet_points_html","published_at","created_at","vendor","type","tags","price","price_min","price_max","available","price_varies","compare_at_price","compare_at_price_max","compare_at_price_varies","requires_selling_plan","selling_plan_groups","images","featured_image","variants","variant featured_image","variant compare_at_price","variant price"])
            for i in range(len(product_name_arr)):
                writer.writerow([str(id_by_id_arr[i]),str(full_link_arr[i]),str(handle_arr[i]),str(related_collections_handle_arr[i]),str(related_collections_handle_arr[i]),str(title_arr[i]),str(title_html_arr[i]),str(ceo_title_arr[i]),str(ceo_description_arr[i]),str(product_name_arr[i]), str(full_description_arr[i]),str(full_description_html_arr[i]),'','','','','','','',str(price_arr[i]),str(price_arr[i]),str(price_arr[i]),"TRUE","0","","","0","","",str(images_arr[i]),str(imge_primary_arr[i]),str(variants_arr[i][0]),str(images_arr[i]),"",str(price_arr[i])])
    def get_handle_and_collection_handle(self,link):
        # remove all after ? in link
        l_pos = link.find('?')
        if l_pos > 0:
            link = link[:l_pos]
        hendle = link.split('/')[-1]
        collection_handle = link.split('/')[-3]
        return hendle,collection_handle
    def request_link_by_link_to_get_ids(self,link,proxy,s):
        # request link
        # time.sleep(1)
        response = self.make_request(link,proxy,s)
        soup = bs(response.text, 'html.parser')

        # get div with class swatch_options and find all data-id and append to id_arr
        id_arr = []
        variants_arr = []
        #
        scripts = soup.find_all('script')
        for script in scripts:
            # find var meta
            if script.text.find('var meta') > -1:
                # cut string
                # find { from start and }}; from end
                script_text = script.text[script.text.find('var meta = {') + 10:script.text.rfind('}};') + 2]
                print("script_text")
                # convert string to json
                script_json = json.loads(script_text)
                id_by_ids = script_json['product']['variants']
                for id_by_id in id_by_ids:
                    id_arr.append(id_by_id['id'])
        for div in soup.find_all('div', class_='swatch_options'):
            for variant in div.find_all('div', class_='option_title'):
                # if div has text, append to handle_arr
                if len(div.text) > 0:
                    variants_arr.append(variant.text)
        return id_arr,variants_arr
    def make_request(self,url,proxy,s):
        response = s.get(url, proxies=proxy, verify=False, timeout=5)
        if response.status_code == 200:
            return response
        else:
            time.sleep(40)
            return self.make_request(url,proxy,s)

    def scrap_shopify(self,all_categpries):
        domain = self.domain
        index = 0
        for category in all_categpries:
            url = domain + category
            print(f"url {url}")
            ip_addresses = [
                "173.176.14.246:3128",
                "129.153.157.63:3128",
                "141.101.115.2:80",
                "172.67.34.58:80",
                "172.67.177.251:80",
                "203.22.223.150:80"
            ]
            try:
                proxy_index = random.choice(ip_addresses)
                proxy = {"http": proxy_index}
                s = requests.Session()
                response = self.make_request(url,proxy,s)
            except:
                response = requests.get(url, timeout=5, headers={'User-Agent': 'Mozilla/5.0'})
                print("Skipping. Connnection error")
            # request link with proxy
            # time.sleep(1)
            # response = requests.get(url, timeout=5, headers={'User-Agent': 'Mozilla/5.0'})
            soup = bs(response.text, 'html.parser')

            # find all 'a' with class product-info__caption and get href
            for link in soup.find_all('a', class_='product-info__caption'):
                link = link.get('href')
                fill_link = domain + link
                print(f"fill_link {fill_link}")
                id_arr,variants_arr = self.request_link_by_link_to_get_ids(fill_link,proxy,s)
                # print(id_arr)
                for id in id_arr:
                    full_link = fill_link + '?variant=' + str(id)
                    print(full_link)
                    product_name, price,full_link,data_value_list,full_description,full_description_html,title,title_html,ceo_title,ceo_description,images_arr,variants,bullet_points_arr,h2_html,id_by_id,tags_arr,vendor, type, related_collections_arr,secure_url = self.request_link_by_link(full_link,proxy,s)
                    # print(product_name, price,full_link,data_value_list)
                    # print(id_by_id,id)

                    self.id_by_id_arr.append(id)
                    self.product_name_arr.append(product_name)
                    if product_name not in self.product_counter:
                        self.product_counter.append(product_name)
                    self.price_arr.append(price)
                    self.full_link_arr.append(full_link)
                    self.data_value_list_arr.append(data_value_list)
                    self.variants_arr.append(','.join(variants_arr))
                    # title section
                    self.title_arr.append(title)
                    self.title_html_arr.append(title_html)
                    self.ceo_title_arr.append(ceo_title)
                    self.ceo_description_arr.append(ceo_description)

                    handle, collection_handele = self.get_handle_and_collection_handle(full_link)

                    self.related_collections_handle_arr.append(','.join(related_collections_arr))
                    self.handle_arr.append(handle)
                    self.full_description_arr.append(full_description)
                    if self.domain == 'https://miss-minceur.com':
                        full_description_html_res = full_description_html.split('✂')
                        print(f"full_description_html_res {len(full_description_html_res)}")
                        self.full_description_html_arr.append(full_description_html_res[0])
                        bullet_points_arr.clear()
                        # full_description_html_res remove first element
                        print(len(full_description_html_res[1:]))
                        full_description_html_step = full_description_html_res[1:]
                        for elem in full_description_html_step:
                            bullet_points_arr.append(elem)

                        try:
                            self.bullet_points_arr.append(bullet_points_arr)
                        except:
                            self.bullet_points_arr.append('')
                    else:
                        self.full_description_html_arr.append(full_description_html)
                        try:
                            self.bullet_points_arr.append(bullet_points_arr)
                        except:
                            self.bullet_points_arr.append('')

                    try:
                        self.imge_primary_arr.append(images_arr[0])
                    except:
                        self.imge_primary_arr.append('')
                    self.images_arr_variant.append(images_arr[0])
                    self.images_arr.append(','.join(images_arr))
                    self.variants_arr_primary.append(variants)

                    self.h2_html_arr.append(h2_html)
                    self.product_id_arr.append(id_by_id)
                    self.tags_arr.append(','.join(tags_arr))
                    self.vendor_arr.append(vendor)
                    self.type_arr.append(type)
                    self.secure_url_arr.append(secure_url)
                    print("++++++++++++")
                    print(f'INDEX {index}, {variants}')
                    print(f"ID COUNT {len(self.product_counter)}")
                    # remove duplicates from list
                    print("++++++++++++")
                    if index% 200 == 0:
                        self.save_to_xlsx(self.id_by_id_arr,self.product_name_arr, self.price_arr,self.full_link_arr,
                         self.data_value_list_arr,self.variants_arr,self.related_collections_handle_arr,
                         self.handle_arr,self.full_description_arr,self.full_description_html_arr,
                         self.title_arr,self.title_html_arr,self.ceo_title_arr,self.ceo_description_arr,
                          self.images_arr,self.imge_primary_arr,self.variants_arr_primary,
                          self.bullet_points_arr,self.h2_html_arr,self.product_id_arr,self.tags_arr,
                          self.vendor_arr,self.type_arr)
                    index += 1
                    if index == 84:
                        break

                if index == 84:
                    break
            if index == 84:
                break
        self.save_to_xlsx(self.id_by_id_arr, self.product_name_arr, self.price_arr, self.full_link_arr,
                          self.data_value_list_arr, self.variants_arr, self.related_collections_handle_arr,
                          self.handle_arr, self.full_description_arr, self.full_description_html_arr,
                          self.title_arr, self.title_html_arr, self.ceo_title_arr, self.ceo_description_arr,
                          self.images_arr, self.imge_primary_arr, self.variants_arr_primary,
                          self.bullet_points_arr, self.h2_html_arr, self.product_id_arr, self.tags_arr,
                          self.vendor_arr, self.type_arr)
        self.save_to_xlsx_product_count()





    def get_menu_links(self):
        url = self.domain
        all_categpries = []
        response = requests.get(url)
        soup = bs(response.text, 'html.parser')
        # for link in soup.find('div', class_='main-nav__wrapper').find_all('a'):
        for link in soup.find('div', class_='nav nav--combined clearfix').find_all('a'):
            menu_link = link.get('href')
            if menu_link.startswith('/collections'):
                all_categpries.append(menu_link)

        return all_categpries


    def create_xls_file(self):
        # create shopify.xlsx file

        wb = Workbook()
        ws = wb.active
        ws.append(["id","product ID","full_link","handle","collection_handele","related_collections_handle","title","title_html","ceo_title","ceo_description","product_name","full_description","full_description_html","h2_html","description_html","bullet_points_html","bullet_points_html","bullet_points_html","published_at","created_at","vendor","type","tags","price","price_min","price_max","available","price_varies","compare_at_price","compare_at_price_max","compare_at_price_varies","requires_selling_plan","selling_plan_groups","images","featured_image","variants","option1","option2","option3","variant featured_image","variant compare_at_price","variant price"])
        wb.save("shopify.xlsx")

if __name__ == "__main__":
    shopify_scrapper = ShopifyScrapper()
    shopify_scrapper.domain = "https://miss-minceur.com"
    shopify_scrapper.create_xls_file()
    all_categpries = shopify_scrapper.get_menu_links()
    print(all_categpries)
    shopify_scrapper.scrap_shopify(all_categpries)

