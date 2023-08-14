import time
from bs4 import BeautifulSoup as bs
import requests
from openpyxl import Workbook,load_workbook
from openpyxl.styles import PatternFill
import json
import sys
import pandas as pd
import random

class ShopifyScrapper:

    def __init__(self):
        self.webarchive = False
        self.webarchive_url = ''
        self.webarchive_url_domain = ''
        self.dublicate = []
        self.increment_index = 0
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
        self.price_min_arr = []
        self.price_max_arr = []

        # full decsription arr
        self.total_description_html_arr = []
        self.clean_description_html_arr = []

        self.hanle = ''

    def remove_all_none_tags(self,soup):
        # print('remove_all_none_tags')
        # remove all ul is not None and len(ul.text) > 40
        uls = soup.find_all(['p','h4','ul','ol','table'])
        for ul in uls:
            # print(ul)
            if ul is None or len(ul.text) < 40:
                ul.decompose()
        # print('remove_all_none_tags END')
        return soup

    def cut_full_description(self,soup,full_description):
        fill_description_primary = ''
        bullet_points_arr = []
        related_col_arr = []
        h2_html_origin = ''

        # remove all  attributes data-mce-fragment
        style_tags = soup.find_all(style=True)
        for style_tag in style_tags:
            # remove style attr
            del style_tag['style']
            del style_tag['data-mce-style']
            del style_tag['data-mce-fragment']
            del style_tag['class']
            del style_tag['data-mce-fragment']
            del style_tag['data-mce-selected']
            del style_tag['width']
            del style_tag['border']
            del style_tag['data-sheets-value']
        soup = soup


        # full_description_html = soup.find('div', class_='description')
        # try:
        #     bullet_points_arr = []
        #     bullet_points = full_description_html.find_all('dl', class_='accordion')
        #     for bullet_point in bullet_points:
        #         # get html
        #         bullet_points_arr.append(str(bullet_point))
        # except Exception as e:
        #     print(e)
        #     bullet_points_arr = []
        #     print(e)
        #     print("Error on line {}".format(sys.exc_info()[-1].tb_lineno))

        # print("cut_full_description3")

        fill_description_primary = ''
        full_description_html = ''
        if self.domain == 'https://miss-minceur.com' or self.domain == 'https://www.univers-fleuri.com':
            # print("cut_full_description4")
            # full_description_html = soup.find('div', class_='description bottom')
            #
            # # remove all style tags
            # full_style_tags = full_description_html.find_all(style=True)
            # for style_tag2 in full_style_tags:
            #     # remove style attr
            #     del style_tag2['style']
            #     del style_tag2['data-mce-style']
            #     del style_tag2['data-mce-fragment']
            #     del style_tag2['class']
            #     del style_tag2['data-mce-fragment']
            #     del style_tag2['data-mce-selected']
            #     del style_tag2['width']
            #     del style_tag2['border']
            #     del style_tag2['data-sheets-value']
            #
            # full_description_html = full_description_html


            # remove h2
            # h2_html = full_description_html.find_all('h2')
            # index = 0
            # for h2 in h2_html:
            #     if index == 0:
            #         h2_html_origin = str(h2)
            #     h2.decompose()
            #     index += 1
            # full_description_html = bs(str(full_description_html), 'html.parser')
            # get html
            full_description_html = str(full_description)
            # print(full_description_html)
            # quit()
            full_description_html_res = []
            # if str(full_description_html).find('✂') != -1:
            #     full_description_html_res = str(full_description_html).split('✂')
            # elif str(full_description_html).find('▶') != -1:
            #     full_description_html_res = str(full_description_html).split('▶')
            # elif str(full_description_html).find('👗') != -1:
            #     full_description_html_res = str(full_description_html).split('👗')
            # elif str(full_description_html).find('<table') != -1 and str(full_description_html).find('<ul') != -1:
            #     # cut data inside table
            #     tb_data = bs(full_description_html, 'html.parser')
            #     tbale_data = tb_data.find_all('table')
            #     table_data = ''
            #     for table in tbale_data:
            #         table_data = str(table)
            #         print(table_data)
            #         table.decompose()
            #
            #
            #     pall_primary = ''
            #     pall_data = tb_data.find_all('p')
            #     for p in pall_data:
            #         pall_primary += str(p)
            #     full_description_html_res.append(str(pall_primary))
            #
            #     ulale_data = tb_data.find_all('ul')
            #     tul_data = ''
            #     for uls in ulale_data:
            #         tul_data = uls
            #     full_description_html_res.append(str(tul_data))
            #     full_description_html_res.append(str(table_data))
            #
            # elif str(full_description_html).find('<ul') != -1:
            # cut data inside table
            # get data inside table
            # print(f"Ul pattern")



            ul_data = bs(full_description_html, 'html.parser')
            # remove all h2
            h2_data = ul_data.find_all('h2')

            for h2 in h2_data:
                h2.decompose()

            ul_data = ul_data

            # find all ul and p
            ul_data = self.remove_all_none_tags(ul_data)

            ul_data = ul_data.find_all(['p', 'h4', 'ul', 'ol', 'table'])
            ul_index = 0
            tag_type = ''
            tag_value = ''
            description_status = False
            count_of_tags = len(ul_data)
            tga_index = 0
            # print(f"count_of_tags {count_of_tags}")
            not_iquel_status = False
            for ul in ul_data:
                # get tag name


                if ul is not None and len(ul.text) > 40:

                    tag_name = ul.name
                    if tag_type == '':
                        tag_type = tag_name

                    # print(f"tag_type {tag_type} tag_name {tag_name} ")
                    # print(ul)
                    # print(f"tag_name {tag_name}")
                    if tag_type == tag_name:
                        # print('Iquil')
                        # print(full_description_html_res)
                        # quit()
                        # tag_value += str(ul)
                        tag_value +=  str(ul) #str(full_description_html_res[-1]) +
                        # full_description_html_res.pop(-1)
                        # full_description_html_res.append(tag_value)
                        # print(tag_value)
                        # full_description_html_res[-1] = str(full_description_html_res[-1]) + str(ul)
                        if ul_index == len(ul_data) - 1:
                            # print(f"Last index")
                            # print(tag_type)
                            # print(f"description_status {description_status}")
                            if tag_type == 'h4' and description_status == False or tag_type == 'p' and description_status == False:
                                full_description_html_res.insert(0, str(tag_value))
                                description_status = True
                            else:
                                full_description_html_res.append(tag_value)

                            # add tag_value to last element in array
                            #
                            # print(full_description_html_res)
                            # quit()
                            # last_iteration_data = str(full_description_html_res[-1]) + str(tag_value)
                            # # # remove last element
                            # full_description_html_res.pop(-1)
                            # full_description_html_res.append(tag_value)

                    elif tag_type != tag_name:
                        # print(f"Not iquil")
                        # print(tag_type)
                        # print(not_iquel_status)
                        # print(ul)
                        # print(f"description_status {description_status}")
                        # if not_iquel_status == True:
                        if tag_type == 'h4' and description_status == False or tag_type == 'p' and description_status == False:
                            full_description_html_res.insert(0, str(tag_value))
                            description_status = True
                        else:
                            full_description_html_res.append(tag_value)

                        # reinstall values
                        tag_type = tag_name
                        # tag_value = ''
                        tag_value = str(ul)
                        # print(f"Reinstall variables")
                        #
                        # print("______________________________________")
                        # tag_value = ''

                ul_index += 1
            # print('===============================')
            # print('full_description_html_res')
            # print(len(full_description_html_res))
            # for tex in full_description_html_res:
            #     print(tex)
            #     print('===============================')
            # quit()
            # print('full_description_html_res')
            #
            # print(len(full_description_html_res))
            # for full_description_html_res_item in full_description_html_res:
            #     print(full_description_html_res_item)
            #     print('------------------')
            #
            # quit()

            #     pall_primary = ''
            #     pall_data = ul_data.find_all('p')
            #     for p in pall_data:
            #         pall_primary += str(p)
            #     full_description_html_res.append(str(pall_primary))
            #
            #     ulale_data = ul_data.find_all('ul')
            #     tul_data = ''
            #     for uls in ulale_data:
            #         tul_data = uls
            #     full_description_html_res.append(str(tul_data))
            #
            # elif str(full_description_html).find('<table') != -1:
            #     # cut data inside table
            #     # get data inside table
            #     tb_data = bs(full_description_html, 'html.parser')
            #
            #     tbale_data = tb_data.find_all('table')
            #     table_data = ''
            #     for table in tbale_data:
            #         table_data = table
            #         table.decompose()
            #
            #     pall_primary = ''
            #     pall_data = tb_data.find_all('p')
            #     for p in pall_data:
            #         pall_primary+=str(p)
            #     full_description_html_res.append(str(pall_primary))
            #     full_description_html_res.append(str(table_data))

            # print(f"full_description_html_res {full_description_html_res}")
            try:
                # print('++++++++++++++++++++++++++')
                # print(type(full_description_html_res))
                # print(len(full_description_html_res))
                fill_description_primary = full_description_html_res[0]
            except Exception as e:
                # print("--------------------------")
                # print(full_description_html)
                fill_description_primary = full_description_html
                print(e)
                print("Error on line {}".format(sys.exc_info()[-1].tb_lineno))


            bullet_points_arr.clear()

            # full_description_html_res remove first element
            full_description_html_step = full_description_html_res[1:]
            for elem in full_description_html_step:
                bullet_points_arr.append(elem)
            # print(soup)
            try:
                related_col_arr = []
                arr1 = []
                arr1.append('related_collections')
                related_col_arr.append(arr1)
                # related_collections = soup.find('div', class_='product_links').find_all('a')
                # for related_collection in related_collections:
                #     # print(related_collection)
                #     related_collection = related_collection['href']
                #     handle, collection_handele = self.get_handle_and_collection_handle(related_collection)
                #     # print(handle)
                #     related_col_arr.append(handle)
            except Exception as e:
                print(e)
                related_col_arr.append('')
                print(e)
                print("Error on line {}".format(sys.exc_info()[-1].tb_lineno))


        elif self.domain == 'https://univers-chinois.com':
            related_col_arr = []
            fill_description_primary = ''
            full_description_html = ''
            ss = bs(full_description, 'html.parser')

            full_style_tags = ss.find_all(style=True)
            for style_tag2 in full_style_tags:
                # remove style attr
                del style_tag2['class']
                del style_tag2['data-mce-fragment']
                del style_tag2['data-mce-selected']
                del style_tag2['width']
                del style_tag2['border']
                del style_tag2['data-sheets-value']
            ss = ss

            # find all p tags
            p_tags = ss.find_all('p')
            index = 0
            for p_tag in p_tags:
                # if p tag has parent div
                if index < 2:
                    fill_description_primary += str(p_tag)
                index += 1

            try:

                related_collections = soup.find('div', class_='product-links').find_all('a')
                for related_collection in related_collections:
                    related_collection = related_collection['href']
                    handle, collection_handele = self.get_handle_and_collection_handle(related_collection)
                    related_col_arr.append(handle)
            except:
                related_col_arr = []


        # print(fill_description_primary)
        # print(bullet_points_arr)
        # print(related_col_arr)
        # quit()

        return fill_description_primary,bullet_points_arr,related_col_arr

    def clena_bad_tags(self,soup):
        style_tags = soup.find_all()
        for style_tag in style_tags:
            # remove style attr
            del style_tag['class']
            del style_tag['data-mce-fragment']
            del style_tag['data-mce-selected']
            del style_tag['width']
            del style_tag['border']
            del style_tag['style']
            del style_tag['data-mce-style']
            del style_tag['data-sheets-value']
            del style_tag['src']
            del style_tag['alt']
            del style_tag['data-mce-src']


        return soup

    def remove_all_css_style(self,full_description_html_primary):

        # remove all style tags
        soup = bs(full_description_html_primary, 'html.parser')
        # find all tags

        full_description_html_primary = str(soup)
        return full_description_html_primary

    def request_link_by_link(self,link_by_item,proxy_index,s):
        # link_by_item = "http://web.archive.org/web/20210617082123/https://www.univers-fleuri.com/products/collier-fleur-resine"
        # print("request_link_by_link")

         # make request
        response_item = self.make_request(link_by_item, proxy_index, s)
        # print('converrt into soup html')



        # converrt into soup html
        # print(response_item)
        if response_item != False and response_item != None:
            soup_item = bs(response_item.text, 'html.parser')

            # remove bad tags
            print("clena_bad_tags")
            soup_item = self.clena_bad_tags(soup_item)



            # get attributes data-product
            try:
                product_data = soup_item.find('div', class_='product_form')
                product_data = product_data['data-product']
            except:
                product_data = soup_item.find('script', id='ProductJson-product-template')
                # remove script tag
                # get only value
                try:
                    product_data = str(product_data.text)
                except:
                    product_data = False

            if product_data != False:
                images_arr = []
                product_data = json.loads(product_data)
                # print(product_data)


                produc_id = product_data['id']
                product_title = product_data['title']
                product_handle = product_data['handle']

                published_at = product_data['published_at']
                created_at = product_data['created_at']
                vendor = product_data['vendor']
                product_type = product_data['type']
                tags = product_data['tags']
                variants_arr = []
                secure_url = ''
                full_description_html_primary = ''
                bullet_points_arr = []

                full_description = product_data['description']
                full_description = str(self.clena_bad_tags(bs(full_description, 'html.parser')))
                total_description_html_arr = full_description


                all_desc = bs(full_description, 'html.parser')

                full_description_html_primary = ''
                bullet_points_arr = []
                related_col_arr = []
                r = []
                related_col_arr.append(r)

                try:
                    # print('a')
                    full_description_html_primary, bullet_points_arr, related_col_arr = self.cut_full_description(soup_item,full_description)
                except Exception as e:
                    # display a line of error
                    print(e)
                    print("Error on line {}".format(sys.exc_info()[-1].tb_lineno))

                index = 0
                for h2 in all_desc.find_all('h2'):
                    if index == 0:
                        # remove style attr from h2
                        del h2['style']
                        del h2['class']
                        # remove all css style
                        h2_html_origin = str(h2)

                    h2.decompose()
                    index += 1
                    break


                full_description = bs(full_description, 'html.parser').text

                # get title from head
                ceo_title = soup_item.find('title').text
                ceo_title = ceo_title.strip()
                ceo_description = soup_item.find('meta', {'name': 'description'})['content']

                images = product_data['images']
                for img in images:
                    images_arr.append(img)


                title_html = soup_item.find('h1')

                # for div in soup_item.find_all('div', class_='swatch_options'):
                #     for variant in div.find_all('div', class_='option_title'):
                #         # if div has text, append to handle_arr
                #         if len(div.text) > 0:
                #             variants_arr.append(variant.text)
                data_value_list = []
                # if div has data-value, append to data_value_list

                id_by_id = product_data['id']
                vendor = product_data['vendor']
                type = product_data['type']
                tags = product_data['tags']

                for product in product_data['variants']:
                    if product['id'] not in self.dublicate:
                        try:
                            secure_url = ''
                            self.dublicate.append(product['id'])

                            self.id_by_id_arr.append(product['id'])
                            self.product_name_arr.append(product_title)
                            self.total_description_html_arr.append(str(self.remove_all_css_style(total_description_html_arr)))
                            self.clean_description_html_arr.append(self.remove_all_css_style(full_description_html_primary))
                            self.full_description_html_arr.append(self.remove_all_css_style(full_description_html_primary))
                            self.price_arr.append(self.cut_compare_price(product_data['price']))
                            self.price_min_arr.append(self.cut_compare_price(product_data['price_min']))
                            self.price_max_arr.append(self.cut_compare_price(product_data['price_max']))
                            self.full_link_arr.append(link_by_item)
                            self.data_value_list_arr.append(data_value_list)
                            self.variants_arr.append(','.join(variants_arr))
                            # title section
                            self.title_arr.append(product_title)
                            self.title_html_arr.append(title_html)
                            self.ceo_title_arr.append(ceo_title)
                            self.ceo_description_arr.append(ceo_description)
                            self.full_description_html_arr.append(full_description_html_primary)
                            self.bullet_points_arr.append(bullet_points_arr)
                            try:
                                secure_url = 'https:' + str(product['featured_image']['src'])
                            except:
                                secure_url = ''

                            variants = []
                            if product['option1'] == None:
                                variants.append('')
                            else:
                                variants.append(product['option1'])

                            if product['option2'] == None:
                                variants.append('')
                            else:
                                variants.append(product['option2'])

                            if product['option3'] == None:
                                variants.append('')
                            else:
                                variants.append(product['option3'])

                            self.related_collections_handle_arr.append(product_handle)
                            # print(related_col_arr)
                            # if len(related_col_arr[0]) == 0:
                            #     hh_arr = []
                            #     hh_arr.append(product_handle)
                            #     self.handle_arr.append(hh_arr)
                            # else:
                            #     self.handle_arr.append(related_col_arr)
                            hh = []
                            hh.append(product_handle)
                            # print(f"product_handle {product_handle}")
                            self.handle_arr.append(hh)
                            self.full_description_arr.append(full_description)

                            try:
                                self.imge_primary_arr.append(images_arr[0])
                            except:
                                self.imge_primary_arr.append('')

                            self.images_arr_variant.append(images_arr[0])
                            self.images_arr.append(','.join(images_arr))
                            self.variants_arr_primary.append(variants)
                            self.h2_html_arr.append(str(self.remove_all_css_style(h2_html_origin)))
                            self.product_id_arr.append(id_by_id)
                            self.tags_arr.append(','.join(tags))
                            self.vendor_arr.append(vendor)
                            self.type_arr.append(type)
                            self.secure_url_arr.append(secure_url)

                            self.published_at.append(product_data['published_at'])
                            self.created_at.append(product_data['created_at'])
                            self.available_arr.append(product_data['available'])
                            self.compare_at_price_varies_arr.append(product_data['compare_at_price_varies'])
                            self.price_varies_arr.append(product_data['price_varies'])

                            if str(product_data['compare_at_price']) != 'None':
                                self.compare_at_price_arr.append(self.cut_compare_price(product_data['compare_at_price']))
                            else:
                                self.compare_at_price_arr.append(product_data['compare_at_price'])

                            if product_data['compare_at_price_max'] == 'None':
                                self.compare_at_price_max_arr.append(product_data['compare_at_price_max'])
                            else:
                                self.compare_at_price_max_arr.append(self.cut_compare_price(product_data['compare_at_price_max']))
                        except Exception as e:
                            print(e)
                            print("Error on line {}".format(sys.exc_info()[-1].tb_lineno))



    def cut_compare_price(self,compare_at_price):

        if compare_at_price != 0:
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
            print(self.increment_index)
            sheet.cell(row=next_row, column=1, value="Product Quantity:"+str(self.increment_index))

            wb.save(xlsx_file_path)
            wb.close()


    def save_to_xlsx(self):
        xlsx_file_path = "shopify.xlsx"
        wb = load_workbook(xlsx_file_path)

        # Step 2: Select the worksheet where you want to append data
        sheet_name = "Sheet"  # Change this to the name of your target sheet
        sheet = wb[sheet_name]

        # Calculate the next row to append data to (assuming data starts from row 2)
        next_row = sheet.max_row + 1

        for i in range(len(self.id_by_id_arr)):
            try:
                option1 = self.variants_arr_primary[i][0]
            except:
                option1 = ''
            try:
                option2 = self.variants_arr_primary[i][1]
            except:
                option2 = ''
            try:
                option3 = self.variants_arr_primary[i][2]
            except:
                option3 = ''

            try:
                bullet_points_variant1 = self.remove_all_css_style(self.bullet_points_arr[i][0])
            except:
                bullet_points_variant1 = ''

            try:
                bullet_points_variant2 = self.remove_all_css_style(self.bullet_points_arr[i][1])
            except:
                bullet_points_variant2 = ''

            try:
                bullet_points_variant3 = self.remove_all_css_style(self.bullet_points_arr[i][2])
            except:
                bullet_points_variant3 = ''
            #
            # try:
            #     # remove all style attribute
            #     total_description_html = ''
            #     total_description_html = str(h2_html_arr[i][0]) + str(full_description_html_arr[i][0]) + str(bullet_points_variant1) + str(bullet_points_variant2) + str(bullet_points_variant3)
            #
            # except:
            #     total_description_html = ''

            # Step 4: Append the data to the selected worksheet
            # print(self.handle_arr[i])
            try:
                sheet.cell(row=next_row, column=1, value=str(self.id_by_id_arr[i]))
                sheet.cell(row=next_row, column=2, value=str(self.product_id_arr[i]))
                sheet.cell(row=next_row, column=3, value=str(self.full_link_arr[i]))
                sheet.cell(row=next_row, column=4, value=str(self.hanle))
                sheet.cell(row=next_row, column=5, value=str(self.hanle)) #
                sheet.cell(row=next_row, column=6, value=str(self.hanle)) # self.related_collections_handle_arr[i])
                sheet.cell(row=next_row, column=7, value=str(self.title_arr[i]))
                sheet.cell(row=next_row, column=8, value=str(self.title_html_arr[i]))
                sheet.cell(row=next_row, column=9, value=str(self.ceo_title_arr[i]))
                sheet.cell(row=next_row, column=10, value=str(self.ceo_description_arr[i]))
                sheet.cell(row=next_row, column=11, value=str(self.product_name_arr[i]))
                sheet.cell(row=next_row, column=12, value=str(self.full_description_arr[i]))
                sheet.cell(row=next_row, column=13, value=str(self.total_description_html_arr[i])).fill = PatternFill(start_color='ADFF2F', end_color='ADFF2F', fill_type='solid')
                sheet.cell(row=next_row, column=14, value=str(self.h2_html_arr[i])).fill = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
                sheet.cell(row=next_row, column=15, value=str(self.clean_description_html_arr[i])).fill = PatternFill(start_color='B0E0E6', end_color='B0E0E6', fill_type='solid')
                sheet.cell(row=next_row, column=16, value=str(bullet_points_variant1)).fill = PatternFill(start_color='CD853F', end_color='CD853F', fill_type='solid')
                sheet.cell(row=next_row, column=17, value=str(bullet_points_variant2)).fill = PatternFill(start_color='FFEBCD', end_color='FFEBCD', fill_type='solid')
                sheet.cell(row=next_row, column=18, value=str(bullet_points_variant3)).fill = PatternFill(start_color='FF7F50', end_color='FF7F50', fill_type='solid')
                sheet.cell(row=next_row, column=19, value=str(self.published_at[i]))
                sheet.cell(row=next_row, column=20, value=str(self.created_at[i]))
                sheet.cell(row=next_row, column=21, value=str(self.vendor_arr[i]))
                sheet.cell(row=next_row, column=22, value=str(self.type_arr[i]))
                sheet.cell(row=next_row, column=23, value=str(self.tags_arr[i]))
                sheet.cell(row=next_row, column=24, value=str(self.price_arr[i]))
                sheet.cell(row=next_row, column=25, value=str(self.price_min_arr[i]))
                sheet.cell(row=next_row, column=26, value=str(self.price_max_arr[i]))
                sheet.cell(row=next_row, column=27, value=str(self.available_arr[i]))
                sheet.cell(row=next_row, column=28, value=str(self.price_varies_arr[i]))
                sheet.cell(row=next_row, column=29, value=str(self.compare_at_price_arr[i]))
                sheet.cell(row=next_row, column=30, value=str(self.compare_at_price_max_arr[i]))
                sheet.cell(row=next_row, column=31, value=str(self.compare_at_price_varies_arr[i]))
                sheet.cell(row=next_row, column=32, value=str(''))
                sheet.cell(row=next_row, column=33, value=str(''))
                sheet.cell(row=next_row, column=34, value=str(self.remove_webarchive_from_img(self.images_arr[i])))
                sheet.cell(row=next_row, column=35, value=str(self.remove_webarchive_from_img(self.imge_primary_arr[i])))
                sheet.cell(row=next_row, column=36, value=str(self.variants_arr[i]))
                sheet.cell(row=next_row, column=37, value=str(option1))
                sheet.cell(row=next_row, column=38, value=str(option2))
                sheet.cell(row=next_row, column=39, value=str(option3))
                sheet.cell(row=next_row, column=40, value=str(self.remove_webarchive_from_img(self.secure_url_arr[i])))
                sheet.cell(row=next_row, column=41, value=str(''))
                sheet.cell(row=next_row, column=42, value=str(self.price_arr[i]))
                next_row += 1
            except:
                pass


        # Step 5: Save the changes back to the XLSX file
        wb.save(xlsx_file_path)

        # Optional: Close the workbook
        wb.close()
        self.increment_index += len(dict.fromkeys(self.product_id_arr))

        # clear
        # get all variables from __init__
        for name, value in vars(self).items():
            # clear all variables
            if type(value) is list :# and name != 'product_counter' and name != 'dublicate'
                value.clear()

    def remove_webarchive_from_img(self, img):
        if img != '':
            if self.webarchive == True:
                img_arr = img.split(',')
                for i,im in enumerate(img_arr):
                    end_pos = im.find('https://cdn')
                    img_arr[i] = im[end_pos:]

            return ",".join(img_arr)
        else:
            return img

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
        print(link)
        hendle = link.split('/')[-1]
        collection_handle = link.split('/')[-3]
        return hendle,collection_handle

    def request_link_by_link_to_get_ids(self,link,proxy,s):
        # request link
        response = self.make_request(link,proxy,s)
        if response != None:
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

    def make_request(self,url,proxy,s,reconnect=None):
        response = ''
        timeout = 20
        try:
            if reconnect == True:
                timeout = 80
            response = requests.get(url, timeout=timeout)
            # else:
            #     response = s.get(url, proxies=proxy, timeout=20)
            #     if response.status_code == 429 or response.status_code == 409:
            #         response = requests.get(url, timeout=20)
        except:
            print('Except')
            print(url)
            print(response)
            self.make_request(url,proxy,s,reconnect=True)

            # response = requests.get(url, timeout=20)
            # print(f"make_request {response.status_code}")
        try:
            if response.status_code == 200:
                return response
            if response.status_code == 404:
                return False
            else:
                time.sleep(40)
                return self.make_request(url,proxy,s)
        except:
            pass


    def cut_collection_link(self,link):
        if "/collections/" in link:
            col_p = link.find('/collections/')
            prod_pos = link.find('/products/')
            self.hanle = link[col_p+13:]
            print("====================================")
            print(self.hanle)
            print("====================================")

    def scrap_shopify(self,all_categpries):
        if self.webarchive == True:
            domain = self.webarchive_url_domain
        else:
            domain = self.domain

        index = 0
        for category in all_categpries:
            if "/collections/" in category:
                self.cut_collection_link(category)

                url = domain + category
                print(f"url for collection {url}")
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
                    response = requests.get(url, timeout=200, headers={'User-Agent': 'Mozilla/5.0'})
                    print("Skipping. Connnection error")

                # request link with proxy
                if response != False and response != None:

                    soup = bs(response.text, 'html.parser')

                    # find all 'a' with class product-info__caption and get href
                    for link in soup.find_all('a'): # , class_='product-info__caption'
                        link = link.get('href')
                        fill_link = ''

                        if link != None:

                            if self.webarchive == True:
                                # print(link)
                                if link.find("/products/") != -1 and "/collections/" not in link:
                                    if "web.archive" in link:
                                        # fill_link = self.cut_collection_link(link)
                                        fill_link = link
                                        print(f"fl {fill_link}")
                                    else:
                                        fill_link = domain + link
                                        print(f"fl {fill_link}")

                            if self.webarchive == False:
                                fill_link = domain + link
                                fill_link = ''

                            if len(fill_link) > 0:
                                # try:
                                    print(f"Link origin {fill_link}")
                                    self.request_link_by_link(fill_link,proxy,s)

                                    print("++++++++++++")
                                    print(f'INDEX {index}')
                                    print(f"ID COUNT {len(self.product_counter)}")
                                    print(f"handle = {self.hanle}")


                                    # remove duplicates from list
                                    print("++++++++++++")
                                    # if index % 2 == 0:
                                    try:
                                        self.save_to_xlsx()
                                    except Exception as e:
                                        print(e)
                                        print("Error on line {}".format(sys.exc_info()[-1].tb_lineno))
                                    index += 1
                                    print(f"Prim INDEX = {index}")


                                # except Exception as e:
                                #     print(e)
                                #     print("Error on line {}".format(sys.exc_info()[-1].tb_lineno))
                    #         if index == 30:
                    #             break
                    #     if index == 30:
                    #         break
                    # if index == 30:
                    #     break


                # try:
                #     self.save_to_xlsx()
                # except:
                #     pass
        self.save_to_xlsx()
        # try:
        #     time.sleep(15)
        #     self.save_to_xlsx_product_count()
        # except:
        #     pass


    def get_menu_links(self):
        url = ''
        if self.webarchive == True:
            url = self.webarchive_url+""+self.domain

        if self.webarchive == False:
            url = self.domain

        all_categpries = []
        response = requests.get(url,timeout=60)

        soup = bs(response.text, 'html.parser')
        # for link in soup.find('div', class_='nav nav--combined clearfix').find_all('a'):
        # for link in soup.find('div', class_='nav nav--combined center').find_all('a'):
        for link in soup.find('div', class_='grid-item text-center large--text-right').find_all('a'):
            menu_link = link.get('href')
            if menu_link.find('/collections') != -1:
                all_categpries.append(menu_link)

        return all_categpries


    def create_xls_file(self):
        # create shopify.xlsx file

        wb = Workbook()
        ws = wb.active
        ws.append(["id","product ID","full_link","handle","collection_handele","related_collections_handle","title","title_html","ceo_title","ceo_description","product_name","full_description","full_description_html","h2_html","description_html","bullet_points_html","bullet_points_html","bullet_points_html","published_at","created_at","vendor","type","tags","price","price_min","price_max","available","price_varies","compare_at_price","compare_at_price_max","compare_at_price_varies","requires_selling_plan","selling_plan_groups","images","featured_image","variants","option1","option2","option3","variant featured_image","variant compare_at_price","variant price"])
        wb.save("shopify.xlsx")

    def clean_duplicates(self):
        wb = load_workbook("shopify.xlsx")
        sheet = wb.worksheets[0]
        data_arr = []
        hendler_arr = []
        remove_indexs_arr = []
        hendler_origin_arr = []
        # Iterate over the rows in the sheet
        indexs = 0
        for row in sheet:
            ids = row[0].value

            if ids in data_arr:
                index_ids = data_arr.index(ids)
                print(f"Index {indexs}")
                print(f"ID {ids}")
                print(f"Duble value {data_arr[index_ids]}")
                print(f"Handle value {hendler_arr[index_ids]}")
                print(f"Dublicate index {index_ids}")
                print("Dublicate")
                print(f"ORIGIN")
                print()
                if str(row[3].value) != str(hendler_arr[index_ids]):
                    print(f"Related {hendler_arr[index_ids]  +','+ row[3].value}")
                    row[4].value = row[4].value.replace(hendler_arr[index_ids],"")

                    full_collection = hendler_arr[index_ids] + "," + row[3].value

                    row[4].value = full_collection
                    print(f"Real data {row[4].value}")

                remove_indexs_arr.append(ids)
            else:
                data_arr.append(str(ids))
                hendler_arr.append(str(row[3].value))
                hendler_origin_arr.append(str(row[3].value))
            indexs += 1


        wb.save("shopify.xlsx")
        wb.close()

        # remove dublicates using pandas
        df = pd.read_excel('shopify.xlsx')
        df.drop_duplicates(subset=['id'], inplace=True, keep='last')
        df.to_excel('shopify.xlsx', index=False)
        print("Done")


if __name__ == "__main__":
    shopify_scrapper = ShopifyScrapper()
    shopify_scrapper.webarchive = True
    shopify_scrapper.webarchive_url = "http://web.archive.org/web/20210920200301/"
    shopify_scrapper.webarchive_url_domain = "http://web.archive.org"

    shopify_scrapper.domain = "https://www.univers-fleuri.com"
    shopify_scrapper.create_xls_file()
    all_categpries = shopify_scrapper.get_menu_links()
    print(all_categpries)
    shopify_scrapper.scrap_shopify(all_categpries)
    shopify_scrapper.clean_duplicates()




