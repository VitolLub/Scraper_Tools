import time
from bs4 import BeautifulSoup as bs
import requests
from openpyxl import Workbook
import json
import random

class ShopifyScrapper:

    def __init__(self):
        self.url = ''
        self.id_by_id_arr = []
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
    def request_link_by_link(self,link,proxy_index,s):
        print(link)
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
            tags_arr = []
            # get div with class description
            full_description = soup.find('div', class_='description').text
            full_description_html = soup.find('div', class_='description')

            # find all strong tags
            for strong in full_description_html.find_all('strong'):
                print(strong.text)
                tag = strong.text
                if tag.find('@') == -1:
                    tags_arr.append(strong.text)

        except:
            full_description = ''
            full_description_html = ''


        try:
            title = soup.find('h1',class_="product_name").text
            title_html = soup.find('h1',class_="product_name")
        except:
            title = ''
            title_html = ''

        try:
            # get title from head
            ceo_title = soup.find('title').text
            ceo_description = soup.find('meta', {'name': 'description'})['content']
        except:
            ceo_title = ''
            ceo_description = ''
        images_arr = []
        try:
            images = soup.find('div', class_='product_gallery_nav').find_all('img')
            if len(images) > 0:
                for image in images:
                    print(self.https+image['src'])
                    if image['src'].find('_300x.') > -1:
                        # replace all _300x. to _800x.
                        images_arr.append(self.https+image['src'].replace('_300x.', '_800x.'))
                    else:
                        images_arr.append(self.https+image['src'])
            else:
                images = soup.find('div', class_='image__container').find_all('img')
                for image in images:
                    print(f"real images {self.https+image['data-src']}")
                    if image['data-src'].find('_300x.') > -1:
                        # replace all _300x. to _800x.
                        images_arr.append(self.https+image['data-src'].replace('_300x.', '_800x.'))
                    else:
                        images_arr.append(self.https+image['data-src'])
        except:
            images_arr = []

        try:
            # get selected option
            variants = soup.find('option', selected='selected').text
            if variants.find('/') > 0:
                variants = variants.split('/')
            else:
                variants = variants.split(' ')
        except:
            variants = []

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
            related_collections = soup.find('div', class_='product-links').find_all('a')
            print("related_collections")
            print(related_collections)
            # get href
            for related_collections_a in related_collections:
                related_collections_href = related_collections_a['href']
                hendle_pos = related_collections_href.find('/collections/')
                related_collections_handle = related_collections_href[hendle_pos + 13:]
                related_collections_arr.append(related_collections_handle)
        except:
            related_collections_arr = []


        return product_name, price,link,",".join(data_value_list), full_description, full_description_html,title,title_html,ceo_title,ceo_description,images_arr,variants,bullet_points_arr,h2_html,id_by_id,tags_arr, vendor, type,related_collections_arr
    def save_to_xlsx(self,id_by_id_arr,product_name_arr, price_arr,full_link_arr,
                    data_value_list_arr,variants_arr,related_collections_handle_arr,handle_arr,
                    full_description_arr,full_description_html_arr,title_arr,title_html_arr,ceo_title_arr,
                    ceo_description_arr,images_arr,imge_primary_arr,variants_arr_primary,bullet_points_arr,h2_html_arr,product_id_arr,tags_arr,
                     vendor_arr,type_arr):
        wb = Workbook()
        ws = wb.active
        ws.append(["id","product ID","full_link","handle","collection_handele","related_collections_handle","title","title_html","ceo_title","ceo_description","product_name","full_description","full_description_html","h2_html","bullet_points_html","published_at","created_at","vendor","type","tags","price","price_min","price_max","available","price_varies","compare_at_price","compare_at_price_max","compare_at_price_varies","requires_selling_plan","selling_plan_groups","images","featured_image","variants","option1","option2","option3","variant featured_image","variant compare_at_price","variant price"])
        for i in range(len(product_name_arr)):
            try:
                option1 = variants_arr_primary[i][0]
                print(f"option1: {option1}")
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
            ws.append([str(id_by_id_arr[i]),str(product_id_arr[i]),str(full_link_arr[i]),str(handle_arr[i]),str(related_collections_handle_arr[i]),str(related_collections_handle_arr[i]),str(title_arr[i]),str(title_html_arr[i]),str(ceo_title_arr[i]),str(ceo_description_arr[i]),str(product_name_arr[i]), str(full_description_arr[i]),str(full_description_html_arr[i]),str(h2_html_arr[i]),str(bullet_points_arr[i]),'','',str(vendor_arr[i]),str(type_arr[i]),str(tags_arr[i]),str(price_arr[i]),str(price_arr[i]),str(price_arr[i]),"TRUE","0","","","0","","",str(images_arr[i]),str(imge_primary_arr[i]),str(variants_arr[i]),option1,option2,option3,str(images_arr[i])," ",str(price_arr[i])])
        wb.save("shopify.xlsx")
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
        # for div in soup.find_all('div', class_='swatch_options'):
        #     for data_id in div.find_all('div', {'data-id': True}):
        #         print(data_id['data-id'])
        #         id_arr.append(data_id['data-id'])
        #     # find all div with class option_title
        #     for variant in div.find_all('div', class_='option_title'):
        #         # if div has text, append to handle_arr
        #         if len(div.text) > 0:
        #             variants_arr.append(variant.text)
        #
        # print(id_arr)
        # print(variants_arr)
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
        domain = 'https://univers-chinois.com'
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
                print(proxy_index)
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
                print(f"link: {fill_link}")
                id_arr,variants_arr = self.request_link_by_link_to_get_ids(fill_link,proxy,s)
                print(id_arr)
                for id in id_arr:
                    full_link = fill_link + '?variant=' + str(id)
                    print(full_link)
                    product_name, price,full_link,data_value_list,full_description,full_description_html,title,title_html,ceo_title,ceo_description,images_arr,variants,bullet_points_arr,h2_html,id_by_id,tags_arr,vendor, type, related_collections_arr = self.request_link_by_link(full_link,proxy,s)
                    print(product_name, price,full_link,data_value_list)
                    print(id_by_id,id)

                    self.id_by_id_arr.append(id)
                    self.product_name_arr.append(product_name)
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
                    print(f"handle: {related_collections_arr}")
                    self.related_collections_handle_arr.append(','.join(related_collections_arr))
                    self.handle_arr.append(handle)
                    self.full_description_arr.append(full_description)
                    self.full_description_html_arr.append(full_description_html)
                    print(f"images_arr: {images_arr}")
                    try:
                        self.imge_primary_arr.append(images_arr[0])
                    except:
                        self.imge_primary_arr.append('')
                    self.images_arr.append(','.join(images_arr))
                    self.variants_arr_primary.append(variants)
                    try:
                        # bullet_points_arr to string
                        bullet_points_arr = ','.join(bullet_points_arr)
                        # print(f"bullet_points_arr: {len(bullet_points_arr)}")
                        self.bullet_points_arr.append(bullet_points_arr)
                    except:
                        self.bullet_points_arr.append('')
                    self.h2_html_arr.append(h2_html)
                    self.product_id_arr.append(id_by_id)
                    self.tags_arr.append(','.join(tags_arr))
                    self.vendor_arr.append(vendor)
                    self.type_arr.append(type)

                    print("++++++++++++")
                    print(f'INDEX {index}, {variants}')
                    print("++++++++++++")
                    index += 1
            #         if index == 300:
            #             break
            #
            #     if index == 300:
            #         break
            # if index == 300:
            #     break

        self.save_to_xlsx(self.id_by_id_arr,self.product_name_arr, self.price_arr,self.full_link_arr,
                         self.data_value_list_arr,self.variants_arr,self.related_collections_handle_arr,
                         self.handle_arr,self.full_description_arr,self.full_description_html_arr,
                         self.title_arr,self.title_html_arr,self.ceo_title_arr,self.ceo_description_arr,
                          self.images_arr,self.imge_primary_arr,self.variants_arr_primary,
                          self.bullet_points_arr,self.h2_html_arr,self.product_id_arr,self.tags_arr,
                          self.vendor_arr,self.type_arr)





    def get_menu_links(self):
        url = self.domain
        all_categpries = []
        response = requests.get(url)
        soup = bs(response.text, 'html.parser')
        for link in soup.find('div', class_='main-nav__wrapper').find_all('a'):
            menu_link = link.get('href')
            if menu_link.startswith('/collections'):
                all_categpries.append(menu_link)

        return all_categpries


if __name__ == "__main__":
    shopify_scrapper = ShopifyScrapper()
    all_categpries = shopify_scrapper.get_menu_links()
    print(all_categpries)
    shopify_scrapper.scrap_shopify(all_categpries)
