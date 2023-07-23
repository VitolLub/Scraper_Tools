from google_images_search import GoogleImagesSearch
import random
import requests
import base64
import sys
import json
import names
from playwright.sync_api import sync_playwright
from dataclasses import dataclass, asdict, field
import pandas as pd
import argparse

random_name = ['John', 'Jane', 'Jack', 'Jill', 'Joe', 'Jenny', 'Jerry', 'Judy', 'Jim', 'Jen', 'John', 'Jane', 'Jack', 'Jill', 'Joe', 'Jenny', 'Jerry', 'Judy', 'Jim', 'Jen', 'John', 'Jane', 'Jack', 'Jill', 'Joe', 'Jenny', 'Jerry', 'Judy', 'Jim', 'Jen', 'John', 'Jane', 'Jack', 'Jill', 'Joe', 'Jenny', 'Jerry', 'Judy', 'Jim', 'Jen', 'John', 'Jane', 'Jack', 'Jill', 'Joe', 'Jenny', 'Jerry', 'Judy', 'Jim', 'Jen', 'John', 'Jane', 'Jack', 'Jill', 'Joe', 'Jenny', 'Jerry', 'Judy', 'Jim', 'Jen', 'John', 'Jane', 'Jack', 'Jill', 'Joe', 'Jenny', 'Jerry', 'Judy', 'Jim', 'Jen', 'John', 'Jane', 'Jack', 'Jill', 'Joe', 'Jenny', 'Jerry', 'Judy', 'Jim', 'Jen' ]
random_surname = ['Doe', 'Smith', 'Jones', 'Taylor', 'Williams', 'Brown', 'Davies', 'Evans', 'Wilson', 'Thomas']
random_phone = ['+1 123 456 7890', '+1 234 567 8901', '+1 345 678 9012', '+1 456 789 0123', '+1 567 890 1234', '+1 678 901 2345', '+1 789 012 3456', '+1 890 123 4567', '+1 901 234 5678', '+1 012 345 6789','+1 123 456 7890', '+1 234 567 8901', '+1 345 678 9012', '+1 456 789 0123', '+1 567 890 1234', '+1 678 901 2345', '+1 789 012 3456', '+1 890 123 4567', '+1 901 234 5678', '+1 012 345 6789','+1 123 456 7890', '+1 234 567 8901', '+1 345 678 9012', '+1 456 789 0123', '+1 567 890 1234', '+1 678 901 2345', '+1 789 012 3456', '+1 890 123 4567', '+1 901 234 5678', '+1 012 345 6789']
gender_val = ['M','W']

# create JSON formatted data with 1000 random records
def create_json_data():
    json_data = []
    for i in range(101):
        try:
            random_email = random.choice(random_name).lower() + '.' + random.choice(random_surname).lower() + '@gmail.com'
            json_data.append({
                'name': names.get_full_name(),
                'email': random_email,
                'phone': gen_phone(),
                # 'websites':'',
                #'gender':random.choice(gender_val)
                'address':get_address(),
                'rend_typr':get_gend_type()
            })
        except Exception as e:
            print(e)
    return json_data

def get_address():
    # generate random address in USA
    from faker import Faker
    fake = Faker()
    return fake.address()
def get_gend_type():
    arr1 = ['Sale','Rent','Sale/Rant']

    return random.choice(arr1)
def get_gend_type():
    arr1 = ['Sale','Rent','Sale/Rant']

    return random.choice(arr1)
def gen_phone():
    first = str(random.randint(100, 999))
    second = str(random.randint(1, 888)).zfill(3)

    last = (str(random.randint(1, 9998)).zfill(4))
    while last in ['1111', '2222', '3333', '4444', '5555', '6666', '7777', '8888']:
        last = (str(random.randint(1, 9998)).zfill(4))

    return '+1 {}-{}-{}'.format(first, second, last)
def save_json_data():
    with open('data2.json', 'w') as outfile:
        try:
            json.dump(create_json_data(), outfile)
        except Exception as e:
            print(e)

def save_csv_data():
    res = create_json_data()
    print(res)
    # read JSON liny by line and save to CSV
    with open('sofi_zillow_data.csv', 'w') as outfile:
        for line in res:
            outfile.write(line['name'] + ',' + line['email'] + ',' + line['phone'] + ',' + line['address'] + ',' + line['rend_typr']+ '\n')


if __name__ == "__main__":
    # save in
    save_csv_data()
