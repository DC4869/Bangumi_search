import json
import openpyxl as xl
import random
import requests
import time
from bs4 import BeautifulSoup

def get_bangumi_link(keyword, type, maxNumber):
    keyword = keyword.replace('!', ' ')
    keyword = keyword.replace('/', ' ')
    keyword = keyword.replace(' ', '-')
    url = f'https://api.bgm.tv/search/subject/{keyword}?type={type}&responseGroup=small&max_results={maxNumber}'
    headers = {
        'User-Agent': 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.11 (KHTML, like Gecko) Chrome/23.0.1271.64 Safari/537.11'
        #'Accept' : '*/*',
        #'Accept-Language': 'en-US,en,ja,zh,cn;q=0.8',
        #'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/98.0.4758.82',
    }
    print(url)
    time.sleep(random.uniform(0., 1.))
    response = requests.get(url, headers=headers)
    #print(response.status_code)
    if response.status_code == 200:
        if 'json' in response.headers.get('Content-Type'):
            html = response.json()
            if html['list'] is None:
                bgm_id = ''
                bgm_link = ''
                bgm_name = ''
                bgm_name_cn = ''
            else:
                bgm_id = html['list'][0]['id']
                bgm_link = html['list'][0]['url']
                bgm_name = html['list'][0]['name']
                bgm_name_cn = html['list'][0]['name_cn']
        else:
            #print('Response content is not in JSON format.')
            html = 'spam'
            bgm_id = ''
            bgm_link = ''
            bgm_name = ''
            bgm_name_cn = ''
    else:
        #print('Response content is not in JSON format.')
        html = 'spam'
        bgm_id = ''
        bgm_link = ''
        bgm_name = ''
        bgm_name_cn = ''

    return bgm_id, bgm_link, bgm_name, bgm_name_cn

#Years = ['2000', '2001', '2002', '2003', '2004', '2005', '2006', '2007', '2008', '2009', '2010', '2011', '2012', '2013', '2014', '2015', '2016', '2017', '2018', '2019', '2020', '2021', '2022', '2023', '2024']
Years = ['2013', '2014', '2015', '2016', '2017', '2018', '2019', '2020', '2021', '2022', '2023', '2024']
for year in Years:
    input_xlsx = 'Index/Index_{}.xlsx'.format(year)
    wb = xl.load_workbook(input_xlsx)
    sheets = wb.sheetnames
    for sheet_name in sheets:
        worksheet = wb[sheet_name]
        for row in range(3,worksheet.max_row+1):
            #for column in 'B':  #Here you can add or reduce the columns
            cell_name = '{}{}'.format('B', row)
            anime_name = worksheet[cell_name].value # the value of the specific cell
            print(anime_name) 
            bgm_id, bgm_link, bgm_name, bgm_name_cn = get_bangumi_link(anime_name, 2, 3)
            print(bgm_id, bgm_link, bgm_name, bgm_name_cn)
            
            bgm_id_cell = '{}{}'.format('I', row)
            bgm_link_cell = '{}{}'.format('J', row)
            bgm_name_cell = '{}{}'.format('K', row)
            bgm_name_cn_cell = '{}{}'.format('L', row)
        
            worksheet[bgm_id_cell] = bgm_id
            worksheet[bgm_link_cell] = bgm_link
            worksheet[bgm_name_cell] = bgm_name
            worksheet[bgm_name_cn_cell] = bgm_name_cn
    
    output_xlsx = 'output/Index_{}_bgm.xlsx'.format(year)
    wb.save(output_xlsx)
