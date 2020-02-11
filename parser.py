import pandas as pd
import xml.etree.ElementTree as ET
from datetime import datetime as dt
from openpyxl import load_workbook
import xlsxwriter


pd.set_option('display.max_columns', 500)

xml_file = 'feed-20200210_sh.xml'
tree = ET.parse(xml_file)
root = tree.getroot()
columns = ['external_offer_id', 'offer_name', 'offer_items_per_pack',
           'offer_pack_type', 'offer_vat', 'retailer_regular_price_per_item',
           'retailer_regular_price_per_kilo',
           'retailer_regular_price_per_pack',
           'shop_stock', 'price_update_date', 'stock_update_date', 'retailer_id']
dataFrame = pd.DataFrame(columns=columns)
weight = ['Весовой', 'Кусок', 'Нарезка']
date_change = (dt.strptime(root.attrib['date'], '%Y-%m-%d %H:%M')).date()
items_amount = 1
vat = 20
for node in root.findall('shop/offers/offer'):
    offer_id = node.get('id')
    name = node.find('name').text
    shop_price = node.find('price').text
    pack_type = 'Поштучно'
    for item in node.iterfind('param[@name="Фасовка"]'):
        if item.text in weight:
            pack_type = "Весовой"
    for item in node.iterfind('param[@name="weight"]'):
        pack_weight = float(item.text)
    for retailer in node.iterfind('outlets/outlet'):
        retailer_id = retailer.get("id")
        retailer_stock = retailer.get("instock")
        retailer_price = float(retailer.find('price').text)
        if pack_type == 'Весовой':
                retailer_price_kilo = retailer_price / pack_weight
        else:
            retailer_price_kilo = 0
        dataFrame = dataFrame.append(pd.Series([offer_id, name, items_amount,
                                                pack_type, vat, retailer_price,
                                                retailer_price_kilo,
                                                retailer_price,
                                                retailer_stock, date_change, date_change, retailer_id],
                                                index=columns), ignore_index=True)
book = load_workbook('Data.xlsx')
writer = pd.ExcelWriter('Data.xlsx', engine='openpyxl', mode='a')
writer.book = book
for ret_id in dataFrame.retailer_id.unique():
    cur_ret = pd.DataFrame(dataFrame[dataFrame['retailer_id'] == ret_id])
    cur_ret.drop(['retailer_id'], axis=1)
    cur_ret.to_excel(writer, sheet_name=ret_id, index=False)
writer.save()
writer.close()
