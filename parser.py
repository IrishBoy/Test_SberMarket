import pandas as pd
import xml.etree.ElementTree as ET
from datetime import datetime as dt
from openpyxl import load_workbook


xml_file = 'feed-20200210_sh.xml'
outpup_file = 'Data.xlsx'
tree = ET.parse(xml_file)
root = tree.getroot()
columns = ['external_offer_id', 'offer_name', 'offer_items_per_pack',
           'offer_pack_type', 'offer_vat', 'retailer_regular_price_per_item',
           'retailer_regular_price_per_kilo',
           'retailer_regular_price_per_pack',
           'shop_stock', 'price_update_date',
           'stock_update_date', 'retailer_id']
dataFrame = pd.DataFrame(columns=columns)
weight = ['Весовой']
fas = ['Кусок']
date_change = (dt.strptime(root.attrib['date'], '%Y-%m-%d %H:%M')).date()
items_amount = 1
vat = 20
for node in root.findall('shop/offers/offer'):
    offer_id = node.get('id')
    name = node.find('name').text
    shop_price = float(node.find('price').text)
    pack_type = 'Поштучно'
    for item in node.iterfind('param[@name="Фасовка"]'):
        if item.text in weight:
            pack_type = "Весовой"
        elif item.text in fas:
            pack_type = "Фасованный"
    for item in node.iterfind('param[@name="weight"]'):
        pack_weight = float(item.text)
    for retailer in node.iterfind('outlets/outlet'):
        retailer_id = retailer.get("id")
        retailer_stock = int(retailer.get("instock"))
        retailer_price = float(retailer.find('price').text)
        if ((pack_type in ['Весовой', 'Фасованный'] or
                abs(retailer_price - shop_price) >= max(retailer_price, shop_price) * 0.3)) and (retailer_price != 0):
    #здесь очень странный момент, очевидно(как мне кажется), что наценка не может быть 30%
    #однако, есть товары, в которых тип фасовки не указан, однако они продаются на развес,
    #напрмер, Азу или же какая-то разделка, к сожалению, мне не удалось найти схожести в их параметрах
    #поэтому я говорю, что если цена отличается более чем на 30%(беря в рассчёт, что максимальная порция
    # не превышает 700 грамм), значит, это цена уже за киллограмм и я также сразу меняю тип фасовки
            retailer_price_kilo = retailer_price / pack_weight
            if pack_type == 'Поштучно':
                pack_type = 'Весовой'
        else:
            retailer_price_kilo = 0
        dataFrame = dataFrame.append(pd.Series([
                                                offer_id, name, items_amount,
                                                pack_type, vat, retailer_price,
                                                retailer_price_kilo,
                                                retailer_price,
                                                retailer_stock, date_change,
                                                date_change, retailer_id],
                                                index=columns),
                                                ignore_index=True)
book = load_workbook(outpup_file)
writer = pd.ExcelWriter(outpup_file, engine='openpyxl', mode='a')
writer.book = book
for ret_id in dataFrame.retailer_id.unique():
    cur_ret = pd.DataFrame(dataFrame[dataFrame['retailer_id'] == ret_id])
    cur_ret = cur_ret.drop(['retailer_id'], axis=1)
    cur_ret.to_excel(writer, sheet_name=ret_id, index=False)
writer.save()
writer.close()
