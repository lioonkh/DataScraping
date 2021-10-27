from bs4 import BeautifulSoup
import requests
import pandas as pd



links=['https://www.astrotools.com/product/1-2-extra-heavy-duty-reversible-air-drill-500rpm/',
       'https://www.astrotools.com/product/industrial-1-4-air-die-grinder/',
       'http://www.astrotools.com/1-2-air-ratchet-wrench-50ft-lb-torque.html',
       'http://www.astrotools.com/1-2-super-duty-impact-wrench-twin-hammer.html',
       'http://www.astrotools.com/1-heavy-duty-air-impact-wrench-with-2-anvil.html',
       'http://www.astrotools.com/1-heavy-duty-air-impact-wrench-with-6-anvil.html',
       'http://www.astrotools.com/onyx-heavy-duty-long-barrel-air-hammer-with-4pc-chisels-250mm.html',
       'http://www.astrotools.com/needle-scaler-flux-hammer-combo-4-400-blows-per-minute.html']
allproducts=[]
for index, url in enumerate(links) :
    page = requests.get(url).text
    sp = BeautifulSoup(page, 'html.parser')
    product_info = sp.findAll("div", id="main-content")
    allproducts.append(product_info)
productset= []

for product_info in allproducts:
    for spp in product_info:
        productdetails=dict()
        productdetails[ 'title' ]= spp.find('div','et_pb_module et_pb_wc_title et_pb_wc_title_0_tb_body et_pb_bg_layout_light').text.strip('\n\r\t": ').strip('\n\r\t": ').strip('\n\r\t": ')
        productdetails['number'] = spp.find('div', 'et_pb_text_inner').text.strip('\n\r\t": ').strip('\n\r\t": ').strip('\n\r\t": ')
        for div in spp.find_all('div', class_='et_pb_tab_content'):
            for p in div.find_all('p'):
                spec = p.text  # to get the specification text
        productdetails['spec']=spec.strip('\n\r\t": ').strip('\n\r\t": ').strip('\n\r\t": ')
        im = spp.find('img', class_='wp-post-image')
        productdetails['image'] = im['src']
        productset.append(productdetails)



df = pd.DataFrame(productset)
df.to_excel('AstroPr.xlsx')






























