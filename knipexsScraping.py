from bs4 import BeautifulSoup
import requests
import pandas as pd

headers= {'Useer_Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.45 Safari/537.36'}

products_list=[]
def getProducts(identifier):
   # the identifier in the url will change according to each product id
    url=f'https://www.knipex.com/products/crimping-pliers/knipex-multicrimp-crimping-pliers-with-changer-magazine/knipex-multicrimpcrimping-pliers-changer-magazine/{identifier}'
    r=requests.get(url,headers=headers)
    soup = BeautifulSoup(r.text, 'html.parser')

   # to scrape the technical details then append it to a list
    tech_details = []
    for i in soup.find_all('div', class_ ="key-value-class-item"):
        tech_details.append(i.text.replace("\n", "").strip().replace(" ", "-"))

    products_info = {
        'ProductTitle': soup.find('div', class_ ="ProductDetailTitle").text.strip('\n').split('\n')[1],
        'ProductNumber': soup.find('div', class_ ="ProductDetailTitle").text.strip('\n').split('\n')[0],
        'Image': 'https://www.knipex.com/' + soup.find('div', class_ ="SliderProductDetailContainer").findNext().findNext().findNext()['src'],
        'TechnicalDetails': tech_details,
        'Description': soup.find('div', class_ ="bulletpoint-container").findNext().findNext().text.strip('\n').replace("\n", "-----")
        }
    products_list.append(products_info)
    return

product_key=["973301","975110","975236","975304","975314","1262180","1382200",
"1640150","001101","169501SB","6801160","6801180","6801200","6801280","7001125","7001160"
,"7002160","7002180","7101200","7172460","8601250","8603125"
,"8603150"
,"8603180"
,"8603250"
,"8603400"
,"8701150"
,"8701180"
,"8701250"
,"8801250"
,"8801300"
,"4811J1"
,"4811J2"
,"4821J21"
,"4821J31"
,"4911A2"
,"4921A21"
,"2502160"
,"2612200"
,"2616200"
,"2622200"
,"0301180"
,"0302180"
,"0902240"
,"9852"
,"9855"
,"002120"
,"0306180"
,"0306200"
,"1106160"
,"1386200"
,"7006160"
,"7006180"
,"9516165"
,"9516200"
,"9511165"
,"9512165"
,"9512200"
,"9531250"
]

for x in product_key:
    getProducts(x)
    print(x + " DONE!")


df=pd.DataFrame(products_list)
df.to_excel('KnipexDataset.xlsx', index=False)
print('Finished')
