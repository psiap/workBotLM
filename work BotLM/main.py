import lxml
import requests
# establishing session
from bs4 import BeautifulSoup

s = requests.Session()
s.headers.update({
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10.9; rv:45.0) Gecko/20100101 Firefox/45.0'
    })

def load_user_data(session):
    url = 'https://inled.ru/magazin/folder/prozhektory-machtovye-dlya-sportivnyh-sooruzhenij'
    request = session.get(url)
    return request.text

def contain_movies_data(text):
    soup = BeautifulSoup(text)
    print(soup)
    film_list = soup.xpath('.//div[@class = "product-name"]/a/@href')[0]
    print(film_list)
    return film_list is not None

# loading files
while True:
    data = load_user_data(s)
    if contain_movies_data(data):
        with open('test.html', 'w+',encoding='utf-8') as output_file:
          output_file.write(data)
    else:
        break