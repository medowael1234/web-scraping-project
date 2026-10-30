import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook

url = "https://books.toscrape.com/"
page = requests.get(url)
page.encoding = "utf-8"

soup = BeautifulSoup(page.text, "html.parser")

wb = Workbook()
ws = wb.active
ws.append(["اسم الكتاب", "السعر"])

for book in soup.select("article.product_pod"):
    title = book.h3.a["title"]
    price = book.select_one(".price_color").text
    ws.append([title, price])

wb.save("books.xlsx")
print("خلصت! شوف ملف books.xlsx")