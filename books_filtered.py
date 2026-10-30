import time
import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook

wb = Workbook()
ws = wb.active
ws.append(["Title", "Price", "Rating", "Stock", "Link"])

for n in range(1, 51):
    url = f"https://books.toscrape.com/catalogue/page-{n}.html"
    page = requests.get(url)
    page.encoding = "utf-8"
    soup = BeautifulSoup(page.text, "html.parser")

    for book in soup.select("article.product_pod"):
        title = book.h3.a["title"]
        price = float(book.select_one(".price_color").text.replace("£", ""))
        rating = book.select_one("p.star-rating")["class"][1]
        stock = book.select_one(".availability").text.strip()
        link = "https://books.toscrape.com/catalogue/" + book.h3.a["href"]
        if price < 20 and rating == "Five":
            ws.append([title, price, rating, stock, link])

    print(f"Done page {n}")
    time.sleep(0.5)

wb.save("books_filtered.xlsx")
print("All done! books_filtered.xlsx is ready")