import time
import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook

wb = Workbook()
ws = wb.active
ws.append(["Quote", "Author", "Tags"])

url = "https://quotes.toscrape.com/"
page_number = 1

while url:
    page = requests.get(url)
    page.encoding = "utf-8"
    soup = BeautifulSoup(page.text, "html.parser")

    for q in soup.select("div.quote"):
        text = q.select_one("span.text").text
        author = q.select_one("small.author").text
        tags = ", ".join(t.text for t in q.select("a.tag"))
        ws.append([text, author, tags])

    print(f"Done page {page_number}")
    page_number += 1

    next_btn = soup.select_one("li.next a")
    if next_btn:
        url = "https://quotes.toscrape.com" + next_btn["href"]
    else:
        url = None

    time.sleep(0.5)

wb.save("quotes.xlsx")
print("All done! quotes.xlsx is ready")