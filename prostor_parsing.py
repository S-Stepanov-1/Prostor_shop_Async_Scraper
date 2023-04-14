import random
import time
import asyncio
import requests
import xlsxwriter as xlsxwriter
from aiocfscrape import CloudflareScraper
from bs4 import BeautifulSoup
from fake_useragent import UserAgent

URL_katalog = "https://prostor.ua/ru/katalog/"
URL_root = "https://prostor.ua"
ua = UserAgent(browsers=["chrome", "firefox"])


def get_soup(url, params=None):
    headers = {"User-Agent": ua.random}

    response = requests.get(url, headers=headers, params=params)
    return BeautifulSoup(response.content, "lxml")


async def get_data(card_url):
    headers = {"User-Agent": ua.random}

    async with CloudflareScraper(headers=headers) as session:
        async with session.get(card_url) as response:
            card_soup = BeautifulSoup(await response.text(), "lxml")  # Soup of each cards

            try:
                name = card_soup.find("h1", class_="product-title").text.strip()
            except Exception:
                name = "No name"
                print(f"No name\t{card_url}")

            try:
                article = card_soup.find("div", class_="product-header__code").text.strip()
            except Exception:
                article = "No article"
                print(f"No article\t{card_url}")

            try:
                price = card_soup.find("div", class_="product-price__item").text.strip()
            except Exception:
                price = "No price"
                print(f"No price\t{card_url}")


            try:
                image_link = card_soup.find("span", class_="gallery__link j-gallery-zoom j-gallery-link").get("data-href")
            except Exception:
                image_link = "Нет изображений"

            try:
                description = card_soup.find("div", class_="product-description j-product-description").text.strip()
            except Exception:
                description = "Нет описания для данного товара"

            if card_soup.find("div", class_="product-header__availability product-header__availability--out-of-stock"):
                return "No more products"

            return {
                "name": name,
                "article": article,
                "price": price,
                "image_link": URL_root + image_link,
                "description": description
            }


async def create_tasks(page_soup):
    article_cards = page_soup.find_all("div", class_="catalogCard-box j-product-container")

    tasks = []
    for card in article_cards:
        card_link = card.find_next("a").get("href")
        task = asyncio.create_task(get_data(URL_root + card_link))
        tasks.append(task)

    return await asyncio.gather(*tasks)


def write_to_file(page_xlsx, row, info_dict, column=0):
    try:
        for product in info_dict:
            page_xlsx.write(row, column, product["name"])
            page_xlsx.write(row, column + 1, product["article"])
            page_xlsx.write(row, column + 2, product["price"])
            page_xlsx.write(row, column + 3, product["image_link"])
            page_xlsx.write(row, column + 4, product["description"])
            row += 1
    except Exception:
        pass


def main():
    cur_time = time.perf_counter()

    book = xlsxwriter.Workbook("Prostor_shop.xlsx")  # Creation of xlsx book
    page_xlsx = book.add_worksheet("Products")
    page_xlsx.set_column(0, 5, 15)

    # Format settings
    sheet_format = book.add_format()
    sheet_format.set_bold()

    page_xlsx.write(0, 0, "Наименование", sheet_format), page_xlsx.write(0, 1, "Артикул", sheet_format)
    page_xlsx.write(0, 2, "Цена", sheet_format), page_xlsx.write(0, 3, "Изображение", sheet_format)
    page_xlsx.write(0, 4, "Описание", sheet_format)
    row = 1

    main_soup = get_soup(URL_katalog)
    pages_num = int(main_soup.find_all("a", class_="pager__item j-catalog-pagination-btn")[-1].text.strip())

    for page in range(1, pages_num + 1):
        if int(page) % 10 == 0:
            delay = random.randint(7, 12)
            print(f"\nWaiting {delay} seconds...\n")
            time.sleep(delay)

        page_soup = get_soup(URL_katalog, params={"page": page})
        print(page_soup.find("title").text.strip())  # Page number + title(name)

        products_info = asyncio.run(create_tasks(page_soup))

        write_to_file(page_xlsx, row, products_info)
        row += 20

        if "No more products" in products_info:
            break

    book.close()

    end = time.perf_counter()
    print(f"It took {end - cur_time} seconds")


if __name__ == '__main__':
    main()
