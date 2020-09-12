import requests
import xlsxwriter
from selenium import webdriver
from bs4 import BeautifulSoup
from time import sleep
from unicodedata import normalize


#https://stackoverflow.com/questions/20986631/how-can-i-scroll-a-web-page-using-selenium-webdriver-in-python
def scroll_to_bottom(driver):
    SCROLL_PAUSE_TIME = 1

    # Get scroll height
    last_height = driver.execute_script("return document.body.scrollHeight")

    while True:
        # Scroll down to bottom
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

        # Wait to load page
        sleep(SCROLL_PAUSE_TIME)

        # Calculate new scroll height and compare with last scroll height
        new_height = driver.execute_script("return document.body.scrollHeight")
        if new_height == last_height:
            break
        last_height = new_height


def find_all_links():
        urls = set()
        try:
            driver = webdriver.Firefox()
            url = "https://bullionexchanges.com/silver-products-page"
            driver.get(url)
            sleep(2)

            scroll_to_bottom(driver)

            product_list = driver.find_element_by_class_name("all-products-by-metal")
            links = product_list.find_elements_by_tag_name("a")
            for link in links:
                url = link.get_attribute("href")
                urls.add(url)
        finally:
            driver.close()
            print(f"Found {len(urls)} total links")
            return urls


def extract_relevant_data(links, workbook):
    worksheet = workbook.get_worksheet_by_name("silver")
    possible_metals = ["Gold", "Silver", "Platinum"]
    max_col_used = 3

    count = 0
    for link in links:
        count += 1
        print(f"Examining link {count}: {link}")
        sleep(1)
        try:
            html = requests.get(link)
            soup = BeautifulSoup(html.content, 'html.parser')

            title = soup.find(class_="col-main").find(class_="product-name").get_text().strip()

            metals = []
            for metal in possible_metals:
                if metal in title:
                    metals.append(metal)
            metals = ", ".join(metals)

            description = normalize('NFKC',
                                soup.find(class_="box-collateral box-description box-active").get_text().strip())

            images = []
            carousel = soup.find("div", {"class": "list_carousel"}).find_all("a")
            for a in carousel:
                images.append(a["href"])
        except:
            print("Link failed: " + link)
            continue

        current_data = [title, metals, description, images]

        col_used = add_data_to_worksheet(current_data, worksheet, count)
        if col_used > max_col_used:
            max_col_used = col_used

    return max_col_used


def add_data_to_worksheet(data, worksheet, row):
    cur_col = -1

    for num in range(3):
        cur_col += 1
        value = data[num]
        worksheet.write(row, cur_col, value)

    images = data[3]

    for image in images:
        cur_col += 1
        worksheet.write(row, cur_col, image)

    return cur_col


def initialize_workbook():
    workbook = xlsxwriter.Workbook('bullionexchanges_silver.xlsx')
    worksheet = workbook.add_worksheet("silver")

    bold = workbook.add_format({"bold": True})

    worksheet.write(0, 0, "Title", bold)
    worksheet.write(0, 1, "Metal", bold)
    worksheet.write(0, 2, "Description", bold)
    worksheet.freeze_panes(1, 0)

    return workbook


def close_workbook(workbook, max_col_used=3):
    worksheet = workbook.get_worksheet_by_name("silver")
    bold = workbook.add_format({"bold": True})

    letter = chr(ord('A') + max_col_used)
    cols_to_change_width = "A:" + letter
    worksheet.set_column(cols_to_change_width, 40)

    if max_col_used >= 3:
        for num in range(3, max_col_used+1):
            cur_url = "URL_" + str(num - 2)
            worksheet.write(0, num, cur_url, bold)

    workbook.close()



if __name__ == '__main__':
    max_col_used = 0
    workbook = initialize_workbook()
    links = find_all_links()
    if len(links) > 0:
        max_col_used = extract_relevant_data(links, workbook)
    close_workbook(max_col_used=max_col_used, workbook=workbook)

