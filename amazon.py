import csv
from time import sleep
from datetime import datetime
from random import random
from selenium.common import exceptions
from selenium import webdriver
import pandas as pd

def generate_filename(search_term):
    timestamp = datetime.now().strftime("%Y%m%d%H%S%M")
    stem = '_'.join(search_term.split(' '))
    filename = stem + '_' + timestamp + '.csv'
    return filename

def save_data_to_csv(record, filename, new_file=False):
    header = ['description', 'price', 'rating', 'review_count', 'url']
    if new_file:
        with open(filename, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(header)
    else:
        with open(filename, 'a+', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(record)

def create_webdriver() -> webdriver.Edge:
    options = webdriver.EdgeOptions()
    options.use_chromium = True
    options.headless = True
    driver = webdriver.Edge(executable_path=r'C:\Users\msedgedriver\msedgedriver.exe', options=options)
    return driver

def generate_url(search_term, page):
    base_template = 'https://www.amazon.com/s?k={}&ref=nb_sb_noss_1'
    search_term = search_term.replace(' ', '+')
    stem = base_template.format(search_term)
    url_template = stem + '&page={}'
    if page == 1:
        return stem
    else:
        return url_template.format(page)

def extract_card_data(card):
    description = card.find_element_by_xpath('.//h2/a').text.strip()
    url = card.find_element_by_xpath('.//h2/a').get_attribute('href')
    try:
        price = card.find_element_by_xpath('.//span[@class="a-price-whole"]').text
    except exceptions.NoSuchElementException:
        return
    try:
        temp = card.find_element_by_xpath('.//span[contains(@aria-label, "out of")]')
        rating = temp.get_attribute('aria-label')
    except exceptions.NoSuchElementException:
        rating = ""
    try:
        temp = card.find_element_by_xpath('.//span[contains(@aria-label, "out of")]/following-sibling::span')
        review_count = temp.get_attribute('aria-label')
    except exceptions.NoSuchElementException:
        review_count = ""
    return description, price, rating, review_count, url

def collect_product_cards_from_page(driver):
    cards = driver.find_elements_by_xpath('//div[@data-component-type="s-search-result"]')
    return cards

def sleep_for_random_interval():
    time_in_seconds = random() * 2
    sleep(time_in_seconds)

def run(search_term):
    filename = generate_filename(search_term)
    save_data_to_csv(None, filename, new_file=True)
    driver = create_webdriver()
    num_records_scraped = 0

    for page in range(1, 21):
        search_url = generate_url(search_term, page)
        print(search_url)
        driver.get(search_url)
        driver.set_page_load_timeout(30)  # Set the timeout to 30 seconds (adjust as needed)

        cards = collect_product_cards_from_page(driver)
        for card in cards:
            record = extract_card_data(card)
            if record:
                save_data_to_csv(record, filename)
                num_records_scraped += 1
        sleep_for_random_interval()

    driver.quit()
    print(f"Scraped {num_records_scraped:,d} for the search term: {search_term}")

    # Read CSV into a DataFrame
    df = pd.read_csv(filename)

    # Save DataFrame to Excel file
    excel_filename = filename.replace('.csv', '.xlsx')
    df.to_excel(excel_filename, index=False)
    print(f'Data successfully saved to {excel_filename}')

if __name__ == '__main__':
    term = 'dell laptop'
    run(term)
