from selenium.webdriver import ActionChains
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium import webdriver
from time import sleep
from selenium.webdriver.support.wait import WebDriverWait
import openpyxl

product_name = ''
product_price = ''
product_price_rubly = ''
product_status = ''
link_checker = True

current_url = ''




def open_google_tab(product_name):
    driver.get('https://www.google.com/')
    driver.find_element_by_name('q').send_keys(product_name)
    driver.find_element_by_name('q').send_keys(u'\ue007')
    cite_list = driver.find_elements_by_tag_name('cite')
    tester = 0
    driver.implicitly_wait(3)
    for cite in cite_list:
        if "www.ikea.com" in cite.text:
            cite.click()
            tester = 1
            break

    global product_price_rubly
    global product_status
    global link_checker
    if tester == 1 and link_checker is True:
        driver.implicitly_wait(1)
        try:
            button = driver.find_element_by_class_name(u'js-cookie-info__accept-button')
            driver.implicitly_wait(1)
            ActionChains(driver).move_to_element(button).click(button).perform()
        except Exception:
            print("НЕТУ КУКИ")
        driver.implicitly_wait(3)
        try:
            product_price_rubly = driver.find_element_by_class_name('range-revamp-pip-price-package__main-price').text
            driver.implicitly_wait(1)
        except:
            print('Нет цены')
        try:
            product_price_rubly = driver. \
                find_element_by_xpath(
                '//*[@id="content"]/div/div/div/div[2]/div[3]/div/div[1]/div/div[2]/div/span/span[1]').text
            driver.implicitly_wait(1)
        except:
            print('Нет цены')
        try:
            product_price_rubly = driver. \
                find_element_by_xpath(
                '//*[@id="content"]/div/div/div/div[2]/div[3]/div/div[1]/div/div[2]/div').text
            driver.implicitly_wait(1)
        except:
            print('Нет цены')

        try:
            driver.find_element_by_xpath('//*[@id="content"]/div/div[1]/div/div[2]/div[3]/div/div[5]/div[2]').click()
        except Exception:
            print('Другая ссылка1')
        try:
            driver.find_element_by_link_text('Проверка наличия в офлайн-магазине').click()
        except:
            print('Другая ссылка2')

        driver.implicitly_wait(3)
        try:
            button1 = driver.find_element_by_class_name('range-revamp-change__search-store')
            ActionChains(driver).move_to_element(button1).send_keys('омск').perform()

            if driver.find_element_by_class_name('range-revamp-stockcheck__store-text').text == "В наличии":
                product_status = "В наличии"
            elif driver.find_element_by_class_name('range-revamp-stockcheck__store-text').text == "Заканчивается":
                product_status = "Заканчивается"
            elif driver.find_element_by_class_name('range-revamp-stockcheck__store-text').text == "Почти закончился":
                product_status = "Почти закончился"
            else:
                product_status = "Нет в наличии"
        except:
            link_checker = False
            product_status = ' '
            product_price_rubly = ' '
            print("Не нашел ссылку")
    else:
        product_price_rubly = ''
        product_status = ''
    link_checker = True


def links_in_each_page():
    driver.implicitly_wait(3)
    elements = driver.find_elements_by_class_name('item-card__name-link')
    page_urls = [element.get_attribute('href') for element in elements]
    driver.implicitly_wait(1)
    global current_url
    global index
    current_url = driver.current_url
    for url in page_urls:
        driver.get(url)
        global product_name
        product_name = driver.find_element_by_class_name('item__heading').text
        prod_name = product_name
        global product_price
        try:
            product_price = driver.find_element_by_class_name('item__price-once').text
        except:
            product_price = 'Нет в наличии на каспи'

        product_name = product_name + ' ikea омск'
        open_google_tab(product_name)

        print(product_name, product_price, product_price_rubly, product_status)
        worksheet['A' + str(index)] = prod_name
        worksheet['B' + str(index)] = product_price
        worksheet['C' + str(index)] = product_price_rubly
        worksheet['D' + str(index)] = product_status
        index = index + 1
    driver.get(current_url)


xpath_next_page = input('Xpath кнокпи следующей страницы: ')
url = input('Url подкатегории: ')
phone_number = input('Номер телефона без 8ки: ')
kaspi_password = input('Пароль от каспи: ')
excel_name = input('Название эксель файла: ')

wb = openpyxl.load_workbook(excel_name + '.xlsx')
wb.create_sheet('Sheet1')
worksheet = wb['Sheet1']
index = 1

driver = webdriver.Chrome(executable_path="chromedriver.exe")
driver.get("https://kaspi.kz/entrance?action=login&returnUrl=https%3A%2F%2Fkaspi.kz%2Fshop%2F%3Fat%3D3")
driver.find_element_by_xpath('//*[@id="txtLogin"]').send_keys(phone_number)
driver.find_element_by_xpath('//*[@id="txtPassword"]').send_keys(kaspi_password)

try:
    driver.find_element_by_link_text('Алматы').click()
except:
    print("Уже Алмата")

driver.get(url)

try:
    driver.find_element_by_link_text('Алматы').click()
except:
    print("Уже Алмата")


def loop(xpath_next_page, current_url):
    while True:
        next_page_btn = driver.find_element_by_xpath(xpath_next_page)
        next_page_checker = next_page_btn.get_attribute('class')
        if 'pagination__el _disabled' == next_page_checker:
            sleep(1)
            links_in_each_page()
            break
        elif driver.current_url == current_url:
            break
        else:
            sleep(1)
            links_in_each_page()
            WebDriverWait(driver, 10). \
                until(
                EC.element_to_be_clickable(
                    (By.XPATH, xpath_next_page))).click()
            print(current_url)


loop(xpath_next_page, url)
wb.save(excel_name+'.xlsx')
driver.close()
