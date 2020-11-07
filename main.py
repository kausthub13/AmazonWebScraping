from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.keys import Keys
from tkinter import *
import time
from datetime import date
import mmap
from random import randint
import csv
import pandas as pd
import tkinter as tk
from tkinter import filedialog
import tkinter.font as font
import os
import ntpath

error_occurred = False
filename = None
row_count = 0
total_lines = 0
flipkart_val = 0
amazon_val = 0
directory = None
output_file = None


def UploadAction(event=None):
    global directory
    directory = filedialog.askdirectory()


def mapcount(filename):
    f = open(filename, "r+", encoding='utf8')
    buf = mmap.mmap(f.fileno(), 0)
    lines = 0
    readline = buf.readline
    while readline():
        lines += 1
    return lines


def read_csv(file):
    isbn_list = []
    with open(file, encoding='utf8') as csv_file:
        csv_reader = csv.reader(csv_file, delimiter=',')
        line_count = 0
        for row in csv_reader:
            if line_count == 0:
                pass
            else:
                isbn_list.append(row[1])
            line_count += 1
    return isbn_list


def next_isbn(filename):
    # with open(filename, 'r', encoding='utf8') as input_file:
    #     for current_row in csv.reader(input_file):
    #         yield current_row[1]
    csv_file = pd.read_csv(filename)
    isbn_list = csv_file['ISBN13']
    for i in range(len(isbn_list)):
        yield str(isbn_list.iloc[i])


def setup_flipkart_driver():
    options = Options()
    options.headless = False
    # options.add_argument('headless')
    options.add_argument('--no-sandbox')
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    chrome_driver_path = r"chromedriver.exe"
    driver = webdriver.Chrome(executable_path=chrome_driver_path, options=options)
    driver.minimize_window()
    return driver


def setup_flipkart_header(filename):
    global output_file
    base_name = ntpath.basename(filename)[:-4] + '_flipkart_output.csv'
    base_file_path = os.path.join(output_file, base_name)
    writing_file = open(base_file_path, 'a', encoding="utf-8", newline='')
    amazon_flip_writer = csv.writer(writing_file, delimiter=',')
    amazon_flip_writer.writerow(
        ["Date", "ISBN13", "Flipkart Listing","Buybox"])
    writing_file.close()


def flipkart_scraper(filename):
    global row_count
    global total_lines
    global output_file
    isbn_generator = next_isbn(filename)
    # next(isbn_generator)
    setup_flipkart_header(filename)
    driver = setup_flipkart_driver()
    base_file_path = ""
    line_count = 0
    for current_isbn in isbn_generator:
        line_count += 1
        found = False
        listed = "Not Listed"
        buybox = "No"
        # price = "NA"
        # curr_binding = "NA"
        # curr_pages = "NA"
        # curr_per_page_ratio = "NA"
        try:
            driver.get('https://www.flipkart.com/books/pr?sid=bks&q='+str(current_isbn))
            search_result = driver.find_element_by_class_name('_2cLu-l')
            search_result_link = search_result.get_attribute("href")
            driver.get(search_result_link)
            seller_name = driver.find_element_by_id("sellerName")
            time.sleep(2)
            # highlights = driver.find_elements_by_class_name('_2-riNZ')
            # for highlight in highlights:
            #     curr_highlight = str(highlight.text)
            #     if 'Binding' in highlight.text:
            #         curr_binding = curr_highlight.replace('Binding:','').strip()
            #     if 'Pages' in highlight.text:
            #         curr_pages = curr_highlight.replace('Pages:','').strip()
            if 'repro' in str(seller_name.text).lower():
                found = True
                buybox = "Yes"
                # price = driver.find_element_by_css_selector("div[class='_1vC4OE _3qQ9m1']").text
                # price = str(price).replace('₹','').strip()
                # try:
                #     curr_per_page_ratio = round(float(price) / float(curr_pages),2)
                # except:
                #     pass
            else:
                more_sellers = None
                try:
                    more_sellers = driver.find_element_by_class_name('VXac4C')
                except:
                    pass
                if more_sellers:
                    a_tag = more_sellers.find_element_by_tag_name('a')
                    href_link = a_tag.get_attribute("href")
                    driver.get(href_link)
                    time.sleep(2)
                    sellers = driver.find_elements_by_css_selector("div[class='_3fm_R4']")
                    seller_id = -1
                    if sellers:
                        for seller in sellers:
                            seller_id += 1
                            if 'repro' in str(seller.text).lower():
                                found = True
                                break
                        # if found:
                        #     prices = driver.find_elements_by_css_selector("div[class='_1vC4OE']")
                        #     prices_id = -1
                        #     for price_tag in prices:
                        #         prices_id += 1
                        #         if prices_id == seller_id:
                        #             price = price_tag.text
                        #             break
                        #     price = str(price).replace('₹','').strip()
                        #     try:
                        #         curr_per_page_ratio = round(float(price) / float(curr_pages),2)
                        #     except:
                        #         pass
            if found:
                listed = "Listed"
        except NoSuchElementException:
            pass
        except Exception as e:
            with open(ntpath.basename(filename)[:-4] + "flip_error_log.csv", 'a', newline='') as error_file:
                error_writer = csv.writer(error_file)
                error_writer.writerow(['Flipkart', current_isbn])
        print(str(line_count) + "/" + str(total_lines), "Flipkart", current_isbn, listed,buybox)
        base_name = ntpath.basename(filename)[:-4] + '_flipkart_output.csv'
        base_file_path = os.path.join(output_file, base_name)
        writing_file = open(base_file_path, 'a', encoding="utf-8", newline='')
        amazon_flip_writer = csv.writer(writing_file, delimiter=',')
        amazon_flip_writer.writerow([date.today(), current_isbn] + [listed,buybox])
        writing_file.close()
    driver.close()
    driver.quit()
    return base_file_path


def amazon_scrape(filename):
    global row_count
    global total_list
    global output_file
    options = Options()

    found = False
    listed = "Not Listed"
    price_tag = "NA"
    options.headless = False

    binding = "NA"
    per_page_ratio = "NA"
    # options.add_argument('headless')
    options.add_argument('--no-sandbox')
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    chrome_driver_path = r"chromedriver.exe"
    driver = webdriver.Chrome(executable_path=chrome_driver_path, options=options)
    driver.minimize_window()
    original_link = 'https://www.amazon.in/'
    base_name = ntpath.basename(filename)[:-4] + '_amazon_output.csv'
    base_file_path = os.path.join(output_file, base_name)
    writing_file = open(base_file_path, 'a', encoding='utf8', newline='')
    amazon_flip_writer = csv.writer(writing_file, delimiter=',')
    amazon_flip_writer.writerow(
        ["Date", "ISBN13", "Amazon Listing","Buybox"])
    writing_file.close()
    csv_file = pd.read_csv(filename)
    isbn_list = csv_file['ISBN13']
    line_count = 1
    for i in range(len(isbn_list)):
        found = False
        listed = "Not Listed"
        # price_tag = "NA"
        # start = time.time()
        # pages = "NA"
        current_isbn = str(isbn_list[i])
        buybox = "No"
        ## stack = [original_link]
        # curr_per_page_ratio = "NA"
        # count = 0
        # page = 1
        # time.sleep(randint(1, 5))
        try:
            driver.get('https://www.amazon.in/s?k='+current_isbn)
            # driver.get(stack.pop())
            # time.sleep(3)
            # isbn_field = driver.find_element_by_id('twotabsearchtextbox')
            if True:
            #     isbn_field.send_keys(str(current_isbn))
            #     isbn_field.send_keys(Keys.RETURN)
            #
            #     while driver.execute_script('return document.readyState;') != "complete":
            #         pass

                time.sleep(2)
                search_results = driver.find_elements_by_css_selector("a[class='a-size-base a-link-normal']")
                all_reference_links = []
                for search_result in search_results:
                    link = search_result.get_attribute("href")
                    if 'Paperback' in search_result.text or 'Hardcover' in search_result.text:
                        all_reference_links.append(link)
                other_results = driver.find_elements_by_css_selector(
                    "a[class='a-size-base a-link-normal a-text-bold']")
                for other_result in other_results:
                    link = other_result.get_attribute("href")
                    if 'Paperback' in other_result.text or 'Hardcover' in other_result.text:
                        all_reference_links.append(link)
                first_link = True
                while all_reference_links:
                    try:
                        curr_found = False
                        curr_listed = "Not Listed"
                        # price_tag = "NA"
                        # pages = "NA"
                        driver.get(all_reference_links.pop(0))
                        time.sleep(2)
                        # if first_link:
                        #     first_link = False
                        #     try:
                        #         the_details_id = driver.find_element_by_id('detailBullets_feature_div')
                        #         span_elements = the_details_id.find_elements_by_class_name('a-list-item')
                        #         for span_element in span_elements:
                        #             other_spans = span_element.find_elements_by_tag_name('span')
                        #             for other_span in other_spans:
                        #                 if 'pages' in other_span.text or 'Pages' in other_span.text:
                        #                     pages = other_span.text
                        #     except NoSuchElementException:
                        #         pass
                        # temp_binding = ""
                        # try:
                        #     binding = driver.find_element_by_id('productSubtitle')
                        #     if 'Hardcover' in binding.text:
                        #         temp_binding = 'Hardcover'
                        #     elif 'Paperback' in binding.text:
                        #         temp_binding = 'Paperback'
                        # except NoSuchElementException:
                        #     pass
                        seller_name = None
                        try:
                            # seller_name = driver.find_element_by_id('sellerProfileTriggerId')
                            delay = 5
                            seller_name = WebDriverWait(driver, delay).until(
                                EC.presence_of_element_located((By.ID, 'sellerProfileTriggerId')))
                        except Exception as e:
                            with open(ntpath.basename(filename)[:-4]+"amazon_error_log.csv", 'a',newline='') as error_file:
                                error_writer = csv.writer(error_file)
                                error_writer.writerow(['Amazon', current_isbn])
                        # try:
                        #     mrp_prices = driver.find_elements_by_class_name('a-list-item')
                        #     for mrp_price in mrp_prices:
                        #         if 'M.R.P.:' in mrp_price.text:
                        #             price_tag = str(mrp_price.text)
                        #             price_tag = price_tag.replace('M.R.P.:','')
                        # except e:
                        #     pass

                        # try:
                        #     selling_price = driver.find_element_by_css_selector(
                        #         "span[class='a-size-medium a-color-price inlineBlock-display offer-price a-text-normal price3P']")
                        # except NoSuchElementException:
                        #     pass
                        # if selling_price and price_tag == "NA":
                        #     price_tag = selling_price.text

                        # pages = pages.replace('pages', '').replace('Pages', '').strip()
                        # price_tag = price_tag.replace('₹', '').strip()
                        # try:
                        #     curr_per_page_ratio = round(float(price_tag) / float(pages), 2)
                        # except:
                        #     pass

                        if seller_name and 'repro' in str(seller_name.text).lower():
                            curr_found = True
                            found = True
                            buybox = "Yes"
                            if curr_found:
                                curr_listed = 'Listed'
                            # selling_price = None

                            base_name = ntpath.basename(filename)[:-4] + '_amazon_output.csv'
                            base_file_path = os.path.join(output_file, base_name)
                            writing_file = open(base_file_path, 'a', encoding="utf-8", newline='')
                            amazon_flip_writer = csv.writer(writing_file, delimiter=',')
                            amazon_flip_writer.writerow(
                                [date.today(), current_isbn] + [curr_listed,buybox])
                            writing_file.close()
                            print(str(line_count) + "/" + str(total_lines), "Amazon", current_isbn, curr_listed,buybox)
                        else:
                            try:
                                other_sellers = driver.find_element_by_id('mbc-olp-link')
                                a_tag = other_sellers.find_element_by_tag_name("a")
                                href_link = a_tag.get_attribute("href")
                                driver.get(href_link)
                                time.sleep(2)
                                try:
                                    delay = 5
                                    WebDriverWait(driver, delay).until(EC.presence_of_element_located(
                                        (By.CSS_SELECTOR, "span[class='a-size-medium a-text-bold']")))
                                except Exception as e:
                                    with open(ntpath.basename(filename)[:-4] + "amazon_error_log.csv", 'a',
                                              newline='') as error_file:
                                        error_writer = csv.writer(error_file)
                                        error_writer.writerow(['Amazon', current_isbn])
                                more_pages = True
                                while more_pages:
                                    more_pages = False
                                    search_results = driver.find_elements_by_css_selector(
                                        "span[class='a-size-medium a-text-bold']")
                                    seller_num = -1
                                    for search_result in search_results:
                                        seller_num += 1
                                        seller_name = search_result.text
                                        if 'repro' in str(seller_name).lower():
                                            curr_found = True
                                            curr_listed = 'Listed'
                                            found = True
                                            break
                                    # selling_prices = driver.find_elements_by_css_selector("span[class='a-size-large a-color-price olpOfferPrice a-text-bold']")
                                    if curr_found:
                                        #     curr_sell_num = -1
                                        #     for selling_price in selling_prices:
                                        #         curr_sell_num += 1
                                        #         if curr_sell_num == seller_num:
                                        #             price_tag = selling_price.text
                                        #             break
                                        #     pages = pages.replace('pages', '').replace('Pages', '').strip()
                                        #     price_tag = price_tag.replace('₹', '').strip()
                                        #     try:
                                        #         curr_per_page_ratio = round(float(price_tag) / float(pages),2)
                                        #     except:
                                        #         pass
                                        base_name = ntpath.basename(filename)[:-4] + '_amazon_output.csv'
                                        base_file_path = os.path.join(output_file, base_name)
                                        writing_file = open(base_file_path, 'a', encoding="utf-8", newline='')
                                        amazon_flip_writer = csv.writer(writing_file, delimiter=',')
                                        amazon_flip_writer.writerow(
                                            [date.today(), current_isbn] + [curr_listed,buybox])
                                        writing_file.close()
                                        print(str(line_count) + "/" + str(total_lines), "Amazon", current_isbn,
                                              curr_listed,buybox)
                                    try:
                                        next_button = driver.find_element_by_css_selector('li[class="a-last"]')
                                        next_button_link = next_button.find_element_by_tag_name('a')
                                        href_tag = next_button_link.get_attribute('href')
                                        more_pages = True
                                        driver.get(href_tag)
                                    except NoSuchElementException:
                                        pass

                            except NoSuchElementException:
                                pass
                            except Exception as e:
                                with open(ntpath.basename(filename)[:-4] + "amazon_error_log.csv", 'a',
                                          newline='') as error_file:
                                    error_writer = csv.writer(error_file)
                                    error_writer.writerow(['Amazon', current_isbn])
                            if curr_found == False:
                                try:
                                    new_sellers = driver.find_element_by_id("buybox-see-all-buying-choices-announce")
                                    new_sellers_link = new_sellers.get_attribute("href")
                                    driver.get(new_sellers_link)
                                    time.sleep(2)
                                    try:
                                        delay = 5
                                        WebDriverWait(driver, delay).until(EC.presence_of_element_located(
                                            (By.CSS_SELECTOR, "span[class='a-size-medium a-text-bold']")))
                                    except Exception as e:
                                        with open(ntpath.basename(filename)[:-4] + "amazon_error_log.csv", 'a',
                                                  newline='') as error_file:
                                            error_writer = csv.writer(error_file)
                                            error_writer.writerow(['Amazon', current_isbn])
                                    more_pages = True
                                    while more_pages:
                                        more_pages = False
                                        search_results = driver.find_elements_by_css_selector(
                                            "span[class='a-size-medium a-text-bold']")
                                        seller_num = -1
                                        for search_result in search_results:
                                            seller_num += 1
                                            seller_name = search_result.text
                                            if 'repro' in str(seller_name).lower():
                                                curr_found = True
                                                curr_listed = 'Listed'
                                                found = True
                                                break
                                        # selling_prices = driver.find_elements_by_css_selector("span[class='a-size-large a-color-price olpOfferPrice a-text-bold']")
                                        if curr_found:
                                            #     curr_sell_num = -1
                                            #     for selling_price in selling_prices:
                                            #         curr_sell_num += 1
                                            #         if curr_sell_num == seller_num:
                                            #             price_tag = selling_price.text
                                            #             break
                                            #     pages = pages.replace('pages', '').replace('Pages', '').strip()
                                            #     price_tag = price_tag.replace('₹', '').strip()
                                            #     try:
                                            #         curr_per_page_ratio = round(float(price_tag) / float(pages),2)
                                            #     except:
                                            #         pass
                                            base_name = ntpath.basename(filename)[:-4] + '_amazon_output.csv'
                                            base_file_path = os.path.join(output_file, base_name)
                                            writing_file = open(base_file_path, 'a', encoding="utf-8", newline='')
                                            amazon_flip_writer = csv.writer(writing_file, delimiter=',')
                                            amazon_flip_writer.writerow(
                                                [date.today(), current_isbn] + [curr_listed,buybox])
                                            writing_file.close()
                                            print(str(line_count) + "/" + str(total_lines), "Amazon", current_isbn,
                                                  curr_listed,buybox)
                                        try:
                                            next_button = driver.find_element_by_css_selector('li[class="a-last"]')
                                            next_button_link = next_button.find_element_by_tag_name('a')
                                            href_tag = next_button_link.get_attribute('href')
                                            more_pages = True
                                            driver.get(href_tag)
                                        except NoSuchElementException:
                                            pass
                                except:
                                    pass
                    except NoSuchElementException:
                        pass
                    except Exception as e:
                        with open(ntpath.basename(filename)[:-4] + "amazon_error_log.csv", 'a', newline='') as error_file:
                            error_writer = csv.writer(error_file)
                            error_writer.writerow(['Amazon', current_isbn])
        except Exception as e:
            with open(ntpath.basename(filename)[:-4] + "amazon_error_log.csv", 'a', newline='') as error_file:
                error_writer = csv.writer(error_file)
                error_writer.writerow(['Amazon', current_isbn])
        if not found:
            base_name = ntpath.basename(filename)[:-4] + '_amazon_output.csv'
            base_file_path = os.path.join(output_file, base_name)
            writing_file = open(base_file_path, 'a', encoding="utf-8", newline='')
            amazon_flip_writer = csv.writer(writing_file, delimiter=',')
            amazon_flip_writer.writerow([date.today(), current_isbn] + [listed,buybox])
            writing_file.close()
            # print(row_count,"Amazon",current_isbn, listed,str(price_tag))
            print(str(line_count) + "/" + str(total_lines), "Amazon", current_isbn, "Not Listed",buybox)
        line_count += 1
    driver.close()
    driver.quit()
    return base_file_path


def merge_files():
    amazon_file = open('amazon_bookstore.csv', 'r')
    amazon_reader = csv.reader(amazon_file)
    line_num = 0
    final_file = dict()
    for row in amazon_reader:
        final_file.setdefault(line_num, [])
        final_file[line_num].extend(row[0:4])
        line_num += 1
    amazon_file.close()
    flipkart_file = open('flipkart_bookstore.csv', 'r')
    flipkart_reader = csv.reader(flipkart_file)
    line_num = 0
    for row in flipkart_reader:
        final_file[line_num].extend(row[2:4])
        line_num += 1
    flipkart_file.close()
    for i in range(line_num):
        amazon_flipkart_file = open('amazon_flipkart_bookstore.csv', 'a', newline='')
        af_writer = csv.writer(amazon_flipkart_file)
        if final_file[i] != []:
            af_writer.writerow(final_file[i])
        amazon_flipkart_file.close()


def setup_ui():
    global row_count
    global output_file
    global total_lines
    global flipkart_val
    global amazon_val
    root = tk.Tk()
    root.geometry("600x200")
    root.title("Select Your Folder To Check Whether The Titles are Listed in Amazon and Flipkart")
    select = tk.Label(root)
    select.pack()
    myFont = font.Font(size=20)
    button = tk.Button(root, bg='yellow', text='Select Folder', command=UploadAction, font=myFont)
    button.configure(width=600, height=2)
    button.pack()
    CheckVar1 = IntVar()
    CheckVar2 = IntVar()
    C1 = Checkbutton(root, text="Flipkart", variable=CheckVar1, onvalue=1, offvalue=0, height=2, width=20)
    C2 = Checkbutton(root, text="Amazon", variable=CheckVar2, onvalue=1, offvalue=0, height=2, width=20)
    C1.pack()
    C2.pack()
    label = tk.Label(root)
    label.pack()

    def task():
        if directory:
            select.config(text="You have successfully selected the folder you can now close this window")
            label.config(text="Current Selected Folder Path: " + str(directory))
        if CheckVar1.get():
            global flipkart_val
            flipkart_val = 1
        if CheckVar2.get():
            global amazon_val
            amazon_val = 1
        root.after(1000, task)

    root.after(0, task)
    root.mainloop()
    if directory:
        output_file = os.path.join(directory, 'Output')
        try:
            os.mkdir(output_file)
        except FileExistsError:
            pass
        for name in os.listdir(directory):
            if name[-5:] == '.xlsx' or name[-4:] == '.xls':
                global total_records

                filename = os.path.join(directory, name)
                excel_reader = pd.read_excel(filename)
                filename = os.path.join(directory, name[:-5] + '.csv')
                excel_reader.to_csv(filename, index=None, header=True)
                total_lines = mapcount(filename) - 1
                print("Processing the file now")
                if amazon_val:
                    file_path = amazon_scrape(filename)
                    read_file = pd.read_csv(file_path)
                    if os.path.exists(file_path[:-4] + '.xlsx'):
                        existing_file = pd.read_excel(file_path[:-4] + ".xlsx")
                        read_file = pd.concat([existing_file, read_file], ignore_index=True)
                        with pd.ExcelWriter(file_path[:-4] + '.xlsx', mode='w') as writer:
                            read_file.to_excel(writer, sheet_name='Sheet1', index=None, header=True)
                    else:
                        with pd.ExcelWriter(file_path[:-4] + '.xlsx', mode='w') as writer:
                            read_file.to_excel(writer, sheet_name='Sheet1', index=None, header=True)
                    os.remove(file_path)
                if flipkart_val:
                    # flipkart_scrape(filename)
                    file_path = flipkart_scraper(filename)
                    read_file = pd.read_csv(file_path)
                    if os.path.exists(file_path[:-4] + '.xlsx'):
                        existing_file = pd.read_excel(file_path[:-4] + ".xlsx")
                        read_file = pd.concat([existing_file, read_file], ignore_index=True)
                        with pd.ExcelWriter(file_path[:-4] + '.xlsx', mode='w') as writer:
                            read_file.to_excel(writer, sheet_name='Sheet1', index=None, header=True)
                    else:
                        with pd.ExcelWriter(file_path[:-4] + '.xlsx', mode='w') as writer:
                            read_file.to_excel(writer, sheet_name='Sheet1', index=None, header=True)
                    os.remove(file_path)
                os.remove(filename)
                print("All Files Processed")
        for name in os.listdir(os.path.dirname(sys.argv[0])):
            if name[-4:] == '.csv':
                convert_to_excel = pd.read_csv(name)
                with pd.ExcelWriter(name[:-4] + '.xlsx', mode='w') as writer:
                    convert_to_excel.to_excel(writer, sheet_name='Sheet1', index=None, header=True)




setup_ui()
