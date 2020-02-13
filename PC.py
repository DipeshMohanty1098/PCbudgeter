import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import time
import os
import win32
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException
import pandas as pd
import xlsxwriter
from selenium.webdriver.chrome.options import Options
from tkinter.ttk import *
import threading


root = tk.Tk()
titles = []
prices = []
parts_dict = {'PARTS':[],'PRICE':[]}
driver_path = 'C:/Users/eradi/Desktop/Webdriver/chromedriver'


chrome_options = Options()
chrome_options.add_argument("--headless")
chrome_options.add_argument("--window-size=1366,768");



def calculate():
    cpu = a.get()
    titles.clear()
    prices.clear()
    parts_dict['PARTS'].clear()
    parts_dict['PRICE'].clear()
    driver = webdriver.Chrome(executable_path=driver_path, options = chrome_options)
    pg = Progressbar(root, length=200, orient='horizontal', maximum=60, value=0, mode='determinate')
    pg.grid(row = 11,column = 1)
    pg['value'] = 0
    root.update()
    for i in range(7):
        root.update()
        time.sleep(0.05)
        tk.Label(root,text = 'Extracting Prices {}/7'.format(i+1)).grid(row = 10, column = 1)
        if i == 0:
            root.update()
            if a.get() == '':
                pg['value'] = i*10
                pg.update()
                continue
            else:
                root.update()
                pg['value'] = i*10
                pg.update()
                driver.get("https://amazon.in")
                searchb = driver.find_element_by_xpath(
                    '/html/body/div[1]/header/div/div[1]/div[3]/div/form/div[2]/div/input')
                searchb.click()
                search_box = driver.find_element_by_xpath(
                    '/html/body/div[1]/header/div/div[1]/div[3]/div/form/div[3]/div[1]/input')
                search_box.send_keys(cpu)
                search_box.send_keys(Keys.RETURN)
                SEARCH_RESULT_LINK = driver.find_element_by_xpath("(//div[@class='sg-col-inner']//img[contains(@data-image-latency,'s-product-image')])[1]")
                SEARCH_RESULT_LINK.click()
                driver.switch_to.window(driver.window_handles[0])
                driver.close()
                driver.switch_to.window(driver.window_handles[0])
                price = driver.find_element_by_id("priceblock_ourprice").text
                product_title = driver.find_element_by_id("productTitle").text
                titles.append(product_title)
                root.update()
                w = tk.Label(root, text = "{}: {}".format(product_title,price) ) #.grid(row = 10,columnspan = 1)
                root.update()
                w.grid(row = 10,columnspan = 1)
                shaved_price = price[2:len(price)-3]
                int_price = str(shaved_price).replace(',','')
                prices.append(int(int_price))
                parts_dict['PARTS'].append(product_title)
                parts_dict['PRICE'].append(int_price)
                root.update()
        elif i == 1:
            if b.get() == '':
                pg['value'] = i * 10
                pg.update()
                continue
            else:
                pg['value'] = i * 10
                pg.update()
                driver.get("https://amazon.in")
                searchb = driver.find_element_by_xpath(
                    '/html/body/div[1]/header/div/div[1]/div[3]/div/form/div[2]/div/input')
                searchb.click()
                search_box = driver.find_element_by_xpath(
                    '/html/body/div[1]/header/div/div[1]/div[3]/div/form/div[3]/div[1]/input')
                search_box.send_keys(b.get())
                search_box.send_keys(Keys.RETURN)
                SEARCH_RESULT_LINK = driver.find_element_by_xpath(
                    "(//div[@class='sg-col-inner']//img[contains(@data-image-latency,'s-product-image')])[1]")
                SEARCH_RESULT_LINK.click()
                driver.switch_to.window(driver.window_handles[0])
                driver.close()
                driver.switch_to.window(driver.window_handles[0])
                price = driver.find_element_by_id("priceblock_ourprice").text
                product_title = driver.find_element_by_id("productTitle").text
                titles.append(product_title)
                tk.Label(root, text="{}: {}".format(product_title, price)).grid(row=11, columnspan=1)
                shaved_price = price[2:len(price) - 3]
                int_price = str(shaved_price).replace(',', '')
                prices.append(int(int_price))
                parts_dict['PARTS'].append(product_title)
                parts_dict['PRICE'].append(int_price)
        elif i == 2:
            if c.get() == '':
                pg['value'] = i * 10
                pg.update()
                continue
            else:
                pg['value'] = i * 10
                pg.update()
                driver.get("https://amazon.in")
                searchb = driver.find_element_by_xpath(
                    '/html/body/div[1]/header/div/div[1]/div[3]/div/form/div[2]/div/input')
                searchb.click()
                search_box = driver.find_element_by_xpath(
                    '/html/body/div[1]/header/div/div[1]/div[3]/div/form/div[3]/div[1]/input')
                search_box.send_keys(c.get())
                search_box.send_keys(Keys.RETURN)
                SEARCH_RESULT_LINK = driver.find_element_by_xpath(
                    "(//div[@class='sg-col-inner']//img[contains(@data-image-latency,'s-product-image')])[1]")
                SEARCH_RESULT_LINK.click()
                driver.switch_to.window(driver.window_handles[0])
                driver.close()
                driver.switch_to.window(driver.window_handles[0])
                price = driver.find_element_by_id("priceblock_ourprice").text
                product_title = driver.find_element_by_id("productTitle").text
                titles.append(product_title)
                tk.Label(root, text="{}: {}".format(product_title, price)).grid(row=12, columnspan=1)
                shaved_price = price[2:len(price) - 3]
                int_price = str(shaved_price).replace(',', '')
                prices.append(int(int_price))
                parts_dict['PARTS'].append(product_title)
                parts_dict['PRICE'].append(int_price)
        elif i==3:
            if d.get() == '':
                pg['value'] = i * 10
                pg.update()
                continue
            else:
                pg['value'] = i * 10
                pg.update()
                driver.get("https://amazon.in")
                searchb = driver.find_element_by_xpath(
                    '/html/body/div[1]/header/div/div[1]/div[3]/div/form/div[2]/div/input')
                searchb.click()

                search_box = driver.find_element_by_xpath(
                    '/html/body/div[1]/header/div/div[1]/div[3]/div/form/div[3]/div[1]/input')
                search_box.send_keys(d.get())
                search_box.send_keys(Keys.RETURN)
                SEARCH_RESULT_LINK = driver.find_element_by_xpath(
                    "(//div[@class='sg-col-inner']//img[contains(@data-image-latency,'s-product-image')])[1]")
                SEARCH_RESULT_LINK.click()
                driver.switch_to.window(driver.window_handles[0])
                driver.close()
                driver.switch_to.window(driver.window_handles[0])
                price = driver.find_element_by_id("priceblock_ourprice").text
                product_title = driver.find_element_by_id("productTitle").text
                titles.append(product_title)
                tk.Label(root, text="{}: {}".format(product_title, price)).grid(row=13, columnspan=1)
                shaved_price = price[2:len(price) - 3]
                int_price = str(shaved_price).replace(',', '')
                prices.append(int(int_price))
                parts_dict['PARTS'].append(product_title)
                parts_dict['PRICE'].append(int_price)
        elif i == 4:
            if e.get() == '':
                pg['value'] = i * 10
                pg.update()
                continue
            else:
                pg['value'] = i * 10
                pg.update()
                driver.get("https://amazon.in")
                searchb = driver.find_element_by_xpath(
                    '/html/body/div[1]/header/div/div[1]/div[3]/div/form/div[2]/div/input')
                searchb.click()
                search_box = driver.find_element_by_xpath(
                    '/html/body/div[1]/header/div/div[1]/div[3]/div/form/div[3]/div[1]/input')
                search_box.send_keys(e.get())
                search_box.send_keys(Keys.RETURN)
                SEARCH_RESULT_LINK = driver.find_element_by_xpath(
                    "(//div[@class='sg-col-inner']//img[contains(@data-image-latency,'s-product-image')])[1]")
                SEARCH_RESULT_LINK.click()
                driver.switch_to.window(driver.window_handles[0])
                driver.close()
                driver.switch_to.window(driver.window_handles[0])
                price = driver.find_element_by_id("priceblock_ourprice").text
                product_title = driver.find_element_by_id("productTitle").text
                titles.append(product_title)
                tk.Label(root, text="{}: {}".format(product_title, price)).grid(row=14, columnspan=1)
                shaved_price = price[2:len(price) - 3]
                int_price = str(shaved_price).replace(',', '')
                prices.append(int(int_price))
                parts_dict['PARTS'].append(product_title)
                parts_dict['PRICE'].append(int_price)
        elif i == 5:
            if f.get() == '':
                pg['value'] = i * 10
                pg.update()
                continue
            else:
                pg['value'] = i * 10
                pg.update()
                driver.get("https://amazon.in")
                searchb = driver.find_element_by_xpath(
                    '/html/body/div[1]/header/div/div[1]/div[3]/div/form/div[2]/div/input')
                searchb.click()
                search_box = driver.find_element_by_xpath(
                    '/html/body/div[1]/header/div/div[1]/div[3]/div/form/div[3]/div[1]/input')
                search_box.send_keys(f.get())
                search_box.send_keys(Keys.RETURN)
                SEARCH_RESULT_LINK = driver.find_element_by_xpath(
                    "(//div[@class='sg-col-inner']//img[contains(@data-image-latency,'s-product-image')])[1]")
                SEARCH_RESULT_LINK.click()
                driver.switch_to.window(driver.window_handles[0])
                driver.close()
                driver.switch_to.window(driver.window_handles[0])
                #price = driver.find_element_by_id("priceblock_ourprice").text
                try:
                    price = driver.find_element_by_id("priceblock_ourprice").text
                except NoSuchElementException:
                    price = driver.find_element_by_id("priceblock_dealprice").text
                product_title = driver.find_element_by_id("productTitle").text
                titles.append(product_title)
                tk.Label(root, text="{}: {}".format(product_title, price)).grid(row=15, columnspan=1)
                shaved_price = price[2:len(price) - 3]
                int_price = str(shaved_price).replace(',', '')
                prices.append(int(int_price))
                parts_dict['PARTS'].append(product_title)
                parts_dict['PRICE'].append(int_price)
        else:
            if g.get() == '':
                pg['value'] = i * 10
                pg.update()
                tk.Label(root, text="Extracted All Prices").grid(row=10, column=1)
                break
            else:
                pg['value'] = i * 10
                pg.update()
                driver.get("https://amazon.in")
                searchb = driver.find_element_by_xpath(
                    '/html/body/div[1]/header/div/div[1]/div[3]/div/form/div[2]/div/input')
                searchb.click()
                search_box = driver.find_element_by_xpath(
                    '/html/body/div[1]/header/div/div[1]/div[3]/div/form/div[3]/div[1]/input')
                search_box.send_keys(g.get())
                search_box.send_keys(Keys.RETURN)
                SEARCH_RESULT_LINK = driver.find_element_by_xpath(
                    "(//div[@class='sg-col-inner']//img[contains(@data-image-latency,'s-product-image')])[1]")
                SEARCH_RESULT_LINK.click()
                driver.switch_to.window(driver.window_handles[0])
                driver.close()
                driver.switch_to.window(driver.window_handles[0])
                price = driver.find_element_by_id("priceblock_ourprice").text
                product_title = driver.find_element_by_id("productTitle").text
                titles.append(product_title)
                tk.Label(root, text="{}: {}".format(product_title, price)).grid(row=16, columnspan=1)
                shaved_price = price[2:len(price) - 3]
                int_price = str(shaved_price).replace(',', '')
                prices.append(int(int_price))
                parts_dict['PARTS'].append(product_title)
                parts_dict['PRICE'].append(int_price)
                driver.quit()
                tk.Label(root, text = "Extracted All Prices").grid(row = 10,column = 1)
    root.update()
    if len(prices) == 0:
        messagebox.showerror("Error","Unsuccessful, fill up atleast one field to calculate properly.")
    else:
        total = sum(prices)
        range_u = total+2000
        range_l = total-2000
        tk.Label(root, text = 'Budget Range: ' + 'Rs{} - Rs{}'.format(range_l,range_u)).grid(row = 17,column = 0)
        messagebox._show("Completed","Budget Calculated Successfully!!")
    root.update()



def close():
    root.destroy()



def clear_all():
    a.delete(0,'end')
    b.delete(0,'end')
    c.delete(0,'end')
    d.delete(0,'end')
    e.delete(0,'end')
    f.delete(0,'end')
    g.delete(0,'end')
    titles.clear()
    prices.clear()

def save_it():
    if len(titles)<1:
        messagebox.showerror("Error","Fill in atleast one field and click on 'show' once to save to text file")
    else:
        data = titles[0]
        filename = filedialog.asksaveasfile(initialdir = 'C:/users/eradi/Documents',title = "Save as",defaultextension = ".txt" ,filetypes = (("notepad","*.txt"),("all files","*.*")))
        for i in range(len(titles)):
            print('{}: Rs{}'.format(titles[i],prices[i]),file = filename)
        print('BUDGET RANGE: Rs{} - Rs{}'.format(sum(prices)-2000,sum(prices)+2000),file = filename)
        #filename.write(data)
        #filename.close()

def save_as_excel():
    if len(titles)<1:
        messagebox.showerror("Error", "Fill in atleast one field and click on 'show' once to save as excel file")
    else:
        file_excel = filedialog.asksaveasfile(initialdir = 'C:/users/eradi/Documents',title = "Save as",defaultextension = ".xlsx" ,filetypes = (("excel file","*.xlsx"),("all files","*.*")))
        df = pd.DataFrame(parts_dict)
       # writer = pd.ExcelWriter('{}.xlsx'.format(file_excel), engine='xlsxwriter')
       # df.to_excel(writer, sheet_name='Sheet1')
       # writer.save()
        df.to_excel('C:\\users\\eradi\\Documents\\{}.xlsx'.format(file_excel),sheet_name='sheet_1')






w = root.title("PC builder")

root.iconbitmap(r'pc.ico')
#inputs

tk.Label(root, text = 'CPU: ').grid(row= 0,column = 0)
a = tk.Entry(root)
a.grid(row = 0,column = 1)
tk.Label(root, text = 'GPU: ').grid(row= 1,column = 0)
b = tk.Entry(root)
b.grid(row = 1,column = 1)
tk.Label(root, text = 'Motherboard: ').grid(row= 2,column = 0)
c = tk.Entry(root)
c.grid(row = 2,column = 1)
tk.Label(root, text = 'RAM: ').grid(row= 3,column = 0)
d = tk.Entry(root)
d.grid(row = 3,column = 1)
tk.Label(root, text = 'Storage: ').grid(row= 4,column = 0)
e = tk.Entry(root)
e.grid(row = 4,column = 1)
tk.Label(root, text = 'Case: ').grid(row= 5,column = 0)
f = tk.Entry(root)
f.grid(row = 5,column = 1)
tk.Label(root, text = 'PSU: ').grid(row= 6,column = 0)
g = tk.Entry(root)
g.grid(row = 6,column = 1)


#buttons
cal = tk.Button(root, text = 'Calculate total cost on Amazon', command = calculate)
cal.grid(row = 7,column = 1)
clear = tk.Button(root, text = 'Clear all fields', command = clear_all)
clear.grid(row = 8,column = 1)
save = tk.Button(root, text = 'Save PC parts and prices as text file', command = save_it)
save.grid(row = 9,column = 1)
root.protocol('WM_DELETE_WINDOW',close)
#excel = tk.Button(root, text = 'Save as Excel File',command = save_as_excel)
#excel.grid(row = 9,column = 2)

#progress bar
#pg = Progressbar(root, length = 200, orient = 'horizontal',maximum = 60,value = 0,mode = 'determinate')
#pg.grid(row = 11,column = 1)

root.mainloop()
