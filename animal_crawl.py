import time
import tkinter as tk
import threading
import queue
from selenium.webdriver.chrome.options import Options
from tkinter import *
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import re
import pandas as pd
import os

try:

    def get_animal_data():
        update_queue.put('查詢中')
        chrome_options = Options()
        chrome_options.add_argument('--headless')
        url='https://www.pet.gov.tw/Web/O302.aspx'
        driver=webdriver.Chrome(options=chrome_options)
        driver.get(url)
        time.sleep(3)

        page_height = driver.execute_script("return document.body.scrollHeight;")
        scroll_position = page_height // 4
        driver.execute_script(f"window.scrollTo(0, {scroll_position});") 
        time.sleep(2) 

        start_y=driver.find_element(By.XPATH,"//input[@name='txtSDATE_year']")

        for _ in range(4):      
            start_y.send_keys(Keys.BACKSPACE)
        start_y.send_keys(f'{start_year.get()}')
        time.sleep(1)

        start_m=driver.find_element(By.XPATH,"//input[@name='txtSDATE_month']")
        for _ in range(2):      
            start_m.send_keys(Keys.BACKSPACE)
        start_m.send_keys(f'{start_month.get()}')
        time.sleep(1)

        start_d=driver.find_element(By.XPATH,"//input[@name='txtSDATE_days']")
        for _ in range(2):      
            start_d.send_keys(Keys.BACKSPACE)
        start_d.send_keys(f'{start_day.get()}')
        time.sleep(1)

        end_y=driver.find_element(By.XPATH,"//input[@name='txtEDATE_year']")
        for _ in range(4):      
            end_y.send_keys(Keys.BACKSPACE)
        end_y.send_keys(f'{end_year.get()}')
        time.sleep(1)

        end_m=driver.find_element(By.XPATH,"//input[@name='txtEDATE_month']")
        for _ in range(2):      
            end_m.send_keys(Keys.BACKSPACE)
        end_m.send_keys(f'{end_month.get()}')
        time.sleep(1)

        end_d=driver.find_element(By.XPATH,"//input[@name='txtEDATE_days']")
        for _ in range(2):      
            end_d.send_keys(Keys.BACKSPACE)
        end_d.send_keys(f'{end_day.get()}')
        time.sleep(1)

        animal_dog=driver.find_element(By.XPATH,f"//label[@for='animal_{val.get()}']")
        animal_dog.click()
        time.sleep(1)

        search=driver.find_element(By.XPATH,"//a[@class='btn btn-main']")
        search.click()
        time.sleep(15)

        height_size=3
        scroll_position_2 = page_height // height_size
        driver.execute_script(f"window.scrollTo(0, {scroll_position_2});")
        city_tc=driver.find_element(By.XPATH,f"//a[@data-item='{val_city.get()}']")    
        city_tc.click()
        time.sleep(15)

        update_queue.put('下載中')

        tables=driver.find_elements(By.XPATH,"//td[@data-th='縣市']")
        data=[]
        for table in tables:
            data.append(table.text)

        numbers=[]
        areas=[]

        for item in data:

            match = re.match(r"(\d+)(\D+)",item)
            if match:
                numbers.append(match.group(1))
                areas.append(match.group(2))
        numbers.append('')
        areas.append('合計')

        tables_A=driver.find_elements(By.XPATH,"//td[@data-th='登記數(A)']")
        data_A=[]
        for table_A in tables_A:
            data_A.append(table_A.text)

        tables_B=driver.find_elements(By.XPATH,"//td[@data-th='除戶數(B)']")
        data_B=[]
        for table_B in tables_B:
            data_B.append(table_B.text)

        tables_C=driver.find_elements(By.XPATH,"//td[@data-th='轉讓數(C)']")
        data_C=[]
        for table_C in tables_C:
            data_C.append(table_C.text)

        tables_D=driver.find_elements(By.XPATH,"//td[@data-th='變更數(D)']")
        data_D=[]
        for table_D in tables_D:
            data_D.append(table_D.text)

        df = pd.DataFrame({
            '郵遞區號': numbers,
            '縣市': val_city.get(),
            '地區': areas,
            '物種': val.get(),
            '登記數':data_A,
            '除戶數':data_B,
            '轉讓數':data_C,
            '變更數':data_D
        })

        print(df)

        mk1 = download_file.get() 
        download_folder = os.path.expanduser(f'~/{mk1}') 
        file_name = f'{val.get()}_{end_year.get()}{end_month.get()}{end_day.get()}.csv'
        if not os.path.exists(download_folder):
            download_folder =f'C:/{mk1}' 
            os.makedirs(download_folder)    
            file_path = os.path.join(download_folder, file_name)
        
        else:
            file_path = os.path.join(download_folder, file_name)

        df.to_csv(file_path,index=False,encoding='utf-8-sig')
        driver.quit()
        update_queue.put('下載完畢')
        os.startfile(download_folder)

    def update_label():
        try:
            while True:
                message = update_queue.get_nowait()
                msg.config(text=message)
        except queue.Empty:
            pass
        root.after(100, update_label) 


    def start_scraping():
        scraping_thread = threading.Thread(target=get_animal_data)
        scraping_thread.start()

    root=Tk()
    root.title('數據自動抓取程式')
    label_frame = tk.LabelFrame(root, text='Author：Avice Su', padx=10, pady=10,bg='#26453D',fg='#BDC0BA',font=("Microsoft JhengHei",10,"bold"))
    label_frame.pack(padx=20, pady=(20,10))
    root.configure(bg='#26453D')
    root.resizable(False, False)
    label1=Label(label_frame,text='開始時間：',bg='#26453D',fg='#CAAD5F',font=("Microsoft JhengHei",12,"bold")) 
    label1.grid(row=0,column=0)
    start_year=Entry(label_frame,width=6) 
    start_year.grid(row=0,column=1,sticky='w')
    label_start_year=Label(label_frame,text='年',bg='#26453D',fg='#CAAD5F',font=("Microsoft JhengHei",12,"bold")) 
    label_start_year.grid(row=0,column=2)
    start_month=Entry(label_frame,width=6) 
    start_month.grid(row=0,column=3,sticky='w')
    label_start_month=Label(label_frame,text='月',bg='#26453D',fg='#CAAD5F',font=("Microsoft JhengHei",12,"bold")) 
    label_start_month.grid(row=0,column=4)
    start_day=Entry(label_frame,width=6) 
    start_day.grid(row=0,column=5,sticky='w')
    label_start_day=Label(label_frame,text='日',bg='#26453D',fg='#CAAD5F',font=("Microsoft JhengHei",12,"bold")) 
    label_start_day.grid(row=0,column=6)

    label2=Label(label_frame,text='結束時間：',bg='#26453D',fg='#CAAD5F',font=("Microsoft JhengHei",12,"bold")) 
    label2.grid(row=1,column=0)
    end_year=Entry(label_frame,width=6) 
    end_year.grid(row=1,column=1,sticky='w')
    label_end_year=Label(label_frame,text='年',bg='#26453D',fg='#CAAD5F',font=("Microsoft JhengHei",12,"bold")) 
    label_end_year.grid(row=1,column=2)
    end_month=Entry(label_frame,width=6) 
    end_month.grid(row=1,column=3,sticky='w')
    label_end_month=Label(label_frame,text='月',bg='#26453D',fg='#CAAD5F',font=("Microsoft JhengHei",12,"bold")) 
    label_end_month.grid(row=1,column=4)
    end_day=Entry(label_frame,width=6) 
    end_day.grid(row=1,column=5,sticky='w')
    label_end_day=Label(label_frame,text='日',bg='#26453D',fg='#CAAD5F',font=("Microsoft JhengHei",12,"bold")) 
    label_end_day.grid(row=1,column=6)

    val = tk.StringVar()
    label3=Label(label_frame,text='物種選擇：',bg='#26453D',fg='#CAAD5F',font=("Microsoft JhengHei",12,"bold")) 
    label3.grid(row=2,column=0)
    radio_dog=tk.Radiobutton(label_frame, text='狗',variable=val, value='dog',bg='#26453D',fg='#CAAD5F',font=("Microsoft JhengHei",12,"bold"))
    radio_dog.grid(row=2,column=1)
    radio_dog.select()
    radio_cat=tk.Radiobutton(label_frame, text='貓',variable=val, value='cat',bg='#26453D',fg='#CAAD5F',font=("Microsoft JhengHei",12,"bold"))
    radio_cat.grid(row=2,column=3)

    optionList=['新北市','臺北市','臺中市','臺南市','高雄市','桃園市','宜蘭縣','新竹縣','苗栗縣','彰化縣','南投縣','雲林縣','嘉義縣','屏東縣','臺東縣','花蓮縣','澎湖縣','基隆市','新竹市','嘉義市','金門縣','連江縣']
    val_city= tk.StringVar()
    label_city=Label(label_frame,text='城市選擇：',bg='#26453D',fg='#CAAD5F',font=("Microsoft JhengHei",12,"bold")) 
    label_city.grid(row=3,column=0)
    city_menu=tk.OptionMenu(label_frame,val_city,*optionList)
    city_menu.config(width=5,fg='black', bg='white',relief='flat', font=("Microsoft JhengHei",10,"bold"))
    city_menu.grid(row=3,column=1,columnspan=3)

    btn=tk.Button(label_frame,text='查詢', padx=9, pady=10,relief='flat',bg='#CAAD5F',fg='#26453D',font=("Microsoft JhengHei",12,"bold"),command=start_scraping)
    btn.grid(row=2,column=5,rowspan=2,columnspan=2, sticky='ws')

    label_frame_file= tk.LabelFrame(root, text='Download File', padx=10, pady=10,bg='#26453D',fg='#BDC0BA',font=("Microsoft JhengHei",10,"bold"))
    label_frame_file.pack(padx=20, pady=10)
    download_file=Entry(label_frame_file,relief='flat',justify='center',width=29,bg='#26453D',fg='#CAAD5F',font=("Microsoft JhengHei",12,"bold")) 
    download_file.grid(row=4,column=0,sticky='w',columnspan=6)
    download_file.insert(0,'Downloads')

    label_frame_msg= tk.LabelFrame(root, text='Message', padx=10, pady=10,bg='#26453D',fg='#BDC0BA',font=("Microsoft JhengHei",10,"bold"))
    label_frame_msg.pack(padx=20, pady=(10,20))
    msg=Label(label_frame_msg,text='請輸入抓取條件',width=29,bg='#26453D',fg='#CAAD5F',font=("Microsoft JhengHei",12,"bold")) 
    msg.grid(row=5,column=0,sticky='ws',columnspan=6)

    update_queue = queue.Queue()
    update_label()
    root.mainloop()
    
except:
    update_queue.put('錯誤，重新查詢')
    pass

