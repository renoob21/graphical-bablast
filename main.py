import pandas as pd
import selenium
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
import time
from pynput.keyboard import Key, Controller
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
import customtkinter
import tkinter

customtkinter.set_default_color_theme("dark-blue")

class BaBlast(customtkinter.CTk):
    def __init__(self):
        super().__init__()

        self.geometry("600x450")
        self.title("BaBlast - Aplikasi Blasting WhatsApp")
        self.minsize(400, 300)

        self.rowconfigure(0,weight=1)
        self.columnconfigure((0,1,2), weight=1)

        self.msg_input = customtkinter.CTkTextbox(self)
        self.msg_input.grid(row=0, column=0, columnspan=3, padx=20, pady=(20,0), sticky='nsew')

        self.browse = customtkinter.CTkButton(master=self,text="Buka daftar WP", command=self.browse_file)
        self.browse.grid(row=1, column=0, padx=20, pady=20, sticky='ew')

        self.login_button = customtkinter.CTkButton(master=self, text='Login', command=self.login)
        self.login_button.grid(row=1,column=1, padx=20, pady=20, sticky='ew')

        self.blast_button = customtkinter.CTkButton(master=self, text='Blast!', command=self.blast)
        self.blast_button.grid(row=1,column=2, padx=20, pady=20, sticky='ew')
        self.blast_button.configure(state='disabled')

        # self.msg = customtkinter.CTkButton(master=self,text="Get Text", command=self.get_msg)
        # self.msg.grid(row=1, column=1, padx=20, pady=20, sticky='ew')



    def browse_file(self):
        filepath = tkinter.filedialog.askopenfilename(filetypes=(("Excel file", "*.xlsx"), ("All Files","*.*")))
        self.data = pd.read_excel(filepath,dtype=str)
        self.data['status'] = '-'

    def login(self):
        self.driver = webdriver.Chrome(ChromeDriverManager().install(),options=Options())
        self.driver.get('https://web.whatsapp.com/')
        self.blast_button.configure(state='normal')
        #self.blast_button.grid(row=1,column=2, padx=20, pady=20, sticky='ew')
        

    def blast(self):
        for i in range(len(self.data)):
            current_data = self.data.iloc[i]
            text = self.msg_input.get('0.0','end')
            text = text.format(nama_wp=current_data['nama_wp'],npwp=current_data['npwp'], nama_ar=current_data['nama_ar'],nomor_ar=current_data['nomor_ar'])
            url = 'https://web.whatsapp.com/send?phone={}&text&type=phone_number&app_absent=0'.format(current_data.nomor_wp)
            self.driver.get(url)
            try :
                WebDriverWait(self.driver,10).until(EC.element_to_be_clickable((By.XPATH,'//div[@title="Type a message"]')))
                field = self.driver.find_element('xpath','//div[@title="Type a message"]')
                for msg in text.split('\n'):
                    field.send_keys(msg)
                    field.send_keys(Keys.SHIFT, Keys.ENTER)
                send = self.driver.find_element('xpath','//button[@aria-label="Send"]')
                send.click()
                self.data['status'].iloc[i] = 'Sukses!'
                time.sleep(2)
            except:
                self.data['status'].iloc[i] = 'Gagal!'
        self.data.to_excel('hasil_blast.xlsx', index=False)

    def get_msg(self):
        msg = self.msg_input.get("0.0",'end')
        print(msg)

if __name__ == '__main__':
    app = BaBlast()
    app.mainloop()