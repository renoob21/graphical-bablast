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

class BaBlast(customtkinter.CTk):
    def __init__(self):
        super().__init__()

        self.title("BaBlast - Aplikasi Blasting WhatsApp")
        self.minsize(400, 300)

        self.browse = customtkinter.CTkButton(master=self,text="Browse", command=self.browse_file)
        self.browse.pack(padx=20, pady=20)

    def browse_file(self):
        filepath = tkinter.filedialog.askopenfilename(filetypes=(("Excel file", "*.xlsx"), ("File CSV berbatas titik koma (;)","*.csv")))
        self.data = pd.read_excel(filepath,dtype=str)
        print(self.data.head())

if __name__ == '__main__':
    app = BaBlast()
    app.mainloop()