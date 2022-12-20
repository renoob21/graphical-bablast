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

        self.browse = customtkinter.CTkButton(master=self,text="Browse", command=self.browse_file)
        self.browse.grid(row=1, column=0, padx=20, pady=20, sticky='ew')

        self.msg = customtkinter.CTkButton(master=self,text="Get Text", command=self.get_msg)
        self.msg.grid(row=1, column=1, padx=20, pady=20, sticky='ew')



    def browse_file(self):
        filepath = tkinter.filedialog.askopenfilename(filetypes=(("Excel file", "*.xlsx"), ("All Files","*.*")))
        self.data = pd.read_excel(filepath,dtype=str)

    def get_msg(self):
        msg = self.msg_input.get("0.0",'end')
        print(msg)

if __name__ == '__main__':
    app = BaBlast()
    app.mainloop()