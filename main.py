from tkinter import ttk
import tkinter as tk
import random
import pandas  as pd

global sheet1
global sheet2
global x
global list1

excel = pd.read_excel('main.xlsx')

sheet1 = pd.read_excel('main.xlsx', sheet_name="1")
sheet2 = pd.read_excel('main.xlsx', sheet_name="2")

list1 = []
list2 = []
list3 = []
list4 = []
list5 = []
list6 = []
list7 = []
list8 = []
list9 = []
list10= []

list1_name = "list1"
list2_name = ""
list3_name = "list3"
list4_name = "list4"
list5_name = "list5"
list6_name = "list6"
list7_name = "list7" 
list8_name = "list8" 
list9_name = "list9"
list10_name =  "list10"


list_names = ["list1", "list2", "list3", "list4", "list5", "list6", "list7", "list8", "list9", "list10" ]
x = 0


class App():
    def __init__(self):
        self.root = tk.Tk()
        self.root.geometry('1920x1080')
        self.root.title('Randomizer                                               language = English, ver betav0.1, res = 1920x1080')
        self.mainframe = tk.Frame(self.root, background='grey')
        self.mainframe.pack(fill='both', expand=True)

        #               The output is going to be here

        self.text_welcome = ttk.Label(self.mainframe, text='output = ', background='grey', font=('Brass Mono', 20), foreground="white")
        self.text_welcome.grid(row=1, column=0, sticky='NWES', pady=10)

        #                  Output
        self.text_output = ttk.Label(self.mainframe, text='Press the button', background='grey', font=('Brass Mono', 50),  foreground="white")
        self.text_output.grid(row=2, column=0, sticky='NWES', pady=10)

        #               Randomize!
        set_random_button = ttk.Button(self.mainframe, command=self.randomize, )
        set_random_button.grid(row=3, column=0, ipady=20, ipadx=20,)

        self.set_color_field = ttk.Combobox(self.mainframe, values=list_names)
        self.set_color_field.grid(row=2, column=0, sticky = 'NWES', pady=10)


        self.root.mainloop()
        return

    def randomize(self):
        




        output = random.choice(choose_list)
        self.text_output.config(text=output)

    def read_sheet1(self):
  

       for i in range(0, 4):
         text = sheet1.cell_value(i, 0)
         if text != "":
          list1.insert(i, text)
          print(list1)

if __name__ == '__main__':
    App()