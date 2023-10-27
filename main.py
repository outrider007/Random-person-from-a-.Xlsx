version = 'beta v0.2'




from tkinter import *
from tkinter import ttk
import tkinter as tk
import random
import xlrd
import webbrowser



global people
people = ''
global chosen_list
chosen_list =''
global above_text
above_text = 'Press the button'


xlsx_file = xlrd.open_workbook("main.xlsx")

sheet1 = xlsx_file.sheet_by_index(0)


list1 = ['bob', 'james']
list2 = ['']
list3 = ['']
list4 = ['']
list5 = ['']
list6 = ['']
list7 = ['']
list8 = ['']
list9 = ['']
list10 = ['']

list1_name = str(sheet1.cell_value(2, 4))
list2_name = str(xlsx_file.sheet_by_index(1).cell_value(2, 4))        
list3_name = str(xlsx_file.sheet_by_index(2).cell_value(2, 4))
list4_name = str(xlsx_file.sheet_by_index(3).cell_value(2, 4))
list5_name = str(xlsx_file.sheet_by_index(4).cell_value(2, 4))
list6_name = str(xlsx_file.sheet_by_index(5).cell_value(2, 4))
list7_name = str(xlsx_file.sheet_by_index(6).cell_value(2, 4))
list8_name = str(xlsx_file.sheet_by_index(7).cell_value(2, 4))
list9_name = str(xlsx_file.sheet_by_index(8).cell_value(2, 4))
list10_name = str(xlsx_file.sheet_by_index(9).cell_value(2, 4))

resoliton_x = int(sheet1.cell_value(2, 17))
resoliton_y = int(sheet1.cell_value(3, 17))

background = sheet1.cell_value(5, 17)
foreground = sheet1.cell_value(6, 17)

Size_output = sheet1.cell_value(7, 17)
Size_list = sheet1.cell_value(8, 17)
Size_labels = sheet1.cell_value(9, 17)
Size_Error =sheet1.cell_value(10, 17)

list_names = [list1_name, list2_name, list3_name, list4_name, list5_name, list6_name, list7_name, list8_name, list9_name]
x = 0


def read_sheet1():
        global people
        people = []
        person = 'a'
        n = 0
        list = xlsx_file.sheet_by_index(0)
        for row in range(list.ncols):
            non_emptycells=[i for i,x in enumerate(list.col(0)) if x.ctype != 0]                  
            
        for cell in non_emptycells:
            person = list.cell_value(cell, 0)  
            people.append(person)                
            
def read_sheet2():
        global people
        people = []
        person = 'a'
        n = 0
        list = xlsx_file.sheet_by_index(1)
        for row in range(list.ncols):
            non_emptycells=[i for i,x in enumerate(list.col(0)) if x.ctype != 0]                  
            
        for cell in non_emptycells:
            person = list.cell_value(cell, 0)  
            people.append(person)                
            
def read_sheet3():
        global people
        people = []
        person = 'a'
        list = xlsx_file.sheet_by_index(2)
        for row in range(list.ncols):
            non_emptycells=[i for i,x in enumerate(list.col(0)) if x.ctype != 0]                  
            
        for cell in non_emptycells:
            person = list.cell_value(cell, 0)  
            people.append(person)      

def read_sheet4():
        global people
        people = []
        person = 'a'
        list = xlsx_file.sheet_by_index(3)
        for row in range(list.ncols):
            non_emptycells=[i for i,x in enumerate(list.col(0)) if x.ctype != 0]                  
            
        for cell in non_emptycells:
            person = list.cell_value(cell, 0)  
            people.append(person)      
    
def read_sheet5():
        global people
        people = []
        person = 'a'
        list = xlsx_file.sheet_by_index(4)
        for row in range(list.ncols):
            non_emptycells=[i for i,x in enumerate(list.col(0)) if x.ctype != 0]                  
            
        for cell in non_emptycells:
            person = list.cell_value(cell, 0)  
            people.append(person)      

def read_sheet6():
        global people
        people = []
        person = 'a'
        list = xlsx_file.sheet_by_index(5)
        for row in range(list.ncols):
            non_emptycells=[i for i,x in enumerate(list.col(0)) if x.ctype != 0]                  
            
        for cell in non_emptycells:
            person = list.cell_value(cell, 0)  
            people.append(person)    

def read_sheet7():
        global people
        people = []
        person = 'a'
        list = xlsx_file.sheet_by_index(6)
        for row in range(list.ncols):
            non_emptycells=[i for i,x in enumerate(list.col(0)) if x.ctype != 0]                  
            
        for cell in non_emptycells:
            person = list.cell_value(cell, 0)  
            people.append(person)       

def read_sheet8():
        person = 'a'
        list = xlsx_file.sheet_by_index(7)
        for row in range(list.ncols):
            non_emptycells=[i for i,x in enumerate(list.col(0)) if x.ctype != 0]                  
            
        for cell in non_emptycells:
            person = list.cell_value(cell, 0)  
            people.append(person)    
    
def read_sheet9():
        global people
        people = []
        person = 'a'
        list = xlsx_file.sheet_by_index(8)
        for row in range(list.ncols):
            non_emptycells=[i for i,x in enumerate(list.col(0)) if x.ctype != 0]                  
            
        for cell in non_emptycells:
            person = list.cell_value(cell, 0)  
            people.append(person)     

def read_sheet10():
        global people
        people = []
        person = 'a'
        list = xlsx_file.sheet_by_index(9)
        for row in range(list.ncols):
            non_emptycells=[i for i,x in enumerate(list.col(0)) if x.ctype != 0]                  
            
        for cell in non_emptycells:
            person = list.cell_value(cell, 0)  
            people.append(person)    
    
def randomize():
        global above_text
        final_list = ''
        if chosen_list == list1_name:
            read_sheet1()
            list1 = people     
            final_list = list1        
        if chosen_list == list2_name:
            read_sheet2()
            list2 = people     
            final_list = list2          
        if chosen_list == list3_name:
            read_sheet3()
            list3 = people     
            final_list = list3
        if chosen_list == list4_name:
            read_sheet4()
            list4 = people     
            final_list = list4
        if chosen_list == list5_name:
            read_sheet5()
            list5 = people     
            final_list = list5                  
        if chosen_list == list6_name:
            read_sheet6()
            list6 = people     
            final_list = list6                  
        if chosen_list == list7_name:
            read_sheet7()
            list7 = people     
            final_list = list7    
        if chosen_list == list8_name:
            read_sheet8()
            list8 = people     
            final_list = list8    
        if chosen_list == list9_name:
            read_sheet9()
            list9 = people     
            final_list = list9    
        if chosen_list == list10_name:
            read_sheet10()
            list10= people     
            final_list = list10              
        if chosen_list == '':
            above_text='Please select valid sheet'     
                 
        else:
           output = random.choice(final_list)   
           above_text=output     

def Open_GitHub():
    webbrowser.open_new('https://github.com/outrider007/Random-person-from-a.-Xlsx')                  


class App():

    def __init__(self):
        def randomize_init():
           global chosen_list
           chosen_list = self.set_sheet.get()
           randomize()
           if above_text == '"Please select valid sheet"':
            self.text_output.configure(text=above_text, font=('Times', 30))   
           else:
            self.text_output.configure(text=above_text, font=('Brass Mono', 40))         
        
        
        self.root = tk.Tk()
        self.root.geometry(f'{resoliton_x}x{resoliton_y}')
        self.root.title(f'Randomizer   {version}, res = {resoliton_x}x{resoliton_y}             , language = English')
        self.root.iconbitmap('icon.ico')
        self.mainframe = tk.Frame(self.root, background=background)
        self.mainframe.pack(fill='both', expand=True)


        #               The output is going to be here

        self.text_welcome = ttk.Label(self.mainframe, text='The lucky person is....... ', background=background, font=('Brass Mono', 20), foreground=foreground)
        self.text_welcome.grid(row=1, column=0, sticky='NWES', pady=10)

        #                  Output
        self.text_output = ttk.Label(self.mainframe, text=above_text, background=background, font=('Brass Mono', 30),  foreground=foreground)
        self.text_output.grid(row=2, column=0, sticky='NWES', pady=10)

        #               Randomize!
        set_random_button = ttk.Button(self.mainframe,command=randomize_init, text= "Randomize!",)
        set_random_button.grid(row=3, column=0, ipady=20, ipadx=20,)
        
        #              (text)choose a sheet
        self.text_welcome = ttk.Label(self.mainframe, text='Change sheet -', background=background, font=('Brass Mono', 20), foreground=foreground)
        self.text_welcome.grid(row=4, column=0, sticky='NWES', pady=10)
        
        #              (list) choose a sheet
        self.set_sheet = ttk.Combobox(self.mainframe, values=list_names, font=('Brass Mono', 15))
        self.set_sheet.grid(row=5, column=0, sticky = 'NWES', pady=20)
        #               https://github.com/outrider007/Random-person-from-a.-Xlsx
        self.git_link = ttk.Button(self.mainframe, text = 'Source code on GitHub', command = Open_GitHub)
        self.git_link.grid(row=6, column=0, sticky='NWES', pady = 50, padx = 60)
                            
        self.root.mainloop()
        return







        
    

if __name__ == '__main__':
    App()