import tkinter as tk
import tkinter.filedialog
from functools import partial
import Menu_Planner_csv
import subprocess
import os
import sys
import math

class Window(object):
    """
    Main window containing application
    """
        
    def __init__(self, master):
        self.mainFrame = tk.Frame(master)
        self.header = tk.Frame(self.mainFrame)
        self.master = master
        self.directory = None

        self.start_btn = tk.Button(self.header, text="Start", command=self.start_pressed, height=2, width=9)
        self.load_btn = tk.Button(self.header, text='Load Files', command=self.load_file_pressed, height=2, width=9)
        self.start_btn.pack(side=tk.LEFT, pady=25, padx=50)
        self.load_btn.pack(side=tk.RIGHT, pady=25, padx=50)

        self.header.pack(side=tk.TOP, fill=tk.BOTH)
        self.mainFrame.pack(fill=tk.BOTH, expand=1)

        self.infotext = tk.Label(self.mainFrame, text="Please select a CSV file for the orders", font="-weight bold")
        self.infotext.pack(side=tk.TOP, anchor=tk.W, fill=tk.BOTH, padx=20, pady=20)
        self.weekLabel = tk.Label(self.mainFrame, text="We are currently on week " + Menu_Planner_csv.getCurrentWeek()
                                  + '\n and we will be using ' + Menu_Planner_csv.getLastWeek() + '.xlsx for history')
        self.weekLabel.pack(side=tk.TOP, anchor=tk.W, fill=tk.BOTH, padx=20, pady=20)
        self.runningLabel = tk.Label(self.mainFrame)
        self.runningLabel.pack(side=tk.TOP)


        #self.canvas=tk.Canvas(self.mainFrame,width=400,height=300, scrollregion=(0,0,800,800))
        #self.canvas.config(width=400,height=300)
        #self.canvas.pack()
        self.bottomFrame = tk.Frame(self.mainFrame)
        #scrollbar = tk.Scrollbar(self.bottomFrame)
        #scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.bottomFrame.pack()
        

        mealsText = "Open Meal Excel File (" + Menu_Planner_csv.getCurrentWeek() + ".xlsx)"
        self.mealsBtn = tk.Button(self.bottomFrame, text=mealsText, command= self.openMealExcel)
        self.mealsBtn.pack(side=tk.TOP)
        ingredientsText = "Open Ingredients Excel File (Ingredient and meal costing1.xlsx)"
        self.ingredientsBtn = tk.Button(self.bottomFrame, text=ingredientsText, command= self.openIngredientExcel)
        self.ingredientsBtn.pack(side=tk.TOP)
        
        self.maxNumCustoms = 10
        self.make_table()
                  
                
        self.table.pack(pady=30)
        

    def make_table(self):
        self.table = tk.Frame(self.bottomFrame)
        self.mealsTable = {}
        for x in range(6):
            for y in range(11): 
                self.mealsTable[(x,y)] = tk.Label(self.table, text=str(x) + "," + str(y), borderwidth=1, relief="solid")
                self.mealsTable[(x,y)].grid(column=x, row=y, sticky="nsew")

                if (x == 3 or x == 1):
                    self.mealsTable[(x,y)].grid(padx=(0,50))

        self.mealsTable[(0,0)].config(text="Classic", font="-weight bold")
        self.mealsTable[(1,0)].config(text="Num", font="-weight bold")
        self.mealsTable[(2,0)].config(text="Asian", font="-weight bold")
        self.mealsTable[(3,0)].config(text="Num", font="-weight bold")
        self.mealsTable[(4,0)].config(text="Customs", font="-weight bold")
        self.mealsTable[(5,0)].config(text="Num", font="-weight bold")
        
    def start_pressed(self):
        if not(self.directory is None):
            self.runningLabel.config(text="Running...")
            Menu_Planner_csv.main(self.directory)
            
            
            self.populate_table()

            self.runningLabel.config(text="Done!")
        else:
            self.infotext.config(fg="red")


    def load_file_pressed(self):
        self.directory = tk.filedialog.askopenfilename(initialdir = "/",title = "Select file",filetypes = (("CSV file","*.csv"), ("All Files","*.*")))
        self.infotext.config(text= 'We are using \'' + self.directory + '\' for this weeks orders.', fg="black", font="-weight normal")

    def openMealExcel(self):
        os.system('open ' + Menu_Planner_csv.getCurrentWeek() + ".xlsx")
        
    def openIngredientExcel(self):
        os.system('open "Ingredient and meal costing1.xlsx"')

    def populate_table(self):
        POrds = Menu_Planner_csv.getPOrds()
        Customs = Menu_Planner_csv.getCustoms()
        print(Menu_Planner_csv.getCustomsLen())

        self.clean_table()

        for i in range(1, 11):
            self.mealsTable[(0,i)].config(text = POrds.at[i-1, 'Classic'])
            self.mealsTable[(1,i)].config(text = str(int(POrds.at[i-1, 'numC'])))
            self.mealsTable[(2,i)].config(text = POrds.at[i-1, 'Asian'])
            self.mealsTable[(3,i)].config(text = str(int(POrds.at[i-1, 'numA'])))

        if (Menu_Planner_csv.getCustomsLen() > 9):
            if (self.maxNumCustoms < Menu_Planner_csv.getCustomsLen()):
                self.maxNumCustoms = Menu_Planner_csv.getCustomsLen()
                
            for j in range(11, Menu_Planner_csv.getCustomsLen()+1):
                self.mealsTable[(4, j)] = tk.Label(self.table, text="," + str(j), borderwidth=1, relief="solid")
                self.mealsTable[(5, j)] = tk.Label(self.table, text="," + str(j), borderwidth=1, relief="solid")
                
                self.mealsTable[(4,j)].grid(column=4, row=j, sticky="nsew")
                self.mealsTable[(5,j)].grid(column=5, row=j, sticky="nsew")

        for i in range(1, Menu_Planner_csv.getCustomsLen()+1):
            self.mealsTable[(4,i)].config(text = str(Customs.at[i-1, 'Customs']))
            self.mealsTable[(5,i)].config(text = str(int(Customs.at[i-1, 'numCust'])))
        
    def clean_table(self):
        for x in range(1,6):
            for y in range(1,11): 
                self.mealsTable[(x,y)].config(text='-')

        for i in range(1, max(self.maxNumCustoms+1, 10)):
            self.mealsTable[(4,i)].config(text = "-")
            self.mealsTable[(5,i)].config(text = "-")
                
class Main(object):
    def __init__(self, master):
        self._master = master
        self._master.title("Meal Picker")
        self._master.geometry("600x260")
        self._master.minsize("750","610")
        app = Window(master)
        

root = tk.Tk()
main = Main(root)
root.mainloop()
