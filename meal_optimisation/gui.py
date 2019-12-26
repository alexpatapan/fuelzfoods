import tkinter as tk
import tkinter.filedialog
from functools import partial
import Menu_Planner_csv
import subprocess
import os

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

        self.infotext = tk.Label(self.mainFrame, text="Please select a CSV file for the orders")
        self.infotext.pack(side=tk.TOP, anchor=tk.W, fill=tk.BOTH, padx=20, pady=20)
        self.weekLabel = tk.Label(self.mainFrame, text="We are currently on week " + Menu_Planner_csv.getCurrentWeek()
                                  + '\n and we will be using week ' + Menu_Planner_csv.getLastWeek() + ' for history')
        self.weekLabel.pack(side=tk.TOP, anchor=tk.W, fill=tk.BOTH, padx=20, pady=20)
        self.runningLabel = tk.Label(self.mainFrame)
        self.runningLabel.pack(side=tk.TOP)

        self.bottomFrame = tk.Frame(self.mainFrame)
        self.linkLabel = tk.Button(self.bottomFrame, text="Meal Excel File", command= self.openExcel)
        self.linkLabel.pack(side=tk.RIGHT)
        self.bottomFrame.pack()
        
    def start_pressed(self):
        if not(self.directory is None):
            self.runningLabel.config(text="Running...")
            Menu_Planner_csv.main(self.directory)
            self.runningLabel.config(text="Done!")


    def load_file_pressed(self):
        self.directory = tk.filedialog.askopenfilename(initialdir = "/",title = "Select file",filetypes = (("CSV file","*.csv"), ("All Files","*.*")))
        self.infotext.config(text= 'We are using \'' + self.directory + '\' for this weeks orders.')

    def openExcel(self):
        print("a")
        #os.system('start "excel" \"' + Menu_Planner_csv.getCurrentWeek() + '\".xlsx')
        os.system('start "excel" "1952.xlsx"')
        

        
class Main(object):
    def __init__(self, master):
        self._master = master
        self._master.title("Meal Picker")
        self._master.geometry("600x260")
        self._master.minsize("200","300")
        app = Window(master)
        

root = tk.Tk()
main = Main(root)
root.mainloop()
