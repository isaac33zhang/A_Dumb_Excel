#!/usr/bin/env python

import sys
sys.path.append('/Users/fanzhaf/Documents/Personal/PyCharm/PycharmProjects/excel/')
from Tkinter import Tk, Frame, Label, Button, Entry, LEFT, RIGHT, END, W, E, StringVar
from lib.pangpang_helper import PangPangHelper


class PangPangGUI(Frame):
    def __init__(self, master=None):
        Frame.__init__(self, master)
        master.title("PangPangSing")
        self.label = Label(master, text="PangPangSing")
        self.label.pack()

        self.department_file = None
        self.salary_file = None
        self.destination = None

        self.row1 = Label(master, text="Department")
        self.row2 = Label(master, text="Salary")
        self.row3 = Label(master, text="Destination")
        self.department_entry = Entry(master)
        self.salary_entry = Entry(master)
        self.destination_entry = Entry(master)

        self.button = Button(master, text="Sing", command=self.sing)

        self.row1.pack()
        self.department_entry.pack()
        self.row2.pack()
        self.salary_entry.pack()
        self.row3.pack()
        self.destination_entry.pack()
        self.button.pack()

        self.close = Button(master, text="Quit", command=master.quit)
        self.close.pack()

    def sing(self):
        self.department_file = self.department_entry.get()
        self.salary_file = self.salary_entry.get()
        self.destination = self.destination_entry.get()
        helper = PangPangHelper(self.department_file, self.salary_file, self.destination)
        helper.run()


root = Tk()
pangpang_gui = PangPangGUI(root)
root.mainloop()
