
import win32com.client 
import win32com
import tkinter as tk
from tkinter import *
#----------------------------------------------------------------
# Test COM interact with Py to manage an ASCOM driver Telescope
# ASCOMPy program is a simple design only for test purposes
# Tested with EQSIM ASCOM Driver 
# -----------------------------------------------------------------
def park_scope():
    print(telescope.AtPark)
     
    if not telescope.AtPark:
      telescope.Park() 
      textpark.set("Telescope Parked")
    if telescope.AtPark:
      textpark.set("Telescope Parked")
      print ("Le télescope est déjà en park")
         
def unpark_scope():
     print(telescope.AtPark)
     if telescope.AtPark:
      telescope.Unpark() 
      textpark.set("Telescope UnParked")
     else:
      print ("Le télescope n'est pas parké")

def connect_driver():
    telescope.connected=True
    telescope.tracking=False
    if telescope.connected:
        Label(fen1, text="State : "+ str(telescope.connected), bg=color, fg=titlecolor).place(x=150,y=80)
    
def choose_telescope():
    global telescope, driver
    global fen1,b,test,x,textpark 
    x = win32com.client.Dispatch("ASCOM.Utilities.Chooser")
    #driver=(x.Choose('EQMOD_SIM.Telescope'))
    x.DeviceType = "Telescope"
    driver=(x.Choose(0))
    telescope = win32com.client.Dispatch(driver)
       
    Label(fen1, text=driver, bg=color, fg=titlecolor).place(x=150,y=50)
    textpark=StringVar()
    b1=tk.Button(fen1,text='Connect', command=connect_driver).place(x=30, y=80)
    b2=tk.Button(fen1,text='Park', command=park_scope).place(x=30, y=110)
    b3=tk.Button(fen1,text='Unpark', command=unpark_scope).place(x=65, y=110)
    Label(fen1, textvar=textpark, bg=color, fg=titlecolor).place(x=150,y=110)
    
#----------------------------------------------------------------
titlecolor="#CB9731"
color="#414242"
fen1 = Tk()
fen1.geometry("500x300")  
fen1.title("ASCOMPy Telescope")
fen1.config (bg=color)
menubar = Menu(fen1)
filemenu = Menu(fen1, tearoff=0)
filemenu.add_command(label="Quit", command=fen1.destroy)
menubar.add_cascade(label="File", menu=filemenu)
fen1.config(menu=menubar)
b=tk.Button(fen1,text='Choose' ,command=choose_telescope).place(x=30, y=50)
fen1.mainloop()
 
