# Import Required Library
import tkinter
import tkinter as tk
from tkinter import *
from tkcalendar import Calendar
from tkcalendar import DateEntry
import ExcelOrganizer
from datetime import timedelta, date
import babel.numbers
# Create Object
root = Tk()

# Set geometry
root.geometry("350x250")

# Setting icon of master window

root.title('Αναζήτηση Δρομολογίων')

text2 = Label(root, text="\n\n")
text3 = Label(root, text="\n\n\n\n")
# Add Calendar
#cal = Calendar(root, selectmode='day')
cal1=DateEntry(root,selectmode='day',date_pattern='dd-MM-yyyy')
cal2=DateEntry(root,selectmode='day',date_pattern='dd-MM-yyyy')
range1 = Label(root, text="Από: ")
range2 = Label(root, text="Έως:")


range1.grid(row=1,column=1)
cal1.grid(row=1,column=3)
text2.grid(row=1)
range2.grid(row=5,column=1)
cal2.grid(row=5,column=3)
text3.grid(row=5)


tk.Label(root, text="Μερική αναζήτηση:").grid(row=10,column=1)
input_text = tk.Entry(root)
input_text.grid(row=10, column=3)
text3.grid(row=10)



def grad_date():
    date_1 = cal1.get_date()
    date_2 = cal2.get_date()
    #print(date_1,date_2)
    dates = []

    for dt in daterange(date_1, date_2):
        dates.append(str(dt.strftime("%d-%m-%Y")))
        #print(dt.strftime("%d-%m-%Y"))

    ExcelOrganizer.ExportFile(dates, input_text.get())
    #root.quit()


def daterange(date1, date2):
    for n in range(int ((date2 - date1).days)+1):
        yield date1 + timedelta(n)

# Add Button and Label
btn= Button(root, text="Eξαγωγή Δρομολογίων", command=grad_date)
btn.grid(row=50,column=3)



# Execute Tkinter
root.mainloop()