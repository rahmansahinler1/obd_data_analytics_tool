import tkinter as tk
from tkinter import *
from PIL import ImageTk, Image
import Functions_GUI.buttons as bt

"""
This script is responsible to create GUI.
"""

global panelA

root = Tk()
root.title("OBD - DiagRA Shot Evaluation Tool")
canvas = tk.Canvas(root, width=936, height=785, bg="#D3D3D3")

# Min and max sizes are declared so user can't change it.
root.minsize(936, 785)
root.maxsize(936, 785)
canvas.pack()

##### LIST OF LABELS AND SCROLLS ####
label1 = Label(root, text="Select DiagRA Shots ")
label1.config(font=("Arial", 8), fg="white", bg="#333333")
label1.place(x=20, y=20, width=418, height=13)

list1 = Listbox(root, font=("Arial", 10), fg="black", bg="white", selectmode=EXTENDED)
list1.place(x=20, y=40, width=400, height=500)
scroll1 = Scrollbar(root)
scroll1.place(x=422, y=40, height=500, width=15)
list1.config(yscrollcommand=scroll1.set)
scroll1.config(command=list1.yview, orient="vertical")

list2 = Listbox(root, font=("Arial", 10), fg="black", bg="white")
list2.place(x=20, y=625, width=879, height=140)
scroll2 = Scrollbar(root)
scroll2.place(x=901, y=625, height=140, width=15)
list2.config(yscrollcommand=scroll2.set)
scroll2.config(command=list2.yview, orient="vertical")

label2 = Label(root, text="Amount of DiagRA Shots: 0")
label2.config(font=("Arial", 8), fg="white", bg="#333333")
label2.place(x=20, y=547, width=418, height=13)

label3 = Label(root, text="Evaluation Process Status")
label3.config(font=("Arial", 8), fg="white", bg="#333333")
label3.place(x=20, y=608, width=896, height=12)

label4 = Label(root, text="Select Excel File")
label4.config(font=("Arial", 8), fg="white", bg="#333333")
label4.place(x=458, y=20, width=458, height=13)

label5 = Label(root, text="Select the Save Path for Results")
label5.config(font=("Arial", 8), fg="white", bg="#333333")
label5.place(x=458, y=104, width=458, height=13)

label6 = Label(root, text="Basic Settings")
label6.config(font=("Arial", 8), fg="white", bg="#333333")
label6.place(x=458, y=188, width=458, height=13)

label7 = Label(root, text="Evaluation Graph / Table")
label7.config(font=("Arial", 8), fg="white", bg="#333333")
label7.place(x=458, y=289, width=458, height=13)

label8 = Label(root, text="Additional P_Code Update")
label8.config(font=("Arial", 8), fg="white", bg="#333333")
label8.place(x=458, y=499, width=458, height=13)

label9 = Label(root, text="Start / Reset / Stop ")
label9.config(font=("Arial", 8), fg="white", bg="#333333")
label9.place(x=458, y=547, width=458, height=13)

##### LIST OF TEXTBOXES ####
text1 = Text(root)
text1.tag_configure('center', justify='center')
text1.insert(1.0, "Load Configuration Excel (.xlsx)", 'center')
text1.config(font=("Arial", 8), fg="black", bg="white", state="disabled")
text1.place(x=458, y=40, width=458, height=21)

text2 = Text(root)
text2.tag_configure('center', justify='center')
text2.insert(1.0, "Default Path: Same Path as Last DiagRA Shot Selected", 'center')
text2.config(font=("Arial", 8), fg="black", bg="white", state="disabled")
text2.place(x=458, y=122, width=458, height=21)

text3 = Text(root)
text3.tag_configure('center', justify='center')
text3.insert(1.0, "Vehicle Name: ", 'center')
text3.config(font=("Arial", 8), fg="black", bg="white", state="disabled")
text3.place(x=458, y=206, width=226, height=16)

text4 = Text(root)
text4.tag_configure('center', justify='center')
text4.insert(1.0, "Control Unit: ", 'center')
text4.config(font=("Arial", 8), fg="black", bg="white", state="disabled")
text4.place(x=458, y=225, width=226, height=16)

text5 = Text(root)
text5.tag_configure('center', justify='center')
text5.insert(1.0, "Create Powerpoint: ", 'center')
text5.config(font=("Arial", 8), fg="black", bg="white", state="disabled")
text5.place(x=458, y=244, width=226, height=16)

text6 = Text(root)
text6.tag_configure('center', justify='center')
text6.insert(1.0, "X-Axis Display Type: ", 'center')
text6.config(font=("Arial", 8), fg="black", bg="white", state="disabled")
text6.place(x=458, y=264, width=226, height=16)

text7 = Text(root)
text7.tag_configure('center', justify='center')
text7.insert(1.0, "Select All:  ", 'center')
text7.config(font=("Arial", 8), fg="black", bg="white", state="disabled")
text7.place(x=458, y=312, width=301, height=18)

text8 = Text(root)
text8.tag_configure('center', justify='center')
text8.insert(1.0, "Date / Kilometer - Graph:  ", 'center')
text8.config(font=("Arial", 8), fg="black", bg="white", state="disabled")
text8.place(x=458, y=335, width=301, height=18)

text9 = Text(root)
text9.tag_configure('center', justify='center')
text9.insert(1.0, "Used Files In Evaluation - Graph / Table: ", 'center')
text9.config(font=("Arial", 8), fg="black", bg="white", state="disabled")
text9.place(x=458, y=357, width=301, height=18)

text10 = Text(root)
text10.tag_configure('center', justify='center')
text10.insert(1.0, "PCode Defect - Graph / Table: ", 'center')
text10.config(font=("Arial", 8), fg="black", bg="white", state="disabled")
text10.place(x=458, y=379, width=301, height=18)

text11 = Text(root)
text11.tag_configure('center', justify='center')
text11.insert(1.0, "Mode 9 (IUMPR) - Graph / Table: ", 'center')
text11.config(font=("Arial", 8), fg="black", bg="white", state="disabled")
text11.place(x=458, y=402, width=301, height=18)

text12 = Text(root)
text12.tag_configure('center', justify='center')
text12.insert(1.0, "IUMPR Evaluation - Graph / Table: ", 'center')
text12.config(font=("Arial", 8), fg="black", bg="white", state="disabled")
text12.place(x=458, y=425, width=301, height=18)

text13 = Text(root)
text13.tag_configure('center', justify='center')
text13.insert(1.0, "Ramzellen Evaluation - Graph: ", 'center')
text13.config(font=("Arial", 8), fg="black", bg="white", state="disabled")
text13.place(x=458, y=448, width=301, height=18)

text14 = Text(root)
text14.tag_configure('center', justify='center')
text14.insert(1.0, "OBD Radar Statistics - Graph: ", 'center')
text14.config(font=("Arial", 8), fg="black", bg="white", state="disabled")
text14.place(x=458, y=471, width=301, height=18)

text15 = Text(root)
text15.tag_configure('center', justify='center')
text15.insert(1.0, "P_Code Database Update: ", 'center')
text15.config(font=("Arial", 8), fg="black", bg="white", state="disabled")
text15.place(x=458, y=517, width=301, height=18)

##### LIST OF CHECKBOXES ####
##### Radio Button Selections for Basic Settings #####
# Create the two Radiobuttons for VDS and VIN selection
option_var1 = tk.IntVar(value=None)
radio1 = tk.Radiobutton(root, text="VDS   ", variable=option_var1, value=1, anchor='w')
radio1.config(font=("Arial", 8), fg="black", bg="white")
radio1.place(x=694, y=206, width=107, height=15)
radio2 = tk.Radiobutton(root, text="VIN  ", variable=option_var1, value=2, anchor='w')
radio2.config(font=("Arial", 8), fg="black", bg="white")
radio2.place(x=809, y=206, width=107, height=15)

# Create the two Radiobuttons for DataType-1 and DataType-2 selection
option_var2 = tk.IntVar(value=None)
radio3 = tk.Radiobutton(root, text="Data Type 1", variable=option_var2, value=1, anchor='w')
radio3.config(font=("Arial", 8), fg="black", bg="white")
radio3.place(x=694, y=225, width=107, height=15)
radio4 = tk.Radiobutton(root, text="Data Type 2", variable=option_var2, value=2, anchor='w')
radio4.config(font=("Arial", 8), fg="black", bg="white")
radio4.place(x=809, y=225, width=107, height=15)

# Create the two Radiobuttons for PowerPoint Creation
option_var3 = tk.IntVar(value=None)
radio5 = tk.Radiobutton(root, text="Yes   ", variable=option_var3, value=1, anchor='w')
radio5.config(font=("Arial", 8), fg="black", bg="white")
radio5.place(x=694, y=244, width=107, height=15)
radio6 = tk.Radiobutton(root, text="No    ", variable=option_var3, value=2, anchor='w')
radio6.config(font=("Arial", 8), fg="black", bg="white")
radio6.place(x=809, y=244, width=107, height=15)

# Create the two Radiobuttons for PowerPoint Creation
option_var4 = tk.IntVar(value=None)
radio7 = tk.Radiobutton(root, text="Km     ", variable=option_var4, value=1, anchor='w')
radio7.config(font=("Arial", 8), fg="black", bg="white")
radio7.place(x=694, y=263, width=107, height=15)
radio8 = tk.Radiobutton(root, text="Datum", variable=option_var4, value=2, anchor='w')
radio8.config(font=("Arial", 8), fg="black", bg="white")
radio8.place(x=809, y=263, width=107, height=15)

##### Checkbox selections for Evaluation Graph Table Function Usage #####

# Create the "Select All" checkbox
select_all_checkbox_var = tk.IntVar()
select_all_checkbox = tk.Checkbutton(root, text="Select All", variable=select_all_checkbox_var, anchor="w")
select_all_checkbox.config(font=("Arial", 8), fg="black", bg="white")
select_all_checkbox.place(x=769, y=312, width=147, height=18)

# Create the other checkboxes
checkbox_vars = [tk.IntVar() for i in range(8)]
checkboxes = []
for i, checkbox_var in enumerate(checkbox_vars):
    checkbox = tk.Checkbutton(root, text="Check for Enable", variable=checkbox_var, anchor="w")
    checkbox.config(font=("Arial", 8), fg="black", bg="white")
    if i == 7:
        checkbox.place(x=769, y=517, width=147, height=18)
    else:
        checkbox.place(x=769, y=335 + i * 23, width=147, height=18)
    checkboxes.append(checkbox)


def select_all_callback():
    """
    Add a callback function to the "Select All" checkbox
    :return:
    """
    if select_all_checkbox_var.get():
        for checkbox_var in checkbox_vars:
            checkbox_var.set(1)
    else:
        for checkbox_var in checkbox_vars:
            checkbox_var.set(0)


select_all_checkbox.config(command=select_all_callback)

##### LIST OF BUTTONS ####
button_importDiagra = tk.Button(root, text="Import DiagRA Shots", command=lambda: bt.import_diagra_shots(root, list1, label2))
button_importDiagra.place(x=20, y=565, width=124, height=28)

button_DeleteSelectedFiles = tk.Button(root, text="Delete Shot(s)", command=lambda: bt.delete_diagra_shot(list1, label2))
button_DeleteSelectedFiles.place(x=167, y=565, width=124, height=28)

button_DeleteAllShots = tk.Button(root, text="Delete All Shots", command=lambda: bt.delete_all_shots(list1, label2))
button_DeleteAllShots.place(x=314, y=565, width=124, height=28)

button_Start = tk.Button(root, text="Start", command=lambda: bt.button_start(list2,
                                                                             button_Start, button_importDiagra, button_DeleteSelectedFiles, button_DeleteAllShots,
                                                                             button_SelectExcel, button_SavePath, button_SavePathDefault,
                                                                             button_Reset, button_Stop,
                                                                             option_var1, option_var2, option_var3, option_var4,
                                                                             checkbox_vars))
button_Start.place(x=458, y=565, width=124, height=28)

button_Reset = tk.Button(root, text="Reset")
button_Reset.place(x=625, y=565, width=124, height=28)
button_Reset.config(state="disabled")

button_Stop = tk.Button(root, text="Stop")
button_Stop.place(x=792, y=565, width=124, height=28)
button_Stop.config(state="disabled")

button_SelectExcel = tk.Button(root, text="Select Excel File (.xlsx)", command=lambda: bt.select_excel(text1))
button_SelectExcel.place(x=458, y=66, width=124, height=30)

button_SavePath = tk.Button(root, text="Select Save Folder ", command=lambda: bt.select_save_folder(text2))
button_SavePath.place(x=458, y=148, width=184, height=30)

button_SavePathDefault = tk.Button(root, text="Set Save Folder as Default", command=lambda: bt.set_save_default(text2))
button_SavePathDefault.place(x=732, y=148, width=184, height=30)

root.mainloop()
