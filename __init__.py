"""
Creates the GUI from tkinter of the window and buttons.
Assigns functions created in button_functions.py to buttons.

"""

import os
import tkinter

from PIL import ImageTk, Image

import button_functions

# GUI of the app, has buttons which call functions
# Spawns the GUI at postion x=125,y=100 on screen.
# Window size is 255x500
root = tkinter.Tk()
root.wm_title('eCase Report Interface')
root.geometry('%dx%d%+d%+d' % (255, 550, 125, 100))

# Loads the SAV logo to img variable
img = ImageTk.PhotoImage(Image.open(r'images\SAVLandscape.jpg').resize((200, 30)))

# Creates two frames within the window,
# for organisation of eCase vs other documents
eCase_frame = tkinter.LabelFrame(root, text='eCase', labelanchor='n')
staff_frame = tkinter.LabelFrame(root, text='Local Files Needed', labelanchor='n')

# Buttons. All functions are in button_functions module.
doctor_Numbers_B = tkinter.Button(eCase_frame, text='Doctor Allocation Numbers',
                                  command=lambda: button_functions.doctor_allocations())

Bowel_B = tkinter.Button(eCase_frame, text='Download eCase Bowel Report',
                         command=lambda: button_functions.bowel_files())

careplans_B = tkinter.Button(eCase_frame, text='Download Care Plans Audits',
                             command=lambda: button_functions.ecase_care_plans())

care_Levels_B = tkinter.Button(eCase_frame, text='Resident Care Level List',
                               command=lambda: button_functions.podiatry_list())

Data_B = tkinter.Button(eCase_frame, text='Download eCase Data and Import',
                        command=lambda: button_functions.ecase_data_download())

pi_Risks_B = tkinter.Button(eCase_frame, text='PI Risk Levels',
                            command=lambda: button_functions.pi_risks())

resident_Printing_B = tkinter.Button(eCase_frame, text='Print Resident Files',
                                     command=lambda: button_functions.printing_files())

birthday_List_B = tkinter.Button(eCase_frame, text='Resident Birthday List',
                                 command=lambda: button_functions.resident_birthdays())

temp_Movements_List_B = tkinter.Button(eCase_frame,
                                       text='Resident Temp Movements List',
                                       command=lambda: button_functions.temp_movements())

training_List_B = tkinter.Button(staff_frame,
                                 text='Mandatory Training List Update',
                                 command=lambda: button_functions.mand_training())

print_ClinicalFiles_B = tkinter.Button(staff_frame,
                                       text='Print Clinical Admission Files',
                                       command=lambda: button_functions.print_clin_files())

staff_Birthday_B = tkinter.Button(staff_frame, text='Print Staff Birthdays',
                                  command=lambda: button_functions.staff_birthdays())

walls_Roche_B = tkinter.Button(staff_frame, text='Walls and Roche stats',
                               command=lambda: button_functions.walls_roche())

help_B = tkinter.Button(root, text='Help',
                        command=lambda: os.startfile(r'Documentation\eCase Help.docx'))
quit_B = tkinter.Button(root, text='Quit', command=lambda: root.destroy())

# Placing SAV logo to the window
panel = tkinter.Label(root, image=img)
panel.image = img
panel.grid(row=0, column=1, columnspan=2, padx=10)

# Placing buttons and organising into the two frames.
# Row is relative to the below order
eCase_frame.grid(row=1, column=0, columnspan=4, padx=25, pady=3)
doctor_Numbers_B.grid(pady=3, padx=5)
Bowel_B.grid(pady=3, padx=5)
careplans_B.grid(pady=3, padx=5)
care_Levels_B.grid(pady=3, padx=5)
Data_B.grid(pady=3, padx=5)
pi_Risks_B.grid(pady=3, padx=5)
resident_Printing_B.grid(pady=3, padx=5)
birthday_List_B.grid(pady=3, padx=5)
temp_Movements_List_B.grid(pady=3, padx=5)

staff_frame.grid(row=2, column=0, columnspan=4)
training_List_B.grid(pady=3)
print_ClinicalFiles_B.grid(pady=3)
staff_Birthday_B.grid(pady=3, padx=38)
walls_Roche_B.grid(pady=3)

help_B.grid(row=3, column=1, pady=5)
quit_B.grid(row=3, column=2, pady=5)

root.mainloop()
