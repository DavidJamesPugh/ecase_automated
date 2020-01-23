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

# # Buttons Dictionaries. All functions are in button_functions module.
ecase_button_dict = {
    'Doctor Allocation Numbers': lambda: button_functions.doctor_allocations(),
    'Download eCase Bowel Report': lambda: button_functions.bowel_files(),
    'Download Care Plans Audits': lambda: button_functions.ecase_care_plans(),
    'Resident Care Level List': lambda: button_functions.podiatry_list(),
    'Download eCase Data and Import': lambda: button_functions.ecase_data_download(),
    'PI Risk Levels': lambda: button_functions.pi_risks(),
    'Print Resident Files': lambda: button_functions.printing_files(),
    'Resident Birthday List': lambda: button_functions.resident_birthdays(),
    'Resident Temp Movements List': lambda: button_functions.temp_movements(),
}

staff_button_dict = {
    'Mandatory Training List Update': lambda: button_functions.mand_training(),
    'Print Clinical Admission Files': lambda: button_functions.print_clin_files(),
    'Print Staff Birthdays': lambda: button_functions.staff_birthdays(),
    'Walls and Roche stats': lambda: button_functions.walls_roche()
}

# Placing SAV logo to the window
panel = tkinter.Label(root, image=img)
panel.image = img
panel.grid(row=0, column=1, columnspan=2, padx=10)

# Placing buttons and organising into the two frames.
# The declare_buttons function establishes a button for each entry
# in the dictionary given
# Row is relative to the below order
eCase_frame.grid(row=1, column=0, columnspan=4, padx=25, pady=3)
button_functions.declare_buttons(eCase_frame, ecase_button_dict, 1)

staff_frame.grid(row=2, column=0, columnspan=4)
button_functions.declare_buttons(staff_frame, staff_button_dict, 1)

tkinter.Button(root,
               text='Help',
               command=lambda: os.startfile(r'Documentation\eCase Help.docx')
               ).grid(row=3, column=1, pady=5)
tkinter.Button(root,
               text='Quit',
               command=lambda: root.destroy()
               ).grid(row=3, column=2, pady=5)

root.mainloop()
