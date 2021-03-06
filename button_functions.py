"""
    Defines the below functions to be used in __init__.py as buttons

"""

import os
import re
import time
import tkinter

from selenium.common.exceptions import NoSuchElementException

import constants
import ecase_data_import
import ecase_downloader
import printing_documents
import staff_docs
import wr_meds


# ###############################################
# ######################eCase Automated Reports##
# ###############################################
def doctor_allocations():
    """
    Opens a selenium browser with ecase_login,
    and downloads eCase reports, then creates two reports,
    one with a list of residents and who their doctor is.
    The second file is a summary of each doctor
    and how many resident’s they are looking after
"""

    doctors = rf'{constants.DOWNLOADS_DIR}\doctor_Allocation_Numbers.xlsx'
    sort_doc = rf'{constants.DOWNLOADS_DIR}\sorted_doctor_Numbers.csv'

    # ##Error checking for if the file is open already.
    # ##Calls the function file_available to save space, rather than
    # ##try and except blocks.
    if file_available(doctors) and file_available(sort_doc):
        ecase_driver = ecase_downloader.ecase_login()
        ecase_downloader.doctor_numbers_download(ecase_driver)
        time.sleep(3)
        ecase_driver.quit()
        ecase_data_import.doctor_numbers()


def bowel_files():
    """
    Opens a new GUI window with two buttons for the below two functions
"""
    nhi_window = tkinter.Tk()
    nhi_window.wm_title("Bowel Records Window")

    tkinter.Button(nhi_window, text="Current Month Bowel Records",
                   command=lambda: ecase_bowel_report(0)).grid(
        row=2, pady=5, column=2)
    tkinter.Button(nhi_window, text="Previous Month's Bowel Records",
                   command=lambda: ecase_bowel_report(1)).grid(
        row=3, column=2, padx=20)
    tkinter.Button(nhi_window, text="Quit",
                   command=lambda: nhi_window.destroy()).grid(
        row=4, column=2, pady=5)


def ecase_bowel_report(age: int):
    """
    Opens a selenium browser with ecase_login,
    and downloads the bowel eCase reports.
    Creates an excel file with a sheet for each area,
    for this month’s Resident’s bowel records
    """

    wings = ['HOUSE 1 - Hector', 'HOUSE 2 - Marion Ross',
             'HOUSE 3 - Bruce', 'HOUSE 4 - Douglas',
             'HOUSE 5 - Henry Campbell',
             'Stirling', 'Iona', 'Balmoral', 'Braemar']

    ecase_driver = ecase_downloader.ecase_login()
    ecase_data_import.bowel_setup()

    for wing in wings:
        ecase_downloader.main_bowel_report(ecase_driver, wing, age)
        ecase_data_import.bowel_import(wing)

    ecase_driver.quit()
    ecase_data_import.bowel_report_cleanup()


def ecase_care_plans():
    """
    Opens a selenium browser with ecase_login,
    and downloads eCase reports, creates an excel file with a
    sheet for each area, showing careplans for each resident.
    
    """

    wings = ['HOUSE 1 - Hector', 'HOUSE 2 - Marion Ross',
             'HOUSE 3 - Bruce', 'HOUSE 4 - Douglas',
             'HOUSE 5 - Henry Campbell',
             'Stirling', 'Iona', 'Balmoral', 'Braemar']

    if file_available(rf'{constants.OUTPUTS_DIR}\Care Plans\eCaseCareplans.xlsx'):
        ecase_driver = ecase_downloader.ecase_login()
        ecase_data_import.care_plans_setup()

        for wing in wings:
            try:
                ecase_downloader.care_plan_audits_download(ecase_driver, wing)
                ecase_data_import.care_plans_import(wing)
            except NoSuchElementException:
                print(f'{wing} care plans could not be downloaded')

        ecase_driver.quit()
        ecase_data_import.care_plans_missing_audits()


def podiatry_list():
    """
    Opens a selenium browser, with ecase_login, and prints, or opens,
    and excel file with a list of Resident’s and their care levels per area

    """
    ecase_driver = ecase_downloader.ecase_login()
    ecase_downloader.care_level_csv(ecase_driver)
    time.sleep(1.5)
    ecase_driver.quit()
    ecase_data_import.care_level_list()


def ecase_data_download():
    r"""
    Opens a selenium browser with ecase_login,
    and downloads eCase data reports, then writes this data into
    eCaseData.xlsx in J:\Quality Data\Clinical Data.
    eCaseGraphs.xlsx has a collection of pivot tables to analyse this data.
    """
    if file_available(rf'{constants.MAIN_DATA_DIR}\Clinical Data\eCaseData.xlsx'):
        ecase_driver = ecase_downloader.ecase_login()
        try:
            ecase_downloader.ecase_data(ecase_driver)
        except NoSuchElementException:
            print("Data report can't be downloaded")

        ecase_driver.quit()

        try:
            ecase_data_import.ecase_data_import()
        except FileNotFoundError:
            pass


# #####################################
# ######################Printing Files#
# #####################################
def printing_files():
    """
    Opens a new window with a text entry for NHI number,
    and three buttons, one for each of the following functions 
    
    """

    nhi_window = tkinter.Tk()
    nhi_window.wm_title("NHI Entry Window")
    tkinter.Label(nhi_window, text="Please enter NHI here:").grid(row=0, column=2)
    nhi_entry = tkinter.Entry(nhi_window)
    nhi_entry.grid(row=1, column=1, columnspan=3, padx=50, pady=5)

    declare_buttons(nhi_window,
                    {"Print Resident Front Sheet": lambda: front_sheet(nhi_entry),
                     "Print RLV Front Sheet": lambda: front_sheet(nhi_entry, village=True),
                     "Print Nurses Front Sheet": lambda: front_sheet(nhi_entry, nurses=True),
                     "Print Door Label": lambda: door_label(nhi_entry),
                     "Create Labels List": lambda: label_list(nhi_entry),
                     "Quit": lambda: nhi_window.destroy()}, 2)


def front_sheet(entry, village=False, nurses=False):
    """
    Opens a selenium browser with ecase_login, and downloads eCase reports,
    resident_Image, and preferred_Name. Uses this to create a formatted
    excel file for Admissions manager and Accountants, with Resident info and
    Resident Contacts.
    :param entry: tkinter Entry object
    :param village: Bool to print only 2 copies of the front sheet, if true
    :param nurses:
    :return:
    """
    printing_resident_sheets(entry, rf'{constants.OUTPUTS_DIR}\front_sheet.xlsx')
    printing_documents.create_front_sheet(village=village, nurses=nurses)


def door_label(entry):
    """
    Opens a selenium browser with ecase_login,
    and downloads eCase reports with resident_Contacts.
    Downloads the resident_Image, and preferred_Name,
    then uses create_Door_Label to create a formatted excel
    file with the resident’s name and photo to place on their door
    """
    printing_resident_sheets(entry, rf'{constants.OUTPUTS_DIR}\door_label.xlsx')
    printing_documents.create_door_label()


def label_list(entry):
    """
    Opens a selenium browser with ecase_login,
    and downloads eCase reports with resident_Contacts.
    Gets the preferred_Name, and then create_Label_List is
    called to generate a formatted excel file to print sticky labels
    """
    printing_resident_sheets(entry, rf'{constants.OUTPUTS_DIR}\label_sheet.xlsx')
    printing_documents.create_label_list()


# #########################################
# ######################Printing Files End#
# #########################################
def pi_risks():
    """
    Creates a file of all resident's Risk factor. The file pir_code.csv needs
    to be manually generated and cleaned, as the natural pir_code.csv from
    eCase has too many duplicates
    """
    ecase_driver = ecase_downloader.ecase_login()
    ecase_downloader.ecase_pi_risk(ecase_driver)
    printing_documents.pi_risk_levels(ecase_driver)
    ecase_driver.quit()


# #########################################
# ######################Birthdays Printout#
# #########################################
def resident_birthdays():
    """
    Opens a new window with two buttons, one for the following two modules
    """
    birthday_window = tkinter.Tk()
    birthday_window.wm_title("Birthdays")

    tkinter.Button(birthday_window, text="Resident Birthdays List",
                   command=lambda: resident_birthday_list()).grid(
        row=2, column=2, padx=10, pady=5)
    tkinter.Button(birthday_window, text="Village Birthdays",
                   command=lambda: resident_birthday_list(only_village=True)).grid(
        row=3, column=2)
    tkinter.Button(birthday_window, text="Quit",
                   command=lambda: birthday_window.destroy()).grid(
        row=5, column=2, pady=10)


def resident_birthday_list(only_village=False):
    """
    Opens a selenium browser with ecase_login, and downloads eCase reports
    with eCase_Birthdays. Then creates an excel file with all current residents
    and their birthdates
    """

    if file_available(rf'{constants.OUTPUTS_DIR}\Resident Birthdays\ResidentBirthdays.xlsx'):
        ecase_driver = ecase_downloader.ecase_login()
        ecase_downloader.ecase_birthdays(ecase_driver)
        time.sleep(4)
        ecase_driver.quit()
        printing_documents.village_birthdays(only_village=only_village)


# #############################################
# ######################Birthdays Printout End#
# #############################################
def temp_movements():
    r"""
    Opens a selenium browser with ecase_login,
    and downloads eCase reports with eCase_Movements, then
    appends new temporary movements to G:\eCase\Downloads\eCaseTempMoves.xlsx
    """

    if file_available(rf'{constants.OUTPUTS_DIR}\eCaseTempMoves.xlsx'):
        ecase_driver = ecase_downloader.ecase_login()
        ecase_downloader.ecase_movements(ecase_driver)
        # waits for the download to finish
        time.sleep(3)
        ecase_driver.quit()
        printing_documents.temp_movements_print()


# #############################################
# ######################Staff Docs Required#
# #############################################
def mand_training():
    r"""
    Takes csv file from MYOB, and updates the Training to be book.xlsx
    file with all mandatory training. In 'J:\Quality Data\Training'. For Jane
    """
    try:
        staff_docs.training_lists()

    except FileNotFoundError:
        popup_error(r"""
Please Generate the file first from MYOB Payroll under the
Employees report and name it Birthday.csv.
Place in J:\Quality Data\Data Technician\StaffDbases\n
The File should have fields
- Employee Code
- Employee First Name
- Employee Last Name
- Training Name
- Date Booked""")

    except PermissionError:
        popup_error("The file is being used by someone else")


def print_clin_files():
    r"""
    Prints all files from 'J:\Forms and Standard letters\
    Clinical Files for Admissions\For Clinical File' Folder
    """

    # Creates a list of all files in clinical_directory ending in .docx
    files = [file for file in os.listdir(constants.ADMISSION_DIR) if file.endswith('.docx')]

    # Prints all files in the files list
    for file in files:
        os.startfile(rf'{constants.ADMISSION_DIR}\{file}', 'print')


def staff_birthdays():
    """
    Takes a csv file from MYOB and creates
    a formatted file for this and the next two months. 
    """
    try:
        staff_docs.birthday_list()

    except FileNotFoundError:
        popup_error(r"""Please Generate the file first from MYOB Payroll 
        under the Employees reportand name it Birthday.csv.
Place in J:\Quality Data\Data Technician\StaffDbases\n
The File should have fields
- Employee Code
- Employee Full Name
- Employee Status
- Employee Occupation
- Employee Start Date
- Employee Birthdate
- Employee Cost Centre Name""")


def walls_roche():
    """
    Takes a csv file, copy pasted from the Walls&Roche medication PDF
    (sourced from Clinical managers), and formats it for pivot table use
    """
    try:
        wr_meds.meds_counts()

    except FileNotFoundError:
        popup_error(r"""
Please extract PDF info into J:\Quality Data\Data Technician\Walls
and Roche, \nand name WRMedication.xlsx""")
    except KeyError:
        popup_error('Please name the sheet in the file "Sheet1"')


# #############################################
# ######################Staff Docs Required End#
# #############################################
def file_available(file: str):
    """
    Checks whether the file given in the argument is open or not.
    Does this by attempting to rename it quickly, and renaming back to
    what it was originally
    """
    import os

    if os.path.isfile(file):
        try:
            os.rename(file, rf'{os.path.dirname(file)}\tempfile.xlsx')
            os.rename(rf'{os.path.dirname(file)}\tempfile.xlsx', file)
            return True

        except OSError:
            popup_error(f'{os.path.basename(file)} is open by someone else '
                        f'and cannot be used.')
            return False
    else:
        return True


def popup_error(msg: str):
    """
    Creates a pop up with a message (msg) on it. Used for general
    error messages. 
    """
    popup = tkinter.Tk()
    popup.wm_title("An Error has Occurred")
    label = tkinter.Label(popup, text=msg)
    label.pack(side="top", fill="x", pady=10)
    b1 = tkinter.Button(popup, text="Okay", command=popup.destroy)
    b1.pack(pady=10)
    popup.mainloop()


def printing_resident_sheets(entry, file):
    """
    Checks if the NHI is valid, then downloads the contacts file from ecase
    Downloads the preferred name and image.
    Then quits the selenium browser.
    :param entry:
    :param file:
    :return:
    """
    nhi = entry.get()
    if re.match("^[A-Za-z]{3}[0-9]{4}$", nhi):
        pass
    else:
        popup_error("Incorrect NHI format entered, please try again")

    if file_available(file):
        ecase_driver = ecase_downloader.ecase_login()
        ecase_downloader.resident_contacts(ecase_driver, nhi)
        ecase_downloader.preferred_name_and_image(ecase_driver, nhi)
        ecase_driver.quit()


def declare_buttons(window, button_dict: dict, column):
    """

    :param window: tkinter window object
    :param button_dict: dictionary of button text as keys, with the related
                        commands as values
    :param column: which column these buttons will be on
    :return:
    """
    for button in button_dict:
        if button == "Quit":
            tkinter.Button(window, text=button, command=button_dict[button]
                           ).grid(column=column, pady=7.5, padx=5)
        else:
            tkinter.Button(window, text=button, command=button_dict[button]
                           ).grid(column=column, pady=2.5, padx=5)
