#! Collateral_Assignments.py - a script that will read the names of personnel
#  and a list of collateral duties from separate Excel spreadsheets, and will
#  ask the user for primary and secondary assignments for each duty. Once selected
#  the system will print out an appropriate appointment letter for each duty.

import calendar
import pyinputplus as pyip
import openpyxl
import docx
import glob
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.style import WD_STYLE_TYPE
from datetime import date


def choose_workbook(message):
    """Asking the user to clarify which excel correlates to personnel or duties list"""
    files = glob.glob('*xlsx')
    print(message)
    output = pyip.inputMenu(files, numbered=True)
    return output


def get_personnel():
    """Reading the current personnel from the provided Excel sheet."""
    file = choose_workbook("Which Excel has the personnel of the department?")
    wb = openpyxl.load_workbook(file)
    sheet = wb.active
    for row in range(1, sheet.max_row + 1):
        personnel.append(sheet[f'A{row}'].value)
    wb.close()

def get_duties():
    """Reading the current duties from the provided Excel sheet."""
    file = choose_workbook("Which Excel lists the collateral duties?")
    wb = openpyxl.load_workbook(file)
    sheet = wb.active
    for row in range(1, sheet.max_row + 1):
        duties.append(sheet[f'A{row}'].value)
    wb.close()

def get_current_date():
    """Get current date with goal format of: dd MMM YYYY"""
    date_year = date.today().year
    date_month = date.today().month
    month_abbr = calendar.month_abbr[date_month]
    date_day = date.today().day
    return f'{date_day} {month_abbr} {date_year}'

def get_officers_name():
    """Getting officer's name with double-check before continuing on."""
    correct_officer = False

    while correct_officer is False:
        output = input("Enter the name of the officer who's signing the letters: ")
        print(f"You entered: {output}. Is this correct?")
        double_check = pyip.inputMenu(['Yes', 'No'], numbered=True)
        if double_check == 'Yes':
            correct_officer = True

    return output

def write_letter(duty, member, holder, officer):
    """Writing the actual appointment letter with the individualized duty, member, and current date."""
    letter = docx.Document()
    styles = letter.styles
    style = styles.add_style('Times New Roman', WD_STYLE_TYPE.PARAGRAPH)
    style.font.name = 'Times New Roman'
    style.font.size = Pt(12)

    date_paragraph = letter.add_paragraph(get_current_date())
    date_paragraph_format = date_paragraph.paragraph_format
    date_paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
    date_paragraph.style = letter.styles['Times New Roman']

    memorandum_paragraph = letter.add_paragraph()
    memorandum_paragraph.add_run('MEMORANDUM').bold = True
    memorandum_paragraph.style = letter.styles['Times New Roman']

    from_paragraph = letter.add_paragraph(f"From: Division Officer, Pharmacy Department, Navy Medicine Readiness and "
                                          f"Training\n           Command Lemoore\nTo:     {member}, USN\nVia:    "
                                          f"Leading Chief Petty Officer, Pharmacy Department")
    from_paragraph.style = letter.styles['Times New Roman']

    sub_paragraph = letter.add_paragraph(f"Subj:  APPOINTMENT AS {holder.upper()} {duty.upper()}")
    sub_paragraph.style = letter.styles['Times New Roman']

    ref_paragraph = letter.add_paragraph("Ref:    (a) NMRTC PHARMACY DEPARTMENT COLLATERAL DUTY \n\t    EXPECTATIONS\n"
                                         "           (b) NMRTC PHARMACY DEPARTMENT PERSONAL QUALIFICATION\n           "
                                         "     STANDARDS")
    ref_paragraph.style = letter.styles['Times New Roman']

    bullet_one_paragraph = letter.add_paragraph()
    bullet_one_paragraph.add_run(f"1.  Effective immediately, you are hereby appointed as the {duty} for the Pharmacy "
                                 f"Department at Navy Medicine Readiness and Training Command Lemoore.  You will be "
                                 f"guided in the conduct of your duties by reference (a) and (b) and directly "
                                 f"responsible to the Division Officer.")
    bullet_one_paragraph.style = letter.styles['Times New Roman']

    bullet_two_paragraph = letter.add_paragraph()
    bullet_two_paragraph.add_run("2.  Failure to properly execute assigned duties will result in remedial training "
                                 "from the department Leading Petty Officer for correction.  If remedial corrective "
                                 "actions prove unsuccessful, it will be the recommendation of leadership to forward "
                                 "to Leading Petty Officer, and or Senior Enlisted Leader for further disciplinary "
                                 "review and actions.")
    bullet_two_paragraph.style = letter.styles['Times New Roman']

    bullet_three_paragraph = letter.add_paragraph()
    bullet_three_paragraph.add_run("3.  This appointment will remain in effect until reassignment, unless otherwise "
                                   "directed.")
    bullet_three_paragraph.style = letter.styles['Times New Roman']

    spacing_paragraph = letter.add_paragraph()
    second_spacing_paragraph = letter.add_paragraph()
    signing_officer_paragraph = letter.add_paragraph()
    signing_officer_paragraph.add_run(f"\t\t\t\t\t\t\t{officer.upper()}")

    another_spacing_paragraph = letter.add_paragraph()
    copy_to_paragraph = letter.add_paragraph()
    copy_to_paragraph.add_run("Copy to:\nDIVO FOLDER")

    letter.save(f"{duty} Appointment Letter-{holder} {member}.docx")

# Setting up blank lists to be filled in from current Excel sheets.
personnel = []
duties = []

# Another list in-case there are duties that aren't yet filled.
come_back_to_duties = []

# Populate lists from Excel.
get_personnel()
get_duties()

# Set officer's name who's signing the letters
signing_officer = get_officers_name()

# Run loop until every duty, both primary and secondary, is accounted for.
# for duty in duties:
#     print(f"Who's going to be the primary for: {duty}?")
#     primary_holder = pyip.inputMenu(personnel, numbered=True)
#     holder = "PRIMARY"
#     write_letter(duty, primary_holder, holder, signing_officer)

write_letter('Crash Cart Coordinator', 'HM2 Kemp', 'PRIMARY', signing_officer)


