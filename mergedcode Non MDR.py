import re
import glob
import os
import shutil
import openpyxl

def ReplicateReportingTemplate():
    # import os
    # import shutil
    # import openpyxl
    os.makedirs('C:\\Reporting tools\\Non MDR tools\\', exist_ok=True)
    list_of_chosen_files = open(
        "C:\\Reporting tools\\Non MDR list of facilities.txt").read().splitlines()
    for i in range(len(list_of_chosen_files)):
        shutil.copy('C:\\Reporting tools\\Non MDR template.xlsx',
                    'C:\\Reporting tools\\Non MDR tools\\' + list_of_chosen_files[i] + ".xlsx")


ReplicateReportingTemplate()


def ChangeTemplateDetails():
    # import os
    # import re
    # import glob
    from openpyxl import Workbook
    from datetime import datetime

    path = 'C:\\Reporting tools\\Non MDR tools\\'
    files = [f for f in glob.glob(path + "**/*.xlsx", recursive=True)]
    for f in files:
        facilityname = os.path.splitext(os.path.basename(f))[0]
        import openpyxl
        excel_file = openpyxl.load_workbook(f)
        excel_sheet = excel_file.active

        excel_sheet.freeze_panes = None
        excel_sheet.sheet_view.topLeftCell = 'A1'
        c = excel_sheet['E11']
        excel_sheet.freeze_panes = c

        excel_sheet.protection.password = 'usaidetb'
        # excel_sheet.protection.disable()

        # Change details
        #excel_sheet['K2'] = datetime.strptime('06-2020', '%m-%Y')
        #excel_sheet['Z2'] = '02/12/2020'
        #excel_sheet['B8'] = "TB Diagnosis (March-2020)"
        #excel_sheet['B19'] = "TB Notifications (March-2020)"
        #excel_sheet['B27'] = "TB/HIV Collaboration (March-2020)"
        #excel_sheet['B54'] = "Cohort at 2 Months (Smear Conversion): December 2019"
        #excel_sheet['B61'] = "TB Treatment Outcomes (March 2019)"
        #excel_sheet['B77'] = "DR-TB NOTIFICATIONS (March 2020)"

        prov = excel_sheet['L4']
        dist = excel_sheet['L3']
        fac = excel_sheet['D2']

        # Update with correct facility details
        facility = facilityname.partition('_')[2]
        fac.value = facility

        district_ = re.search('-(.*)_', facilityname)
        if district_.group(1).strip():
            district = district_
            dist.value = district[1]

        province_ = re.search('(.*)-', facilityname)
        if 'CP' in province_.group(1):
            province = 'Central'
            prov.value = province

        elif 'CBP' in province_.group(1):
            province = 'Copperbelt'
            prov.value = province

        elif 'LP' in province_.group(1):
            province = 'Luapula'
            prov.value = province

        elif 'MP' in province_.group(1):
            province = 'Muchinga'
            prov.value = province

        elif 'NP' in province_.group(1):
            province = 'Northern'
            prov.value = province

        elif 'NWP' in province_.group(1):
            province = 'North-Western'
            prov.value = province

        # rename worksheet with facility name
        #excel_sheet.title = facility

        excel_sheet.protection.sheet = True
        excel_sheet.protection.enable()

        # Save changes to the workbook
        excel_file.save(f)


ChangeTemplateDetails()


def movefilestoprovince():
    path = 'C:\\Reporting tools\\Non MDR tools\\'
    os.makedirs('C:\\Reporting tools\\Provinces\\Central\\', exist_ok=True)
    os.makedirs('C:\\Reporting tools\\Provinces\\Copperbelt\\', exist_ok=True)
    os.makedirs('C:\\Reporting tools\\Provinces\\Luapula\\', exist_ok=True)
    os.makedirs('C:\\Reporting tools\\Provinces\\Muchinga\\', exist_ok=True)
    os.makedirs('C:\\Reporting tools\\Provinces\\Northern\\', exist_ok=True)
    os.makedirs('C:\\Reporting tools\\Provinces\\North western\\', exist_ok=True)
    files = [f for f in glob.glob(path + "**/*.xlsx", recursive=True)]
    for f in files:
        facilityname = os.path.splitext(os.path.basename(f))[0]
        province_ = re.search('(.*)-', facilityname)
        if 'CP' in province_.group(1):
            destination = 'C:\\Reporting tools\\provinces\\Central\\'
            shutil.move(f, destination)

        elif 'CBP' in province_.group(1):
            destination = 'C:\\Reporting tools\\provinces\\Copperbelt\\'
            shutil.move(f, destination)

        elif 'LP' in province_.group(1):
            destination = 'C:\\Reporting tools\\provinces\\Luapula\\'
            shutil.move(f, destination)

        elif 'MP' in province_.group(1):
            destination = 'C:\\Reporting tools\\provinces\\Muchinga\\'
            shutil.move(f, destination)

        elif 'NP' in province_.group(1):
            destination = 'C:\\Reporting tools\\provinces\\Northern\\'
            shutil.move(f, destination)

        elif 'NWP' in province_.group(1):
            destination = 'C:\\Reporting tools\\provinces\\North western\\'
            shutil.move(f, destination)


movefilestoprovince()

import os

cpt = sum([len(files) for r, d, files in os.walk('C:\\Reporting tools\\provinces\\')])
print(cpt)
