'''
Date created: 09-28-2020
Last edit: 09-29-2020
Simple script to automate loging box serial numbers and mask serial 
numbers along with datecode and time stamps into an excel sheet.
'''

print("Importing openpyxl library...")
print("Opening workbook...")

import time, openpyxl

workbook_path = "C:/Breathe_Mask/breathe_serial_log.xlsx"
sheet_name = 'SerialNumbers'

wb = openpyxl.load_workbook(workbook_path)

sheet = wb[sheet_name]

while True:                                                                             # Loop script to continuously enter SNs 
    print("Start new box.")
    box_serial = input("Input BOX serial number: ")
    if box_serial == "" or box_serial.upper() == 'Q':                                   # Breakout of loop if blank or q
        break
    else:
        mask_sns = []

        for x in range(1,13):
            mask_serial = input(f"Input MASK #{x} serial number: ")
            if mask_serial == "":
                break
            else:
                named_tuple = time.localtime()                                          # Get struct_time
                date_stamp = time.strftime("%Y-%m-%d", named_tuple)
                time_stamp = time.strftime("%H:%M:%S", named_tuple)
                mask_sns.append((date_stamp, time_stamp, box_serial, mask_serial))      # Date/Time/Box SN/Mask SN tuple appended

        for mask_sn in mask_sns:                                                        # Appends tuple rows to end of excel sheet
            sheet.append(mask_sn)
        try: 
            wb.save(workbook_path)                                                      # Saves into excel WB
            print("\nBox SN " + box_serial + " saved.\n")
            print("-"*20)
        except PermissionError:                                                         # Error handling if file is left opened
            print("\n" + "-"*30 + "ERROR" + "-"*30)
            print("Permission Error: File may be left opened, please close it.")
            workbook_backup = workbook_path.split('.', 1)[0] + "_backup_" + date_stamp + ".xlsx"
            wb.save(workbook_backup)
            print("Workbook saved as backup due to error. Please notify supervisor.")
            print("-"*65)
