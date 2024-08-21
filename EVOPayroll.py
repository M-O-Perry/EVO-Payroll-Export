from ShiftStatus import ShiftReport
from TimeSheet import TimeSheetReport
from results import ReportOutput

from tkinter import simpledialog, messagebox
import tkinter as tk
import os
import time
from datetime import date as dt


def getInputs():
    root = tk.Tk()
    root.withdraw()
    #root.iconbitmap(default="payrollExport1.ico")
    
    dateIsGood = False
    
    while not dateIsGood:
        start_date = simpledialog.askstring(title = "Start Date", prompt = "What is the first date of the pay cycle?")
        if start_date is None:
            os._exit(0)
        
        end_date = simpledialog.askstring(title = "End Date", prompt = "What is the last date of the pay cycle?")
        if end_date is None:
            os._exit(0)
        
        
        dateIsGood = check_date(start_date) and check_date(end_date)
        
        if not dateIsGood:
            messagebox.showerror("Invalid Date", "Invalid date format entered.\nPlease enter date in MM/DD/YY or MMDDYY format.")
            continue
        
        start_date = format_date(start_date)
        end_date = format_date(end_date)
        
        try:
            if not ( int(start_date[4:]) < int(end_date[4:]) or 
                    (int(start_date[4:]) == int(end_date[4:]) and int(start_date[:2]) < int(end_date[:2])) or 
                    (int(start_date[4:]) == int(end_date[4:]) and int(start_date[:2]) == int(end_date[:2]) and int(start_date[2:4]) <= int(end_date[2:4]))
                    ):
                
                messagebox.showerror("Invalid Date", "Start date must be before end date.")
                dateIsGood = False
                continue
        except:
            messagebox.showerror("Invalid Date", "Invalid date format entered.\nPlease enter date in MM/DD/YY or MMDDYY format.")
            dateIsGood = False
            continue
                
                
    return start_date, end_date
        
def check_date(date):
    date = date.replace("/", "-")
    date = date.replace(" ", "-")
    
    if date.count("-") == 0 :
        if len(date) == 6 or len(date) == 8:
            date = date[:2] + "-" + date[2:4] + "-" + date[4:]
        elif len(date) == 4:
            date = date[:1] + "-" + date[1:2] + "-" + date[2:]
        else:
            print("length incorrect")
            return False
    
    
    
    date = date.split("-")
    
    if not date[0].isnumeric() or not date[1].isnumeric() or not date[2].isnumeric():
        print("not numeric")
        return False
    
    if int(date[0]) > 12 or int(date[0]) < 1:
        print("month out of range")
        return False
    
    if int(date[1]) > 31 or int(date[1]) < 1:
        print("day out of range")
        return False
    
    if int(date[2]) > dt.today().year%100 and int(date[2]) < 90:
        print("year out of range")
        return False
    
    
        
    return True

def format_date(date):
    date = date.replace("/", "-")
    date = date.replace(" ", "-")
    
    if date.count("-") == 0 :
        if len(date) == 6 or len(date) == 8:
            date = date[:2] + "-" + date[2:4] + "-" + date[4:]
        elif len(date) == 4:
            date = date[:1] + "-" + date[1:2] + "-" + date[2:]
        else:
            raise ValueError("Invalid date format")
    date = date.split("-")
    
    if len(date[0]) == 1:
        date[0] = "0" + date[0]
    
    if len(date[1]) == 1:
        date[1] = "0" + date[1]
    
    if len(date[2]) == 4:
        date[2] = date[2][2:4]
    
    date = "".join(date)
    
    return date


startTime = time.time()
start_date, end_date = getInputs()

print("Start Date: ", start_date)
print("End Date: ", end_date)


payrollReport = TimeSheetReport(start_date, end_date)
shiftReport = ShiftReport(start_date, end_date)

shiftReport.export_labor()
shiftReport.parse_labor()
shiftReport.organize_labor()

payrollReport.export_timesheet()
payrollReport.parse_timesheet()

report = ReportOutput(payrollReport.employeeHours, shiftReport.sumEmployeeEntries, start_date, end_date)

report.write_to_excel()
report.formatExcelOutput()
# report.print_all()

# report.createADPFile()
# report.addAllEmployeesToADP()


print("Done")
messagebox.showinfo("Success", "The program has completed successfully.")