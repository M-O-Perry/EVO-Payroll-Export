from PlayActions import send_keys as send
import xlrd
import os
import glob
from EVOUtil import openTASProgram

class TimeSheetReport:  # Automatically Clocked times
    def __init__(self, start_date, end_date):
        self.employeeHours = {}
        self.start_date = start_date
        self.end_date = end_date
    
    def export_timesheet(self):
        save_directory = "\\\\FS2\\engineer\\WORK\\Outdwg\\SolidworksBomOutputs\\"

        openTASProgram("DCD")
        send(["tab 4", "space", "tab 3", self.start_date, "tab", self.end_date, "enter", "alt p",0.5, "alt o", 5])
        send(["alt f", "tab 2", "ctrl delete", "E", "tab", save_directory, "alt o", 3, "enter", "alt x"])


    def parse_timesheet(self, reportFile = ""):
        if reportFile == "":
            list_of_files = glob.glob("\\\\FS2\\engineer\\WORK\\Outdwg\\SolidworksBomOutputs\\*.xls")
            reportFile = max(list_of_files, key=os.path.getctime)
        
        wb = xlrd.open_workbook(reportFile)
        ws = wb.sheet_by_index(0)
        
        for row in range(2, ws.nrows):
            emp = str(ws.cell_value(row, 2)).strip()
            
            if not emp.isnumeric():
                continue
            
            date = ws.cell_value(row, 0)
            runHours = float(ws.cell_value(row, 8))
            name = ws.cell_value(row, 3)
            
            if emp in self.employeeHours:
                if date in self.employeeHours[emp]:
                    self.employeeHours[emp][date] += runHours
                else:
                    self.employeeHours[emp][date] = runHours
        
            else:
                self.employeeHours[emp] = {"name": name, date: runHours}
                
        # os.remove(reportFile)
