from PlayActions import send_keys as send
import xlrd
import os
import glob
from EVOUtil import openTASProgram


class ShiftReport:  # Manually entered shifts
    def __init__(self, start_date, end_date):
        self.employeeEntriesData = {}
        self.sumEmployeeEntries = {}
        self.start_date = start_date
        self.end_date = end_date
        
        
    def export_labor(self):
        save_directory = "\\\\FS2\\engineer\\WORK\\Outdwg\\SolidworksBomOutputs\\"
        
        openTASProgram("WOLE")
        send([self.start_date, "tab", self.end_date, "enter", "#0", "enter", "#499","enter", "alt p", 0.5, "alt y", "alt o", 14])
        
        send(["alt f", "tab 2", "ctrl delete", "E", "tab", save_directory, "alt o", 6, "enter", "alt x"])
        

    def parse_labor(self, reportFile = ""):
        if reportFile == "":
            list_of_files = glob.glob("\\\\FS2\\engineer\\WORK\\Outdwg\\SolidworksBomOutputs\\*.xls")
            reportFile = max(list_of_files, key=os.path.getctime)
        
        wb = xlrd.open_workbook(reportFile)
        ws = wb.sheet_by_index(0)
        
        for row in range(2, ws.nrows):
            
            emp = str(ws.cell_value(row, 1)).strip()
            
            if not emp.isnumeric():
                continue
            
            date = ws.cell_value(row, 0)
            type = ws.cell_value(row, 6)
            name = ws.cell_value(row, 2)
            
            
            operation = ws.cell_value(row, 10)
            hours = float(ws.cell_value(row, 11))
            
            if emp in self.employeeEntriesData:
                if date in self.employeeEntriesData[emp]:
                    self.employeeEntriesData[emp][date].append((type, operation, hours))
                else:
                    self.employeeEntriesData[emp][date] = [(type, operation, hours)]
                    
            else:
                self.employeeEntriesData[emp] = {"name": name, date: [(type, operation, hours)]}
            
        os.remove(reportFile)

    def organize_labor(self):
        for empID in self.employeeEntriesData:
            emp = self.employeeEntriesData[empID]
            
            hours_worked_per_day = {}

            reg_hours = 0
            ot_hours = 0
            dt_hours = 0
            vac_hours = 0
            sick_hours = 0
            hol_hours = 0
            personal_hours = 0
            
            for date, entries in emp.items():
                if date == "name":
                    continue
                
                for field in entries:
                    type = field[0]
                    operation = field[1]
                    hours = field[2]
                    
                    if operation == "13":
                        personal_hours += hours
                    elif type == "R":
                        reg_hours += hours
                    elif type == "O":
                        ot_hours += hours
                    elif type == "D":
                        dt_hours += hours
                    elif type == "V":
                        vac_hours += hours
                    elif type == "S":
                        sick_hours += hours
                    elif type == "H":
                        hol_hours += hours
                        
                    if type == "R" or type == "O" or type == "D":
                        if date in hours_worked_per_day:
                            hours_worked_per_day[date] += hours
                        else:
                            hours_worked_per_day[date] = hours

            self.sumEmployeeEntries[empID] = (emp["name"], reg_hours, ot_hours, dt_hours, vac_hours, sick_hours, hol_hours, personal_hours, hours_worked_per_day)

        
        