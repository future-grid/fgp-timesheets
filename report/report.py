#!/usr/bin/python3

import openpyxl as op
import datetime
import os

# function will return an object with the date of the friday, and the initials of the creator
# "filepath" is expected to be of form "TimeSheet_wYYYYMMDD_FL" FL = initials
def get_data_from_title(filepath):
    title = filepath[len("timesheets\\TimeSheet_w"):-len(".xlsx")]
    initials = title[-2:]
    date = datetime.datetime.strptime(title[:-3], '%Y%m%d')

    return {"initials": initials, "date": date}

# returns a full day with arbitrary many projects / hours. uses the friday date to determine the day's date
def get_day_from_data(row,friday_date):
    # this array defines the difference between the given day and friday. eg friday-friday = 0, friday-wednesday=2
    week_order = ["Friday", "Thursday", "Wednesday", "Tuesday", "Monday", "Sunday"]
    date = friday_date - datetime.timedelta(days=week_order.index(row[0])) # row[0] is always the day of the week eg. "Tuesday"

    projects=[]
    for i in range(2, len(row)-1, 3):
        if row[i] == None:
            break
        projects.append({"project":row[i], "hours":float(row[i+1]), "description":row[i+2]})

    return {"date":date, "projects":projects}

# returns a list of the entire worksheet using very cool python stuff
def iter_rows(ws):
    for row in ws.iter_rows():
        yield [cell.value for cell in row]

def read_timesheet(filepath):
    print("Loading " + filepath)
    wb = op.load_workbook(filepath)
    ws = wb.active # get the active worksheet
    if ws.cell(row=4,column=1).value == "=A3+1":
        print("Old timesheet version!")
        ws.delete_cols(1)

    titledata = get_data_from_title(filepath)
    sheetdata = list(iter_rows(ws))
    days = []

    for i in range(3,9):
        days.append(get_day_from_data(sheetdata[i], titledata["date"]))

    return {"initials": titledata["initials"], "friday": titledata["date"], "days":days}

def load_sheets(root_path):
    directory = os.fsencode(root_path)
    sheets = []
    for file in os.listdir(directory):
        filename = os.fsdecode(file)
        if filename.endswith(".xlsx"):
            try:
                sheets.append(read_timesheet(root_path + "\\" + filename))
            except Exception as e:
                print("Could not read timesheet " + filename + "!")
                print(e)
            continue
        else:
            continue

    return sheets

def generate_project_overview(sheets):
    projects = {}
    for sheet in sheets:
        for day in sheet["days"]:
            for project in day["projects"]:
                project_title = project["project"]
                project_hours = float(project["hours"])
                if project_title in projects:
                    projects[project_title] += project_hours
                else:
                    projects[project_title] = project_hours

    return projects

def generate_employee_overview(sheets):
    employees = {}
    for sheet in sheets:
        employee = sheet["initials"]
        weekly_hours = 0
        date = str(sheet["friday"])[:-9]

        if employee not in employees:
            employees[employee] = {}
        
        for day in sheet["days"]:
            for project in day["projects"]:
                weekly_hours += project["hours"]
        
        employees[employee][date] = weekly_hours

    return employees

def get_earliest_date(sheets):
    early_date = None
    for sheet in sheets:
        for day in sheet["days"]:
            date = day["date"]
            if early_date is None:
                early_date = date
            if date < early_date:
                early_date = date
    return early_date

def get_latest_date(sheets):
    late_date = None
    for sheet in sheets:
        for day in sheet["days"]:
            date = day["date"]
            if late_date is None:
                late_date = date
            if date > late_date:
                late_date = date
    return late_date

def generate_timeline_overview(sheets):
    early_date = get_earliest_date(sheets)
    late_date = get_latest_date(sheets) + datetime.timedelta(days=1)
    current_date = early_date
    timeline = {}
    while True:
        date = str(current_date)[:-9]
        for sheet in sheets:
            employee = sheet["initials"]
            for day in sheet["days"]:
                if day["date"] == current_date:
                    if date not in timeline:
                        timeline[date] = []
                    for project in day["projects"]: # get all projects, hours and descriptions
                        timeline[date].append(employee + ", " + project["project"] + " (" + str(project["hours"]) + "): " + str(project["description"]))

        
        # iterate over the days, get all descriptions
        current_date = current_date + datetime.timedelta(days=1)
        if current_date == late_date:
            break

    return timeline

def adjust_column_widths(ws):
    for column_cells in ws.columns:
        length = max(len(str(cell.value)) for cell in column_cells)
        ws.column_dimensions[column_cells[0].column_letter].width = length * 1.2

def write_timeline_overview(ws_t, timeline_overview):
    i = 1
    for day, contents in timeline_overview.items():
        day_cell = ws_t.cell(row=i, column=1)
        day_cell.value = day

        for item in contents:
            col = (contents.index(item)+2)
            desc_cell = ws_t.cell(row=i, column=col)
            desc_cell.value = item
        
        i += 1

def generate_report(root_path):
    sheets = load_sheets(root_path)
    employee_overview = generate_employee_overview(sheets)
    project_overview = generate_project_overview(sheets)
    timeline_overview = generate_timeline_overview(sheets)

    wb = op.Workbook()
    ws_p = wb.create_sheet("Project Overview")
    ws_e = wb.create_sheet("Employee Overview")
    ws_t = wb.create_sheet("Timeline Overview")
    del wb["Sheet"]

    i = 1
    for employee, weeks in employee_overview.items():
        emp_cell = ws_e.cell(row=i, column=1)
        emp_cell.value = employee

        a = 2
        for week, hours in weeks.items():
            week_cell = ws_e.cell(row=i, column=a)
            week_cell.value = week + " - " + str(hours)
            a += 1
        i += 1

    i = 1
    for project, hours in project_overview.items():
        pro_cell = ws_p.cell(row=i, column=1)
        pro_cell.value = project

        hour_cell = ws_p.cell(row=i, column=2)
        hour_cell.value = hours
        i += 1

    write_timeline_overview(ws_t, timeline_overview)
    adjust_column_widths(ws_p)
    adjust_column_widths(ws_e)
    adjust_column_widths(ws_t)
    
    wb.save('report.xlsx')

generate_report("timesheets")
