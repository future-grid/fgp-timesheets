#!/usr/local/bin/python3

# imports and globals for file locations
import sys, pandas, json, os, colorama, datetime, time
from colorama import init, Fore, Style
from openpyxl import Workbook

# init file variables
full_path = os.path.realpath(__file__)
path, filename = os.path.split(full_path)
folderPath = path + "/timesheet"
tempPath = folderPath + "/today.json" # saving a dict as json

blank_sheet = {
    "projects" : [],
}

# set both quickly
if "devmode" in sys.argv:
    sys.argv.append("verbose")
    sys.argv.append("dry-run")
    sys.argv.remove("devmode")

# init verboseness
if "verbose" in sys.argv:
    print(Style.BRIGHT + Fore.YELLOW + "Running verbosely" + Style.RESET_ALL)
    sys.argv.remove("verbose")
    def verbosePrint(msg):
        print("Verbose: " + msg)
else:
    def verbosePrint(msg):
        pass # do nothing

# init dryness
if "dry-run" in sys.argv:
    sys.argv.remove("dry-run")
    def saveTemp(contents):
        verbosePrint("saveTemp(): Would have saved " + str(contents))
else:
    def saveTemp(contents):
        verbosePrint("saveTemp(): Saving " + str(contents))
        raw = json.dumps(contents)
        f = open(tempPath,"w")
        f.write(raw)
        f.close()

# create folder if it doesn't already exist
if not os.path.exists(folderPath):
    print("Timesheet folder does not exist, attempting to make...")
    try:
        print("Made timesheet folder at " + folderPath)
        os.makedirs(folderPath)
    except OSError as err:
        print("Could not create folder!")
        print(err)
        exit()

def printUsage():
    colorama.init(autoreset=True)
    print(Style.BRIGHT + "Usage:")
    print("Adding projects to timesheet:")
    print(Fore.BLUE + "sheet " + Fore.GREEN + "add " + Style.RESET_ALL + "{project_code} {project_hours} {project_description}")
    print(Fore.BLUE + "sheet " + Fore.GREEN + "start " + Style.RESET_ALL + "{project_code} {project_hours} {project_description}")
    print(Fore.BLUE + "sheet " + Fore.GREEN + "finish")
    print("Viewing and managing timesheets:")
    print(Fore.BLUE + "sheet " + Fore.GREEN + "list")
    print(Fore.BLUE + "sheet " + Fore.GREEN + "rm " + Style.RESET_ALL + "{project_description OR last}")
    print(Fore.BLUE + "sheet " + Fore.GREEN + "export")
    print(Fore.BLUE + "sheet " + Fore.GREEN + "upload")
    print("See " + Fore.BLUE + "sheet " + Fore.GREEN + "count " + "help" + Style.RESET_ALL + " for counting info")

def args():
    return sys.argv[1:]


def readTemp():
    if not os.path.exists(tempPath):
        raise OSError("Today's sheet doesn't exist yet!")
    else:
        verbosePrint("readTemp(): The tempfile exists!")
        with open(tempPath, "r") as f:
            contents = json.load(f)
            verbosePrint("readTemp(): Here it is: " + str(contents))
            return(contents)


def sheetAdd(args):
    if not 2 <= len(args) <= 3:
        raise IndexError("Expected 2 or 3 arguments for sheet add, given " + str(len(args)))

    # set variables
    project_code = args[0]
    try:
        project_hours = float(args[1])
    except:
        raise IOError("Expected number for second argument, got " + str(type(args[1])))
    project_description = args[2] if len(args) == 3 else 'No description'
    project_exists = False
    project = {
        "project_code" : project_code,
        "project_hours" : project_hours,
        "project_description" : project_description,
    }

    # check if tempfile exists
    # read in tempfile
    try:
        sheet = readTemp()
        verbosePrint("sheetAdd(): Found tempfile.")
    except:
        sheet = blank_sheet
        verbosePrint("sheetAdd(): Couldn't find tempfile.")

    sheet["last_added"] = project

    for project in sheet["projects"]:
        if project_code == project["project_code"] and project_description == project["project_description"]:
            project_exists = True

    if project_exists:
        print(Fore.YELLOW + "Project " + Fore.BLUE + project_code + Fore.YELLOW + " already exists! Adding hours." + Style.RESET_ALL)
        project["project_hours"] += project_hours # add hours

    if not project_exists:
        sheet["projects"].append(project)
        
        print(Fore.GREEN + "Added " + Fore.BLUE + project_code + Fore.GREEN + "!" + Style.RESET_ALL)
    saveTemp(sheet)

    exit()

def sheetList(args):
    temp = readTemp()
    verbosePrint("sheetList(): Found tempfile")
    print(Style.BRIGHT + "Code (Hours): Description" + Style.RESET_ALL)
    try:
        working_project = temp["working_project"]
        verbosePrint("sheetList(): Loaded working project" + str(working_project))
        started_at = time.strptime(working_project["started_at"], '%H:%M:%S')
        current_time = time.strptime('%X', time.strftime('%X', time.localtime()))
        delta = time.strftime("%H", (current_time - started_at))
        print(working_project["project_code"] + " (" + str(delta) + "): " + working_project["project_description"])
    except Exception as err:
        raise err
        pass

    for project in temp["projects"]:
        print(project["project_code"] + " (" + str(project["project_hours"]) + "): " + project["project_description"])

def removeFromSheet(sheet, removal_project):
    verbosePrint("removeFromSheet(): Removing " + str(removal_project))
    project_found = False

    for project in sheet["projects"]:
        if removal_project["project_code"] == project["project_code"] and removal_project["project_description"] == project["project_description"]:
            project_found = True
            # if added multiple times, just remove the last set of hours
            if project["project_hours"] > removal_project["project_hours"]:
                project["project_hours"] -= removal_project["project_hours"]
                print(Fore.GREEN + "Removed " + Fore.BLUE + str(removal_project["project_hours"]) + Fore.GREEN + " hours from project." + Style.RESET_ALL)
            else:
                print(Fore.GREEN + "Removed " + Fore.BLUE + str(removal_project["project_code"]) + Fore.GREEN + " from sheet." + Style.RESET_ALL)
                sheet["projects"].remove(project)

    if not project_found:
        raise OSError("Project " + removal_project["project_code"] + " not found!")
    return sheet

def sheetRemove(args):
    sheet = readTemp()

    if len(args) is not 1 and len(args) is not 3:
        raise IndexError("Expected 3 arguments or 'last' for rm, given " + str(len(args)))

    if len(args) is 3:
        try:
            float(args[1])
        except:
            raise IOError("Expected a number for the second argument")

    if args[0] == 'last' and len(args) is 1:
        try:
            sheet = removeFromSheet(sheet, sheet["last_added"])
            del sheet["last_added"]
        except Exception as err:
            raise Exception(err)
            #raise IOError("There is no last added! You may have already removed it")
    else:
        sheet = removeFromSheet(sheet, {
            "project_code" : args[0],
            "project_hours" : float(args[1]),
            "project_description" : args[2]
        })

    saveTemp(sheet)

def sheetStart(args):
    # defensive
    if len(args) is not 0 and len(args) is not 2:
        raise IOError("Expected 2 or 0 args, given " + str(len(args)))

    try:
        sheet = readTemp()
    except:
        sheet = blank_sheet
        verbosePrint("sheetStart(): Could not find sheet, generated new.")

    try:
        working_project = sheet["working_project"]
        # working project already exists!
        raise IOError("Working project already in progress!")
    except IOError as err:
        raise IOError(err)
    except:
        verbosePrint("sheetStart(): Started new working project")
        current_time = time.localtime()

        working_project = {
            "started_at" : time.strftime("%X", current_time)
        }
        
        if len(args) is not 0:
            working_project["project_code"] = args[0]
            working_project["project_description"] = args[1]

        sheet["working_project"] = working_project

        print(Fore.GREEN + "Started project " + Fore.BLUE + working_project["project_code"] + Fore.GREEN + " at " + time.strftime("%X", current_time) + Style.RESET_ALL)
        verbosePrint("sheetStart(): Working project is " + str(working_project))
        saveTemp(sheet)

def main():
    try:
        if len(args()) < 1:
            raise IndexError("No arguments passed!")

        following = tuple(args()[1:]) # grab other args to pass to function

        command = args()[0] # main command is first in args
        if 'help' in following:
            verbosePrint("Found help in following")
            printUsage()
            exit()

        if command == 'help':
            verbosePrint("Found help in command")
            printUsage()
            exit()

        if command == 'add':
            sheetAdd(following)
        elif command == 'list':
            sheetList(following)
        elif command == 'rm':
            sheetRemove(following)
        elif command == 'export':
            sheetExport(following)
        elif command == 'count':
            sheetCount(following)
        elif command == 'start':
            sheetStart(following)
        elif command == 'finish' or command == 'end':
            sheetFinish(following)
        else:
            raise IOError("Not a valid argument!")
        exit() # exit on succeed
    except (Exception) as err:
        print(Fore.RED + "Didn't work! Error was this: " + Style.RESET_ALL)
        print(err)
        exit()

main()
