

# Begin by importing all modules that might be used in the script

#Importing os module
#This is useful for our current intentions regarding "HighRoller"'s data output
from itertools import repeat
import os
#shutil is another library module that carries out a similar function
import shutil
#importing date module
from datetime import date
#importing time module
from datetime import time, timedelta
import time
#importing random module
import random
#Excel handling module
import xlrd
#import keyboard stuff
import msvcrt

#These lines are dedicated to managing the Excel spreadsheet as a repository for the data
spread_path = os.path.join("T:\\", "Project Blue Rose", "Resources", "Big Blind", "Copy of Copy of High_Roller.xlsx")

#Opening workbook
workbook = xlrd.open_workbook(spread_path)
sheet = workbook.sheet_by_index(0)
weekly_sheet = workbook.sheet_by_index(1)
weekly_nrows = weekly_sheet.nrows
current_Date_Row_Index = 0
current_Date_Column_Index = 0
dormant_dir = ""
active_dir = ""





#Introduce automation to the system
print("""
Would the User like for me to automate the process for you?
If so please do not press any keys.
If however you would like to manually oversee the process, please press a key.

""")


#Give the user a chance to override the automated process
time.sleep(10)

#If the user wants to manually oversee the process, this function will ask the user to press a key
#If the system doesnt detect a key within 10 seconds of the above message, it will assume that the user wants to automate the process, setting the user's AFK status to True
if msvcrt.kbhit():
    print("You have chosen to manually oversee the process")
    user_AFK = False
    move_on = input("Hit enter to continue")
else:
    print("You have chosen to automate the process")
    user_AFK = True






#Begin defining functions to be called later



#This function fetches the current date and time and returns it as a string in format dd-mm-yyyy
def DateToday():

    #fetch the current date, to be used in the search for the files
    current_date = date.today()
    #we next need to format the date in a way such that it can be used in the search for the files
    current_date_formatted = current_date.strftime("%d-%m-%Y")
    print(current_date_formatted)
    return(current_date_formatted)

#We need to set the paths to the correct locations
#WireSnatch is a function that will search for the correct directories
#If the directory is not found, it will create it
#This will only do the main active and dormant mother folders
def WireSnatch():
    #not printing paths to console to avoid clutter when calling param repeatedly
    #param has been tested to work

    #set variables to the correct paths
    d_path = os.path.join("T:\\", "Project Blue Rose", "Dormant")
    a_path = os.path.join("T:\\", "Project Blue Rose", "Active")
    


    #check if the paths exist
    #if the dormant directory exists, set the dormant_dir variable to the path
    #if the dormant directory does not exist, create the directory and set the dormant_dir variable to the path
    #repeat for the active directory
    print("Checking for dormant directory")
    if os.path.exists(d_path):
        dormant_dir = d_path
        print("Dormant directory found")
    else:
        os.mkdir(d_path)
        dormant_dir = d_path
        print("Dormant directory was created")
    print("Checking for active directory")
    if os.path.exists(a_path):
        active_dir = a_path
        print("Active directory found")
    else:
        os.mkdir(a_path)
        active_dir = a_path
        print("Active directory was created")

    return(dormant_dir, active_dir)
#This will generate/validate daughter folders for the active directory and the dormant directory
#It is called to generate folder arrays within the active and dormant directories
def docudirectory_validate(dormant_dir, active_dir):
    dormant_type_paths = []
    active_type_paths = []
    document_types = ["Ramp-Maintenance-Register", "Weekly-Maintenance-Register", "360-Maintenance-Register", "Daily-Maintenance-Checklist", "Forklift-Maintenance-Register", "Site-Diary-and-Record"]
    for i in document_types:
        type_path_d = os.path.join(dormant_dir, i)
        type_path_a = os.path.join(active_dir, i)
        if os.path.exists(type_path_d):
            print("A folder for " + i + " already exists in the dormant directory")
        else:
            os.mkdir(type_path_d)
            print("A folder for " + i + " was created in the dormant directory")
        if os.path.exists(type_path_a):
            print("A folder for " + i + " already exists in the active directory")
        else:
            os.mkdir(type_path_a)
            print("A folder for " + i + " was created in the active directory")
        dormant_type_paths.append(type_path_d)
        active_type_paths.append(type_path_a)
    return(dormant_type_paths, active_type_paths)
        




#This function will decide which files need to be moved to the active directory
#THIS FUNCTION IS INCOMPLETE!!!!!
def FileDecider(dormant_dir, active_dir):
    #Decide which files need moving to the active directory
    #This will be done by creating a list of possible options
    #The list will be populated with the tags that will be used to search for the files
    #The system will then select the tags they want to move
    #The tags will be used to search for the files
    #The files will be moved to the active directory
    

    #First we need to get the current date in the format dd-mm-yyyy
    current_date = date.today()
    #Make a standard start date for the search
    start_date = date(2020, 1, 1)
    #Using the current date, we can create a list of possible options for the different documents
    #We need to be able to move retrospective documents as well as the current documents
    delta = current_date - start_date  # as timedelta
    days = [start_date + timedelta(days=i) for i in range(delta.days + 1)]
    formattedDays = days

#This function will search the Excel spreadsheet for filename Data and return the row index of the file
def cell_fetcher():
    cell_date = ""
    #First we will ask the user if they want to search for the current date or a manual date
    if user_AFK == True:
        cell_date = DateToday()
    elif user_AFK == False:
        #Ask the user for the date they want to search for
        date_needed = input("would you like to search for the current date or a manual date? (c/m)")
        if date_needed == "c" or " c" or "current" or "Current" or "CURRENT" or "Current Date" or "current date" or "Current Date" or "CURRENT DATE" or "C":
            cell_date = DateToday()
        elif date_needed == "m" or " m" or "manual" or "Manual" or "MANUAL" or "Manual Date" or "manual date" or "Manual Date" or "MANUAL DATE" or "M":
            cell_date = input("Enter the date you want to search for in the format dd-mm-yyyy: ")
        else:    
            print("Invalid input")
            cell_fetcher()
    #Now we need to search the spreadsheet for the date
    #We will use the xlrd library to search the spreadsheet
    
    for r in range(sheet.nrows):
        for c in range(sheet.ncols):
            if sheet.cell_value(r, c) == cell_date:
                current_Date_Row_Index = r 
                current_Date_Column_Index = c
                print("Row index: " + str(current_Date_Row_Index))
                print("Column index: " + str(current_Date_Column_Index))
    weekly_date_row_index = weekly_dates(cell_date,current_Date_Row_Index,current_Date_Column_Index)
    weekly_date_column_index = 0
    return(current_Date_Row_Index, current_Date_Column_Index, weekly_date_row_index, weekly_date_column_index)
    

#This funtion will create an array of all of the dates previous to the date generated by the cell_fetcher function
def previous_dates(current_Date_Row_Index, current_Date_Column_Index, weekly_date_row_index, weekly_date_column_index):
    dates_list = []
    weekly_dates_list = []
    for i in range(1, current_Date_Row_Index + 1):
        dates_list.append(sheet.cell_value(i, current_Date_Column_Index))
    for n in range(1, weekly_date_row_index + 1):
        weekly_dates_list.append(workbook.sheet_by_index(1).cell_value(n, weekly_date_column_index))

    return(dates_list, weekly_dates_list)



#This function will ask the user what type of document they want to move, and add this to an array to be called later
def document_types():
    #We will create a list of possible options
    #The user will then select the type of document they want to move
    #The type of document will be used to search for the files
    #The files will be moved to the active directory
    document_type_list = ["A", "B", "C", "D", "E", "F"]
    DocTypeTags = []
    for i in document_type_list:
        if user_AFK == False: 
            #ask user if they want to move the document type
            user_input = input("Would you like to move document " + i + "(Y/N)")
            if user_input == "Y" or "y" or "Yes" or "yes" or "YES":
                DocTypeTags.append(i)
                print ("Document type " + i + " added to the list")
            elif user_input == "N" or "n" or "No" or "no" or "NO":
                print("Document type " + i + " not added to the list")
            else:
                print("Invalid input")
                document_types()
        elif user_AFK == True:
            if i != "C":
                DocTypeTags.append(i)
                print ("Document type " + i + " added to the list")
    return(DocTypeTags)


#This function will attempt to find the weekly dates we need to use
def weekly_dates(cell_date, current_Date_Row_Index, current_Date_Column_Index):

    if sheet.cell_value(current_Date_Row_Index, 1) == "Monday":
        weekly_date_row_index = current_Date_Row_Index
        weekly_date_column_index = current_Date_Column_Index
    elif sheet.cell_value(current_Date_Row_Index, 1) != "Monday":
        for i in range(current_Date_Row_Index, 0, -1):
            if sheet.cell_value(i, 1) == "Monday":
                weekly_date_row_index = i
                weekly_date_column_index = current_Date_Column_Index
                weekly_start_date = sheet.cell_value(i, 3)
                print(weekly_start_date)
                break

    for r in range(weekly_nrows):
        if weekly_sheet.cell_value(r,0) == weekly_start_date:
            weekly_start_row_index = r
            weekly_start_column_index = 0

            print("Weekly row index: " + str(weekly_start_row_index))
            print("Weekly column index: " + str(weekly_start_column_index))
            return(weekly_start_row_index)





#This will generate the list of filenames that need moving
def file_Name_Generator(dates_list, weekly_dates_list, DocTypeTags):
    #This function will generate a list of file names based on the dates in the list
    #The list will be used to search for the files
    #The files will be moved to the active directory
    file_name_list_A= []
    file_name_list_B= []
    file_name_list_C= []
    file_name_list_D= []
    file_name_list_E= []
    file_name_list_F= []
    for n in DocTypeTags:
        for i in dates_list:
            if n == "A":
                file_name_list_A.append(n + " - " + i + ".docx")
            elif n == "C":
                file_name_list_C.append(n + " - " + i + ".docx")
            elif n == "D":
                file_name_list_D.append(n + " - " + i + ".docx")
            elif n == "E":
                file_name_list_E.append(n + " - " + i + ".docx")
            elif n == "F":
                file_name_list_F.append(n + " - " + i + ".docx")
        for i in weekly_dates_list:
            if n == "B":
                file_name_list_B.append(n + " - " + i + ".docx")
    return(file_name_list_A, file_name_list_B, file_name_list_C, file_name_list_D, file_name_list_E, file_name_list_F)


def file_shuttle(file_name_list_A, file_name_list_B, file_name_list_C, file_name_list_D, file_name_list_E, file_name_list_F, dormant_directory , active_directory):
    fail_Count = [0, 0, 0, 0, 0, 0]
    success_Count = [0, 0, 0, 0, 0, 0]
    existed_Count = [0, 0, 0, 0, 0, 0]
    file_name_array = [file_name_list_A, file_name_list_B, file_name_list_C, file_name_list_D, file_name_list_E, file_name_list_F]
    for n in range(len(file_name_array)):
        for i in file_name_array[n]:
            if os.path.exists( os.path.join(dormant_directory[n], i )) and os.path.exists( os.path.join(active_directory[n], i )) == False :
                shutil.copy(os.path.join(dormant_directory[n], i), active_directory[n])
                success_Count[n] += 1
            elif os.path.exists( os.path.join(active_directory[n], i )) == True:
                existed_Count[n] += 1
            else:
                fail_Count[n] += 1
    print("New files moved: "+ "A:" + str(success_Count[0]) + " B:" + str(success_Count[1]) + " C:" + str(success_Count[2]) + " D:" + str(success_Count[3]) + " E:" + str(success_Count[4]) + " F:" + str(success_Count[5]))
    print("Files that already existed in the active directory: " + "A:" + str(existed_Count[0]) + " B:" + str(existed_Count[1]) + " C:" + str(existed_Count[2]) + " D:" + str(existed_Count[3]) + " E:" + str(existed_Count[4]) + " F:" + str(existed_Count[5]))
    print("Files that couldn't be found: " + "A:" + str(fail_Count[0]) + " B:" + str(fail_Count[1]) + " C:" + str(fail_Count[2]) + " D:" + str(fail_Count[3]) + " E:" + str(fail_Count[4]) + " F:" + str(fail_Count[5]))
    


#Making row and column variables to be used in the previous_dates function global
current_Date_Row_Index, current_Date_Column_Index, weekly_date_row_index, weekly_date_column_index = cell_fetcher()

#using the WireSnatch function to set the dormant and active directories globally
dormant_dir, active_dir = WireSnatch()

print("Dormant directory variable: " + dormant_dir)
print("Active directory variable: " + active_dir)

#Assigning Globally for ease of use in callbacks
dates_list, weekly_dates_list = previous_dates(current_Date_Row_Index, current_Date_Column_Index, weekly_date_row_index, weekly_date_column_index)
print(weekly_dates_list)
#Final calling of the functions
A, B, C, D, E, F = file_Name_Generator(dates_list, weekly_dates_list, document_types())
dormant_type_paths,active_type_paths = docudirectory_validate(dormant_dir, active_dir)
file_shuttle(A, B, C, D, E, F, dormant_type_paths, active_type_paths)
