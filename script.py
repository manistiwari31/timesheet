import csv
import openpyxl
from openpyxl import load_workbook
from datetime import date
from datetime import datetime
import os

def to_excel(file_path, csv_path):

    
    with open(csv_path, 'r', newline='') as file:
        reader = csv.reader(file)
        rows = list(reader)

    workbook = openpyxl.Workbook()
    sheet = workbook.active

    # Write data to the Excel sheet
    for row in rows:
        sheet.append(row)

    # Save the workbook to the specified file path
    workbook.save(file_path)



def createcsv(file_path, row_data):
    file_exists = os.path.isfile(file_path)

    # Open the CSV file in append mode or create a new file with header
    with open(file_path, 'a', newline='') as file:
        writer = csv.writer(file)

        if not file_exists:
            writer.writerow(['Date', 'Start', 'End', 'Total', 'Project'])

        # Write the new row to the CSV file
        writer.writerow(row_data)

def appendcsv(file_path, row_data):
    with open(file_path, 'a', newline='') as file:
        writer = csv.writer(file)
        writer.writerow(row_data)

def time_diff(target_time_str):

    current_time = datetime.now()
    target_time = datetime.strptime(target_time_str, '%H:%M')
    time_difference = current_time -target_time
    hours, remainder = divmod(time_difference.seconds, 3600)
    minutes, _ = divmod(remainder, 60)

    return str(hours)+":"+ str(minutes)

def end(sheet, file_path, project):
    # Read the existing CSV file
    with open(file_path, 'r', newline='') as file:
        reader = csv.reader(file)
        rows = list(reader)

    # Retrieve the last line
    last_row = rows[-1]
    if last_row[-1] == project:
        print("Already Written")
        return
    exit_time = datetime.now().strftime("%H:%M")
    total_time = time_diff(last_row[1])
    temp = [exit_time,total_time,project]
    last_row += temp

    # Write the updated data back to the CSV file
    with open(file_path, 'w', newline='') as file:
        writer = csv.writer(file)
        writer.writerows(rows)
    to_excel(sheet,file_path)
 
        
def cck(sheet, file_path):
    # Read the existing CSV file
    file_exists = os.path.isfile(file_path)
    if not file_exists:
        return
    with open(file_path, 'r', newline='') as file:
        reader = csv.reader(file)
        rows = list(reader)
    # Retrieve the last line
    last_row = rows[-1]
    last = last_row[0]
    today = str(date.today())
    Last = int(last[-2:])
    Today = int(today[-2:])
    if Last==Today:
        print("Already made an entry for",today)
        return False
    Repeat = Today-Last
    emptydata = []
    for i in range(Repeat):
        data = [str(today[:-2]+str(Last+i)), 0,0 ,0,"Not working"]
        emptydata.append(data)
        with open(file_path, 'r', newline='') as file:
            reader = csv.reader(file)
            rows = list(reader)

        
        rows.append(data)
            
        with open(file_path, 'w', newline='') as file:
            writer = csv.writer(file)
            writer.writerows(rows)  
        to_excel(sheet,file_path)
    return True  
        
    
    # Write the updated data back to the CSV file
    

def start(file):
      
    today = date.today()
    now = datetime.now()

    current_time = now.strftime("%H:%M")
    print("Today's date:", today)

    print("Current Time =", current_time)
    data = [str(today), current_time]
    createcsv(file, data)
  


if __name__ == "__main__":
    project = "CWD Soundbox"
    file = "temp.csv"
    sheet = "Jan.xlsx"
    i = input("Enter i for entry or q for exit: ")
    print()
    if i == "I" or i == "i":
        if cck(sheet, file):
            start(file)
    elif i =="q" or i == "Q":
        end(sheet,file,project)
    else:
        print("wrong entryy")

    print("Have a nice Day.")
