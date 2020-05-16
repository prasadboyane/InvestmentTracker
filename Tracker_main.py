import openpyxl
import os.path
import time
from datetime import datetime

#filepath = "/home/ubun/Desktop/stocksinfo/test101.xlsx"



def create_user():
    #Create User
    user=input('Please Enter New Username: ')
    password=input('Please Enter New Password: ')
    confirm_password=input('Please confiem Password again: ')
    
    if password == confirm_password:
        wb = openpyxl.Workbook()
        sheet = wb.active
        sheet.title = "Metadata"
        
        sheet['A1'] = user
        sheet['B1'] = password
        now = time.strftime("%x")
        sheet['C1'] = now
        
        sheet1 = wb.create_sheet("data")
        wb.save('Tracker_enc.xlsx')
        
    else:
        print()
        print('Passwords did not match. Please try once again !')
        print()
        create_user()
        

def login_user():
    #Create User
    user=input('Please Enter Username: ')
    password=input('Please Enter Password: ')
    wb = openpyxl.load_workbook('Tracker_enc.xlsx')
    ws = wb.active
    if user == ws['A1'].value  and password == ws['B1'].value:
        print('logged in !')
    else:
        print('password is incorrect ! Try again !')
        login_user()


def insert_entry():
    amount_entry = input('Please enter amount: ')
    date = input('Enter a date (YYYY-MM-DD)')
    try:
        year, month, day = map(int, date.split('-'))
        date_entry = datetime(year, month, day)
    except:
        print('Please valid enter date in correct format !')
        insert_entry()
    wb = openpyxl.load_workbook('Tracker_enc.xlsx')
    data_sheet = wb["data"]
    data_sheet['A1'] = 'id'
    data_sheet['B1'] = 'amount'
    data_sheet['C1'] = 'date'
    data_sheet['D1'] = 'assetname'
    data_sheet['E1'] = 'Maturity date'
    
    try:
        max_row_for_c = max((c.row for c in data_sheet['A'] if c.value is not None))  # To find max number of rows in 'A' columns
    except:
        max_row_for_c=0
        
    row = (max_row_for_c,amount_entry,date_entry)
    data_sheet.append(row)
    wb.save('Tracker_enc.xlsx') 
    
    go_next=input('Press "m" for Main Menu & "e" for exit ')
    if go_next=='m':
        show_menu()
    else:
        exit()

def update_entry():
    id_entry = input('Please enter entry id: ')
    
    
    wb = openpyxl.load_workbook('Tracker_enc.xlsx')
    data_sheet = wb["data"]
    
    row = (amount_entry,date_entry)
    data_sheet.append(row)
    wb.save('Tracker_enc.xlsx') 


def view_entry():
    wb = openpyxl.load_workbook('Tracker_enc.xlsx')
    ws = wb["data"]
    max_row_for_a = max((c.row for c in ws['A'] if c.value is not None))
    
    for row in ws.iter_rows(min_row=1, min_col=1, max_row=max_row_for_a, max_col=5):
        for cell in row:
            print(cell.value,'\t\t', end=" ")
        print()
    go_next=input('Press "m" for Main Menu & "e" for exit ')
    if go_next=='m':
        show_menu()
    else:
        exit()


def show_menu():
    print('******   1. Insert new Entry ******')
    print('******   2. Update Entry     ******')
    print('******   3. Delete Entry     ******')
    print('******   4. View Graph       ******')
    print('******   5. Exit             ******')
    choice = input('Please enter your choice: ')
    if choice == '1':
        print('choice 1')
        insert_entry()
    elif choice == '4':
        view_entry()
    elif choice == '5':
        exit()

##START PROGRAM HERE
if os.path.exists('Tracker_enc.xlsx') == True:
    login_user()
else:
    create_user()

show_menu()
