import openpyxl
import os.path
import time
from datetime import datetime
import matplotlib.pyplot as plt

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
    id_entry = input('Please enter id to update: ')
    id_entry_int=int(id_entry)+1
    id_entry_1=str(id_entry_int)
    
    wb = openpyxl.load_workbook('Tracker_enc.xlsx')
    data_sheet = wb["data"]
    i_a = 'A{}'.format(id_entry_1)
    i_b = 'B{}'.format(id_entry_1)
    i_c = 'C{}'.format(id_entry_1)
    i_d = 'D{}'.format(id_entry_1)
    i_e = 'E{}'.format(id_entry_1)
    
    msg="Here's your record: ",data_sheet[i_a].value,data_sheet[i_b].value,data_sheet[i_c].value,data_sheet[i_d].value,data_sheet[i_e].value
    print(msg)
    print('which field you want to edit?')
    print('1. amount')
    print('2. date')
    print('3. assetname')
    print('4. maturity date')
    print('5. Back')
    choice=input('Enter choice number: ')
    
    if choice=='1':
        updated_val=input('Enter updated value for amount: ')
        data_sheet[i_b] = updated_val
    elif choice =='2':
        updated_val=input('Enter updated value for date(YYYY-MM-DD): ')
        try:
            year, month, day = map(int, updated_val.split('-'))
            updated_val_1 = datetime(year, month, day)
        except:
            print('Please valid enter date in correct format !')
            update_entry()
        data_sheet[i_c] = updated_val_1
    elif choice =='3':
        updated_val=input('Enter updated value for assetname: ')
        data_sheet[i_d] = updated_val
    elif choice =='4':
        updated_val=input('Enter updated value for Maturity date(YYYY-MM-DD): ')
        try:
            year, month, day = map(int, updated_val.split('-'))
            updated_val_1 = datetime(year, month, day)
        except:
            print('Please valid enter date in correct format !')
            update_entry()
        data_sheet[i_c] = updated_val_1
    elif choice =='5':
        show_menu()
        
    wb.save('Tracker_enc.xlsx') 
    
    go_next=input('Press "m" for Main Menu & "e" for exit ')
    if go_next=='m':
        show_menu()
    else:
        exit()


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


def delete_entry():
    wb = openpyxl.load_workbook('Tracker_enc.xlsx')
    ws = wb["data"]
    
    id_entry = input('Please enter id to delete: ')
    ws.delete_rows(int(id_entry)+1)
    wb.save('Tracker_enc.xlsx')
    
    go_next=input('Press "m" for Main Menu & "e" for exit ')
    if go_next=='m':
        show_menu()
    else:
        exit()
    

def view_graph():
    wb = openpyxl.load_workbook('Tracker_enc.xlsx')
    ws = wb["data"]
    
    date_arr=[]
    colC = ws['C']
    for cell in colC:
        if cell.value != 'date':
            date_arr.append(cell.value)
    print()    
    
    date_arr.sort()
    min_date=min(date_arr)
    max_date=max(date_arr)
    
    amount_arr=[]
    colB = ws['B']
    for cell in colB:
        if cell.value != 'amount':
            amount_arr.append(cell.value)
    print()    
    
    amount_arr.sort()
    min_am=min(amount_arr)
    max_am=max(amount_arr)
    
    
    plt.plot(date_arr, amount_arr,'go--', linewidth=1, markersize=2)
    plt.axis([min_date, max_date, min_am, max_am])
    plt.ylabel('Amount')
    plt.show()
    
    go_next=input('Press "m" for Main Menu & "e" for exit ')
    if go_next=='m':
        show_menu()
    else:
        exit()
    

def show_menu():
    print('******   1. Insert new Entry ******')
    print('******   2. Update Entry     ******')
    print('******   3. Delete Entry     ******')
    print('******   4. View Entries     ******')
    print('******   5. View Graph       ******')
    print('******   6. Exit             ******')
    choice = input('Please enter your choice: ')
    if choice == '1':
        print('choice 1')
        insert_entry()
    elif choice == '2':
        update_entry()
    elif choice == '3':
        delete_entry()
    elif choice == '4':
        view_entry()
    elif choice == '5':
        view_graph()
    elif choice == '6':
        exit()

##START PROGRAM HERE
if os.path.exists('Tracker_enc.xlsx') == True:
    login_user()
else:
    create_user()

show_menu()
