import gspread

# open service account
gc = gspread.service_account(filename='# insert your google service json file')
# open google spreadsheet
spreadsheet = gc.open("# insert your spreadsheet file name")

# to choose which sheet to open
worksheet = spreadsheet.sheet1

# Insert function


def insert():
    print('Enter a new data.')
    print("------------------------\n")
    user = input('Enter a new user\n')
    age = int(input('Enter a new age?\n'))
    city = input('Enter a new city\n')
    print('\n')

    worksheet.append_row([user, age, city])

# Delete function


def delete():
    print('Update data.')
    print("------------------------\n")

    print('Press 1 to delete row')
    print('Press 2 to delete column')
    option = int(input('Do you want to delete row or column?\n'))
    if option == 1:
        row = int(
            input('Which row do you want to remove? Enter the row number value.\n'))
        worksheet.delete_rows(row)
        print('\n')

    elif option == 2:
        col = int(
            input('Which column do you want to remove? Enter the column number value.\n '))
        worksheet.delete_columns(col)
        print('\n')

    else:
        print('Error! Please choose option 1 or 2 only')
        print('\n')

# Update function


def update():
    print('Update new data.')
    print("------------------------\n")
    cell = input('Enter a cell value. Exp: A3,C1\n')
    datavalue = input('Enter a new data value\n')
    print('\n')

    worksheet.update_acell(cell, datavalue)


# Share spreadsheet function
def share():
    print('Share spreadsheet.')
    print("------------------------\n")
    email = input('Enter an email you want to share to?\n')
    roles = input('Enter a user role(owner,writer,reader) for spreadsheet?\n')
    print('Do you want to have an email message\n')
    print('Press 1 for Yes')
    print('Press 2 for No')

    msg_option = int(input('\n'))
    if msg_option == 1:
        message = input('Enter an email message\n')
    elif msg_option == 2:
        message = input('')
    else:
        print('Error! Please choose option 1 or 2 only')

    spreadsheet.share(email, perm_type='user', role=roles,
                      notify=True, email_message=message)
    print('\n')


# Start App
while True:
    print('1) Insert data\n2) Delete data\n3) Update data\n4) Share spreadsheet\n')

    perform = int(input('Choose an option you want to perform\n'))
    if perform == 1:
        insert()
    elif perform == 2:
        delete()
    elif perform == 3:
        update()
    elif perform == 4:
        share()
    else:
        print('Error! Please choose option from 1-4 only.')
