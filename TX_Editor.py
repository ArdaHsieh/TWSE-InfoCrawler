# -*- coding: utf-8 -*-
"""
Created on Tue Nov 28 2017

@author: I-Ta Hsieh(Arda)
"""
import os, datetime

data = []
filename = "TX Due Date.txt"


def Menu():
    os.system("cls")
    print("TX Due Date Management System")
    print("------------------------------------------")
    print("1. Input New date")
    print("2. Display")
    print("3. Delete")
    print("4. Sort")
    print("0. Exit")
    print("------------------------------------------")
    
    
def ReadData():
    with open (filename, 'r', encoding = 'UTF-8-sig') as f:
        filedata = f.read()
        if (filedata != ""):
            data = filedata.split()
            return data
        else:
            return list()    
        

def sort_date():
    global data
    for i in range(len(data)):
        for j in range(i, len(data)):
            if datetime.date(int(data[i][0:4]), int(data[i][5:7]), int(data[i][8:10])) > datetime.date(int(data[j][0:4]), int(data[j][5:7]), int(data[j][8:10])):
            #if int(delta) > 0:
                data[i], data[j]  = data[j], data[i]
    
    with open(filename, 'w', encoding = 'UTF-8-sig') as f:
        for date in data:
            f.write(date + "\n")
    f.close()
                    
def disp_data():
    print("TX Due Date")
    print("==============")
    for date in data:
        print(date + "\n") 
        
    
def input_data():
    global data
    while True:
        disp_data()
        Due_date = input("New date (Enter => Stop):")
        if Due_date == "":
            break
        if Due_date in data:
            print("%s already exist."% (Due_date))
            continue
        data.append(Due_date)
        
        with open(filename, 'w', encoding = 'UTF-8-sig') as f:
            for date in data:
                f.write(date + "\n")
        print("{} is alreaady stored.".format(Due_date))
        disp_data()
        print("")
        print("I. Input another due date.")
        print("S. Sort date.")
        print("M. Back to Menu.")
        choice = input("Your choice:")
        if (choice == "I" or choice == "i"):
            continue
        elif (choice == "S" or choice == "s"):
            sort_date()
        else:
            break

    
def del_data():
    global data
    while True:
        disp_data()
        Due_date = input("Date you want to delete. (Enter => Stop):")
        if Due_date == "":
            break
        if not Due_date in data:
            print("%s does not exist." % (Due_date))
            continue
        for i in range (len(data)):
            if (data[i] == Due_date):
                index = i      
        yn = input("You want to delete {}? (Y/N)".format(Due_date)) # Confirm
        if (yn == "Y" or yn == "y"):
            del data[index]
            with open(filename, 'w', encoding = 'UTF-8-sig') as f:
                for date in data:
                    f.write(date + "\n")
                print("{} is alreaady edited. (Press any key to menu)".format(Due_date))
                disp_data()
        else:
            continue
        print("")
        print("D. Delete another due date.")
        print("M. Back to Menu.")
        choice = input("Your choice:")
        if (choice == "D" or choice == "d"):
            continue
        else:
            break    


def Main():
    global data
    
    print("      台指期結算日期資料編輯系統      ")
    print("--------------------------------------")
    while True:
        if not (os.path.exists(filename)):
            open (filename, 'w', encoding = 'UTF-8-sig')
            Menu()
            choice = input("Your Choice:")
            if (choice == "1"):
                input_data()
            elif (choice == "2"):
                disp_data()
                input("Press any key to continue")
            elif (choice == "3"):
                del_data()
            else:
                break
        elif (os.path.exists(filename)):
            data = ReadData()
            Menu()
            choice = input("Your Choice:")
            if (choice == "1"):
                input_data()
            elif (choice == "2"):
                disp_data()
                input("Press any key to continue")
            elif (choice == "3"):
                del_data()
            elif (choice == "4"):
                sort_date() 
            else:
                break
    print("\nEdit finish\n")


Main()
    
    