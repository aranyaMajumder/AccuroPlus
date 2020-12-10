# AccuroPlus V-1.0
import openpyxl as op
import os
import random as r
import platform
import time as t
from datetime import datetime
wb = op.load_workbook("UserData.xlsx")
sheet = wb["Sheet1"]

def userSystem():
    global clear
    osx = platform.system()
    if osx == "Windows":
        print("Windows " + platform.release() + " System Detected")
        print()
        clear = 'cls'
    elif osx == "Linux":
        print("Linux System Detected")
        print()
        clear = 'clear'
    else:
        print("Sorry your os does not support this version. Accuroplus supports only Linux or Windows currently.")
        input("Press enter to continue....")
        exit()

userSystem()

def First_Screen():
    print("""Choose a function from below
    * Test yourself - 1
    * Check accuracy - 2
    * Check words per second speed - 3

    """)
    global user_choice
    user_choice = int(input("Choose: "))
    choiceEval(user_choice)

def avg_time():
    print("This will show you the average time you take to write a word")
    u = 0
    total = 0
    while True:
        u += 1
        acc = sheet["B" + str(u)]
        if acc.value == None:
            break
        else:
            total += acc.value
    print("Time taken is " + str(total / (u - 1)) + " seconds")
    print()
    input("Press enter to continue.....")
    os.system(clear)
    First_Screen()

def accuracy_che():
    print("This function will show your average accuracy")
    print()
    print()
    u = 0
    total = 0
    while True:
        u += 1
        acc = sheet["A" + str(u)]
        if acc.value == None:
            break
        else:
            total += acc.value
    print("Accuracy is " + str(total/(u-1)) + "%")
    print()
    input("Press enter to continue.....")
    os.system(clear)
    First_Screen()

def test():
    global accuracy_save
    while True:
        print("Hello There!")
        if dataExist.value == 0:
            print("Note:- Data not found. Creating data from now....")
        elif dataExist.value == 1:
            print("Data found! Editing Existing data...")
        print("")
        limit_of_word = r.randint(3, 10)
        word = ""
        coun = 0
        while coun<limit_of_word:
            letters = chr(r.randint(97, 122))
            word += letters
            coun += 1
        print(word)
        now = datetime.now()
        current = now.strftime("%M")
        current_s = now.strftime("%S")
        cure = int(current) * 60 + int(current_s)
        user_inp = input("Try it out: ")
        dataExist.value = 1
        now1 = datetime.now()
        current1 = now1.strftime("%M")
        current_s1 = now1.strftime("%S")
        cure1 = int(current1) * 60 + int(current_s1)
        if user_inp == word:
            time_taken = cure1-cure
            print("Time taken: " + str(time_taken) + " seconds")
            print("Bravo! Want to go to the next word")
            choice = input("(y/n) ")
            acuracy = 100
            counter = 1
            n = 1
            while True:
                accuracy_save = sheet["A" + str(n)]
                if accuracy_save.value == None:
                    break
                else:
                    n += 1
            accuracy_save.value = acuracy
            wb.save('UserData.xlsx')
            while True:
                time_lists = sheet["B" + str(counter)]
                if time_lists.value == None:
                    time_lists.value = time_taken
                    break
                else:
                    pass
                counter += 1
            if choice == 'y':
                os.system(clear)
            else:
                while True:
                    accuracy_save = sheet["a" + str(n)]
                    if accuracy_save.value == None:
                        break
                    else:
                        n += 1
                accuracy_save.value = acuracy
                wb.save('UserData.xlsx')
                os.system(clear)
                First_Screen()
        else:
            print("Oops! That was wrong!")
            wrong_letters = []
            n = 0
            wrong_num = 0
            while n <= (len(word)-1):
                if word[n] == user_inp[n]:
                    pass
                else:
                    wrong_letters.append(word[n])
                    wrong_num += 1
                    print("You had misspelled " + word[n])
                n += 1
            acuracy = ((len(word)-wrong_num)/len(word))*100
        n = 1
        while True:
            accuracy_save = sheet["a" + str(n)]
            if accuracy_save.value == None:
                break
            else:
                n += 1
        accuracy_save.value = acuracy

        wb.save("UserData.xlsx")
        input("Press enter for the next word....")
        os.system(clear)
def choiceEval(choice): # choice is parameter, pass user_choice at the end through it
    os.system(clear)
    global dataExist
    dataExist = sheet["D7"]
    if dataExist.value == 1: #Data exists
        if choice == 1:
            test()
        elif choice == 2:
            accuracy_che()
        elif choice == 3:
            avg_time()
    else:
        if 1<choice<6:
            print("Sorry, we do not possess enough data to perform this action, first attempt some tests. Going back")
            x = input("Press enter to continue...")
            os.system(clear)
            First_Screen()
        elif choice == 1:
            test()
        else:
            print("Invalid choice. Going back")
            x = input("Press enter to continue....")
            os.system(clear)
            First_Screen()

print("""Choose a function from below
* Test yourself - 1
* Check accuracy - 2
* Check words per second speed - 3


""")

first_choice = int(input("Choose: "))
choiceEval(first_choice)
wb.save("UserData.xlsx")