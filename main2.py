import time

import win32com.client as client
import datetime


# pyinstaller --hiddenimport win32timezone -F main.py
def main():
    # get account name
    account = get_outlook_account()

    choose = input("Alegeti un folder in mod manual? Daca da, introduceti 1, altfel apasati enter\n")
    if choose.isnumeric() and int(choose) == 1:
        inbox_name = input("Introduceti manual numele folderului:\n ")
    else:

        inbox_int = input(
            "NE VOM UITA DOAR IN INBOX.\nAlegeti limba contului. Pentru engleza apasati 1, pentru franceza apasati 2: \n")
        if not inbox_int.isnumeric():
            print("Nu ati introdus un numar, la revedere!")
            exit()
        elif int(inbox_int) > 2:
            print("Nu exista aceasta limba: " + str(inbox_int))
            exit()
        inbox_name = ""
        if int(inbox_int) == 2:
            inbox_name = "Boîte de réception"
        elif int(inbox_int) == 1:
            inbox_name = "Inbox"

    # get the default inbox folder
    try:
        inbox = account.Folders(inbox_name)
    except:
        print("Contul ales nu este in limba aleasa, la revedere")
        exit()

    # get the current date
    date_today = datetime.datetime.today()
    year = input("Introdu anul pentru care vrei sa verifici: ")
    if not year.isnumeric():
        print("Nu ati introdus un numar, la revedere!")
        exit()
    elif int(year) > date_today.year:
        print("Ati introdus un an mai mare decat anul curent, la revedere!")
        exit()

    dict = {}
    emails = []

    month1 = []
    month2 = []
    month3 = []
    month4 = []
    month5 = []
    month6 = []
    month7 = []
    month8 = []
    month9 = []
    month10 = []
    month11 = []
    month12 = []
    print("Inbox name: " + str(inbox_name))

    messages_fldr = []
    total_msg =0
    messages_fldr.append(inbox.Items)
    total_msg = len(inbox.Items)
    for folder in inbox.Folders:
        messages_fldr.append(folder.Items)
        total_msg = total_msg + len(folder.Items)


    print("Total emailuri: " + str(total_msg))
    for msgFolder in messages_fldr:
        message = msgFolder.GetLast()
        try:
            while message:
                try:
                    date_message = str(message.senton.date())
                    date_message = datetime.datetime.strptime(date_message, "%Y-%m-%d")
                    print(str(message.senton.date()))
                    if int(year) == date_message.year:
                        if date_message.month == 1:
                            month1.append(message)
                            print("+1")
                            message.close(0)
                        elif date_message.month == 2:
                            month2.append(message)
                            print("+1")
                            message.close(0)
                        elif date_message.month == 3:
                            month3.append(message)
                            print("+1")
                            message.close(0)
                        elif date_message.month == 4:
                            month4.append(message)
                            print("+1")
                            message.close(0)
                        elif date_message.month == 5:
                            month5.append(message)
                            print("+1")
                            message.close(0)
                        elif date_message.month == 6:
                            month6.append(message)
                            print("+1")
                            message.close(0)
                        elif date_message.month == 7:
                            month7.append(message)
                            print("+1")
                            message.close(0)
                        elif date_message.month == 8:
                            month8.append(message)
                            print("+1")
                            message.close(0)
                        elif date_message.month == 9:
                            month9.append(message)
                            print("+1")
                            message.close(0)
                        elif date_message.month == 10:
                            month10.append(message)
                            print("+1")
                            message.close(0)
                        elif date_message.month == 11:
                            month11.append(message)
                            print("+1")
                            message.close(0)
                        elif date_message.month == 12:
                            month12.append(message)
                            print("+1")
                            message.close(0)
                    message = msgFolder.GetPrevious()
                except Exception as ex:
                    print("ERR: " + str(ex))
                    print("N BEFORE: " + str(
                        len(month1) + len(month2) + len(month3) + len(month4) + len(month5) + len(month6) + len(
                            month7) + len(month8) + len(month9) + len(month10) + len(month11) + len(month12)))
                    message = msgFolder.GetPrevious()
            dict[1] = month1
            dict[2] = month2
            dict[3] = month3
            dict[4] = month4
            dict[5] = month5
            dict[6] = month6
            dict[7] = month7
            dict[8] = month8
            dict[9] = month9
            dict[10] = month10
            dict[11] = month11
            dict[12] = month12

            print("------------------------------------------------")
            print(str(msgFolder))
            for i in range(1, 13):
                print("Month " + str(i) + ": " + str(len(dict[i])))
            print("------------------------------------------------")
        except:
            print("Eroare, probabil nu exista folderul ales: " + str(inbox_name))

    time.sleep(200)


def get_outlook_account():
    """
    Get the outlook account object by account_name and return it
    :param account_name:
    :return:
    """
    # create outlook instance
    outlook = client.Dispatch('Outlook.Application')

    # get the namespace object
    namespace = outlook.GetNameSpace("MAPI")

    counter = 1
    accounts = []
    account_number = 1

    for acc in outlook.Session.Stores:
        print(str(counter) + ": " + str(acc))
        accounts.append(acc)
        counter += 1

    bool_acc = False

    while bool_acc is False:
        account_number = input("Alege un cont din lista, apasand numarul de ordine:\n ")
        if not account_number.isnumeric():
            account_number = input("Introdu un numar:\n ")
        elif int(account_number) > counter:
            account_number = input("Introdu un numar mai mic decat numarul total de conturi:\n ")
        else:
            bool_acc = True

    str_account_name = accounts[int(account_number) - 1]

    account = namespace.Folders(str(str_account_name))
    return account


if __name__ == '__main__':
    main()
