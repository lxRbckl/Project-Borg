# Borg by Highlander #

import matplotlib.pylab as plt
import random, time, subprocess, glob, datetime, xlrd, xlwt, os
from xlrd import open_workbook
from xlutils.copy import copy

# Borg : Functions #

def Borg_Directory():
    ''' finds excel models '''

    global Directory_List

    Directory_List = [Directory_File[:-4] for Directory_File in glob.glob('*.xls')]

def Borg_Menu():
    ''' display commands '''

    Menu_List = ['1', '2', '3']
    while True:

        Menu_Input = input('    Borg Menu\n\n1.\tNew Model\n2.\tRead Model\n3.\tDelete Model\n\nInput : ')
        print()

        if (Menu_Input in Menu_List):

            return Menu_Input

        else:

            print('Input was invalid.\n')
            time.sleep(1.2)

def Borg_New():
    ''' develops a new model '''

    while True:

        New_ID = ''.join(str(random.randint(0, 9)) for i in range(4))
        for Directory_Element in Directory_List:

            if (Directory_Element[-4:] == New_ID):

                continue

        break

    while True:

        New_Input = input('<Press Enter to Exit>\nInput Model Description : ')
        print()

        if (len(New_Input) == 0):

            return

        New_Check = False
        for Directory_Element in Directory_List:

            if (Directory_Element[:-4] == New_Input):

                New_Check = True

        if (New_Check == True):

            print('Input description already exists.\n')
            continue

        elif (New_Check == False):

            print('Model {}{} was created.\n'.format(New_Input, New_ID))
            Directory_List.append('{}{}'.format(New_Input, New_ID))

            Borg_Create('{}{}'.format(New_Input, New_ID))
            break

def Borg_Open():
    ''' reads an existing model '''

    while True:

        Open_Input = input('<Press Enter to Exit>\nInput Model ID : ')
        print()

        if (len(Open_Input) == 0):

            return

        for Directory_Element in Directory_List:

            if (Directory_Element[-4:] == Open_Input):

                Borg_readModel('{}'.format(Directory_Element)) # test # change
                break

        print('Model ID does not exist.\n')

def Borg_Delete():
    ''' deletes an existing model '''

    while True:

        Delete_Input = input('<Press Enter to Exit>\nInput Model ID : ')
        print()

        if (len(Delete_Input) == 0):

            return

        for Directory_Element in Directory_List:

            if (Directory_Element[-4:] == Delete_Input):

                print('Model {} was deleted.\n'.format(Directory_Element))

                os.remove('{}.xls'.format(Directory_Element))
                Directory_List.remove(Directory_Element)
                return

        print('Model ID does not exist.\n')

def Borg_Create(Create_Parameter):
    ''' develop excel model '''

    with open('Borg_Settings.txt', 'r') as Create_File:

        [Create_List] = [Create_Line.strip().split(',') for Create_Line in Create_File]

    Create_Workbook = xlwt.Workbook()
    Create_Worksheet = Create_Workbook.add_sheet(Create_Parameter[-4:])

    Create_Style = xlwt.easyxf('font: bold True; align: horiz centre; borders: left thin, bottom thick, right thin, top thin; font: height 250')
    [Create_Worksheet.write(1, Create_Index + 1, Create_List[Create_Index], Create_Style) for Create_Index in range(len(Create_List))]

    Create_Style = xlwt.easyxf('font: bold True; align: horiz centre; borders: bottom thin, left thin, top thin, right thick; font: height 250')
    [Create_Worksheet.write(Create_Index + 1, 31, Create_List[Create_Index - 1], Create_Style) for Create_Index in range(1, len(Create_List) + 1)]

    Create_Workbook.save('{}.xls'.format(Create_Parameter))

    try:

        os.startfile('{}.xls'.format(Create_Parameter))

    except:

        subprocess.call(['open', '{}.xls'.format(Create_Parameter)])

    input('<Press Enter to Continue>\nPlease save and close model.\n')
    time.sleep(0.6)

    Borg_readModel(Create_Parameter)

def Borg_readModel(Read_Parameter):
    ''' conceptualize excel model '''

    with open('Borg_Settings.txt', 'r') as Read_File:

        [Read_List] = [Read_Line.strip().split(',') for Read_Line in Read_File]

    Read_Workbook = xlrd.open_workbook('{}.xls'.format(Read_Parameter))
    Read_Worksheet = Read_Workbook.sheet_by_index(0)

    Read_List_A = []
    for Read_Column in range(1, 31):

        Read_List_B = []
        for Read_Row in range(2, 103):

            try:

                if (Read_Worksheet.cell_value(Read_Row, Read_Column) == ''):

                    continue

                else:

                    Read_List_B.append(Read_Worksheet.cell_value(Read_Row, Read_Column))

            except:

                continue

        Read_List_A.append(Read_List_B)

    Read_List_C = [[((Read_Index_A + Read_Index_B) % 38) for Read_Index_A in range(-8, 3)] for Read_Index_B in range(38)]

    Read_List_D, Read_List_E, Read_List_F = [], [], []
    for Read_Index_A in range(len(Read_List_A)):

        Read_List_G, Read_List_H, Read_List_I = [], [], []
        for Read_Element_A in Read_List_C:

            Read_Count_A = 0
            for Read_Element_B in Read_Element_A:

                Read_Count_A += Read_List_A[Read_Index_A].count(Read_Element_B)

            if (Read_Count_A != 0):

                Read_List_G.append(Read_Element_A[9])
                Read_List_H.append(Read_Count_A)
                Read_List_I.append(round((Read_Count_A / len(Read_List_A[Read_Index_A]) * 100), 1))

        Read_List_D.append(Read_List_G)
        Read_List_E.append(Read_List_H)
        Read_List_F.append(Read_List_I)

    Read_Workbook = open_workbook('{}.xls'.format(Read_Parameter), formatting_info = True)
    Read_Workbook = copy(Read_Workbook)

    Read_Worksheet = Read_Workbook.get_sheet(0)
    Read_Worksheet.name = Read_Parameter[-4:]

    with open('Borg_WS.txt', 'r') as Read_File:

        [File_List] = [Read_Line.strip().split(',') for Read_Line in Read_File]

    Read_Count_A = 2
    for Read_Index_A in range(len(Read_List_E)):

        Read_Dictionary = {}
        for Read_Index_B in range(len(Read_List_E[Read_Index_A])):

            Read_Dictionary[Read_List_D[Read_Index_A][Read_Index_B]] = Read_List_E[Read_Index_A][Read_Index_B]

        Read_List_G = sorted(Read_Dictionary.items(), key=lambda x: x[1], reverse=True)

        Read_Count_B = 32
        for Read_Element in Read_List_G:

            Read_Value = round((Read_Element[1] / len(Read_List_A[Read_Count_A - 2]) * 100), 1)

            if ((Read_Value <= 100) and (Read_Value >= 75)):

                Read_Style = xlwt.easyxf('font: height 250; align: horiz centre; pattern: pattern solid, fore_colour red')

            elif ((Read_Value <= 74) and (Read_Value >= 50)):

                Read_Style = xlwt.easyxf('font: height 250; align: horiz centre; pattern: pattern solid, fore_colour light_orange')

            elif ((Read_Value <= 49) and (Read_Value >= 25)):

                Read_Style = xlwt.easyxf('font: height 250; align: horiz centre; pattern: pattern solid, fore_colour yellow')

            elif ((Read_Value <= 24) and (Read_Value >= 0)):

                Read_Style = xlwt.easyxf('font: height 250; align: horiz centre; pattern: pattern solid, fore_colour lime')

            Read_Worksheet.write(Read_Count_A, Read_Count_B, '{}/{} {}/{}'.format(Read_Element[0], File_List[Read_Element[0]], Read_Element[1], len(Read_List_A[Read_Count_A - 2])), Read_Style)
            Read_Count_B += 1

        Read_Count_A += 1

    Read_Workbook.save('{}.xls'.format(Read_Parameter))

    try:

        os.startfile('{}.xls'.format(Read_Parameter))

    except:

        subprocess.call(['open', '{}.xls'.format(Read_Parameter)])

    input('<Press Enter to Continue>\nProceed to Graph Model.\n'), time.sleep(0.6)

    Borg_readGraph(Read_List_D, Read_List_F)

def Borg_readGraph(Read_Parameter_A, Read_Parameter_B):
    ''' conceptualize graph of model '''

    with open('Borg_Settings.txt', 'r') as Read_File:

        [Read_List_A] = [Read_Line.strip().split(',') for Read_Line in Read_File]
        Read_List_A.insert(0, '')

    Read_List_A.insert(len(Read_List_A), '')

    while True:

        while True:

            Read_Input_A = input('Amount of Revolutions [1-3]\nInput : ')
            print()

            if ((Read_Input_A == '1') or (Read_Input_A == '2') or (Read_Input_A == '3')):

                break

            else:

                print('Invalid amount.\n')

        while True:

            try:

                Read_Input_B = float(input('Graph Points Greater [0-100]\nInput : '))
                print()

                if ((Read_Input_B <= 100) and (Read_Input_B >= 0)):

                    break

                else:

                    print('\nInvalid input.\n')

            except:

                print('\nInvalid input.\n')

        Read_List_B = []
        for Read_Index in range(int(Read_Input_A)):

            Read_List_B.insert(len(Read_List_B), '')
            [Read_List_B.append(Read_Element) for Read_Element in range(38)]

        Read_List_B.insert(len(Read_List_B), '')

        plt.tick_params(labelsize=7)
        plt.sub

        Read_Count_C, Read_Count_D = 0, 0
        for Read_Count_A in range(int(Read_Input_A)):

            Read_Count_A, Read_Count_B = 0, 0
            for Read_Element_A in Read_Parameter_B:

                if (len(Read_Element_A) == 0):

                    Read_Count_A += 1
                    Read_Count_B = 0
                    continue

                for Read_Element_B in Read_Element_A:

                    if (Read_Element_B <= Read_Input_B):

                        Read_Count_B += 1
                        continue

                    elif ((Read_Element_B <= 100) and (Read_Element_B >= 75)):

                        plt.scatter(Read_Count_A + 1 + Read_Count_C, Read_Parameter_A[Read_Count_A][Read_Count_B] + 1 + Read_Count_D, s = 15, color = 'r')

                    elif ((Read_Element_B <= 74) and (Read_Element_B >= 50)):

                        plt.scatter(Read_Count_A + 1 + Read_Count_C, Read_Parameter_A[Read_Count_A][Read_Count_B] + 1 + Read_Count_D, s = 15,  color = 'darkorange')

                    elif ((Read_Element_B <= 49) and (Read_Element_B >= 25)):

                        plt.scatter(Read_Count_A + 1 + Read_Count_C, Read_Parameter_A[Read_Count_A][Read_Count_B] + 1 + Read_Count_D, s = 15,  color = 'y')

                    elif ((Read_Element_B <= 24) and (Read_Element_B >= 0)):

                        plt.scatter(Read_Count_A + 1 + Read_Count_C, Read_Parameter_A[Read_Count_A][Read_Count_B] + 1 + Read_Count_D, s = 15, color = 'g')

                    Read_Count_B += 1

                Read_Count_A += 1
                Read_Count_B = 0

            Read_Count_C += 0.2
            Read_Count_D += 39

        plt.xticks(range(len(Read_List_A)), Read_List_A)
        plt.yticks(range(len(Read_List_B)), Read_List_B)

        plt.grid(linewidth = 0.3)
        plt.show()

        Read_Input = input('<Press Enter to Exit>\n1. Adjust Graph\nInput : ')
        print()

        if (len(Read_Input) == 0):

            break

# Borg : Main #

Borg_Directory()
while True:

    Borg_Input = Borg_Menu()
    time.sleep(0.6)

    if (Borg_Input == '1'):

        Borg_New()

    elif (Borg_Input == '2'):

        Borg_Open()

    elif (Borg_Input == '3'):

        Borg_Delete()

    time.sleep(0.6)
