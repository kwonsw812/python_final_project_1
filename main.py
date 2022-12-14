import os
import re
import openpyxl
import pandas as pd

file_path = input("파일 경로를 복사 후 붙여넣으시오. : ")

file_names = os.listdir(file_path)
file_names.sort()
print(file_names)

sel = input("바꿀 방식을 선택하시오.\n1.입력한 숫자부터 + 1\n"
            "2.문자열 + 입력한 숫자부터 + 1\n3.파일에서 가져오기\n")

if sel == "1" :
    count = input("시작할 숫자를 입력하세요 : ")
    i = int(count)
    for file_name in file_names:
        print(file_path)
        old = file_path+"\\"+file_name
        dst = str(i) + '.jpg'
        new = file_path+"\\"+dst
        print(dst)
        os.rename(old, new)
        i += 1

elif sel == "2" :
    change = input("문자열을 입력하세요 :")
    count = input("시작할 숫자를 입력하세요 : ")
    i = int(count)
    for file_name in file_names:
        print(file_path)
        old = file_path+"\\"+file_name
        dst = change + str(i) + '.jpg'
        new = file_path+"\\"+dst
        print(dst)
        os.rename(old, new)
        i += 1

elif sel == "3" :
    file_location = input("사용할 파일의 경로를 복사 후 붙여넣으시오. : ")
    files = os.listdir(file_location)

    for file in files :
        if ".xlsx" in file :
            path = file_location + '\\' + file
            df = pd.read_excel(path)
            sel_cell = input("사용할 행 또는 열을 입력하시오. : ")
            num = re.sub(r'[^0-9]', '', sel_cell)

            if "행" in sel_cell :
                name_list = df.iloc[int(num)]
                if len(name_list) <= len(file_names) :
                    i = 0
                    for file_name in file_names:
                        if i <= len(name_list) :
                            print(file_path)
                            old = file_path + "\\" + file_name
                            dst = name_list[i] + '.jpg'
                            new = file_path + "\\" + dst
                            print(dst)
                            os.rename(old, new)
                            i += 1

                if len(name_list) > len(file_names) :
                    i = 0
                    for file_name in file_names:
                        if i <= len(file_names):
                            print(file_path)
                            old = file_path + "\\" + file_name
                            dst = name_list[i] + '.jpg'
                            new = file_path + "\\" + dst
                            print(dst)
                            os.rename(old, new)
                            i += 1

            if "열" in sel_cell :
                name_list = df[int(num)]

                if len(name_list) <= len(file_names) :
                    i = 0
                    for file_name in file_names:
                        if i <= len(name_list):
                            print(file_path)
                            old = file_path + "\\" + file_name
                            dst = name_list[i] + '.jpg'
                            new = file_path + "\\" + dst
                            print(dst)
                            os.rename(old, new)
                            i += 1

                if len(name_list) > len(file_names):
                    i = 0
                    for file_name in file_names:
                        if i <= len(file_names):
                            print(file_path)
                            old = file_path + "\\" + file_name
                            dst = name_list[i] + '.jpg'
                            new = file_path + "\\" + dst
                            print(dst)
                            os.rename(old, new)
                            i += 1

        elif ".xls" in file :
            path = file_location + '\\' + file
            df = pd.read_excel(path)
            sel_cell = input("사용할 행 또는 열을 입력하시오. : ")
            num = re.sub(r'[^0-9]', '', sel_cell)

            if "행" in sel_cell:
                name_list = df.iloc[int(num)]
                if len(name_list) <= len(file_names):
                    i = 0
                    for file_name in file_names:
                        if i <= len(name_list):
                            print(file_path)
                            old = file_path + "\\" + file_name
                            dst = name_list[i] + '.jpg'
                            new = file_path + "\\" + dst
                            print(dst)
                            os.rename(old, new)
                            i += 1

                if len(name_list) > len(file_names):
                    i = 0
                    for file_name in file_names:
                        if i <= len(file_names):
                            print(file_path)
                            old = file_path + "\\" + file_name
                            dst = name_list[i] + '.jpg'
                            new = file_path + "\\" + dst
                            print(dst)
                            os.rename(old, new)
                            i += 1

            if "열" in sel_cell:
                name_list = df[int(num)]
                if len(name_list) <= len(file_names):
                    i = 0
                    for file_name in file_names:
                        if i <= len(name_list):
                            print(file_path)
                            old = file_path + "\\" + file_name
                            dst = name_list[i] + '.jpg'
                            new = file_path + "\\" + dst
                            print(dst)
                            os.rename(old, new)
                            i += 1

                if len(name_list) > len(file_names):
                    i = 0
                    for file_name in file_names:
                        if i <= len(file_names):
                            print(file_path)
                            old = file_path + "\\" + file_name
                            dst = name_list[i] + '.jpg'
                            new = file_path + "\\" + dst
                            print(dst)
                            os.rename(old, new)
                            i += 1

        # elif ".txt" in file :
        #     path = file_location + '\\' + file
        #     file = open(path)
        #     name_list = file.readlines()
        #     for j in name_list :
        #         name_list = name_list.strip("\n")
        #
        #     if len(name_list) <= len(file_names):
        #         i = 0
        #         for file_name in file_names:
        #             if i <= len(name_list):
        #                 print(file_path)
        #                 old = file_path + "\\" + file_name
        #                 dst = name_list[i] + '.jpg'
        #                 new = file_path + "\\" + dst
        #                 print(dst)
        #                 os.rename(old, new)
        #                 i += 1
        #
        #     if len(name_list) > len(file_names):
        #         i = 0
        #         for file_name in file_names:
        #             if i <= len(file_names):
        #                 print(file_path)
        #                 old = file_path + "\\" + file_name
        #                 dst = name_list[i] + '.jpg'
        #                 new = file_path + "\\" + dst
        #                 print(dst)
        #                 os.rename(old, new)
        #                 i += 1