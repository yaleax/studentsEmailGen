import pandas as pd
import xlwt
import random
import re
import datetime
import os
import sys
import tkinter as tk
from tkinter import filedialog
import math
# 获取当前时间20220101格式
def get_current_year_month_day():
    current_date = datetime.datetime.now()
    year = current_date.year
    month = current_date.month
    day = current_date.day
    year_month_day = f"{year}{month:02}{day:02}"
    return year_month_day

# 修改数据成为指定格式
def modify_data(data,domian):
    modified_list = []
    # 调整数据
    for row in data:
        name = row[0]
        # 删除括号和括号内的内容
        name = re.sub(r'\([^)]*\)', '', name)
        # 将名字分成姓和名
        first, last = name.split(',')
        # 调换名字和姓氏
        first, last = last, first
        # 删除任何前/后空格
        first = first.strip()
        last = last.strip().lstrip()
        # 用空字符代替空格
        first = first.replace(" ", "")
        last = last.replace(" ", "")

        # 创建电子邮件地址
        email = f"{first.lower()}.{last.lower()}@students."+ domian
        email2 = f"{first.lower()}.{last.lower()}@students."+ domian
        # 修改年级
        grade = f"Grade {row[1]}"
        # 将修改后的数据添加到修改后的列表中
        modified_list.append([email, f"{first} {last}", first, last, f"{first.lower()}.{last.lower()}", email2, grade, row[2]])
    return modified_list

# 添加数据成为指定格式
def add_data(data,domian,school,company):
    for i, item in enumerate(data):
        # 生成随机字符串，直到生成的字符串在列表中不存在为止
        while True:
            # 生成一个11位的随机字符串
            random_string = ''.join(random.sample('ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnpqrstuvwxyz23456789', k=11))
            # 判断随机字符串是否已经存在
            if random_string not in [x[0] for x in data]:
                # 将随机字符串插入到列表的倒数第二位
                data[i] = item[:len(item)-1] + [random_string] + item[len(item)-1:]
                break
    for i, item in enumerate(data):
        # 将YHIS插入到列表的倒数第二位，并添加组数据，数字和列表最后一位一样
        data[i] = item[:len(item)-1] + [school.upper(), item[len(item)-1], item[len(item)-1]] + ["OU=Students,OU=" + school.upper()+",OU=Locations,DC="+company.upper()+",DC="+company.upper()+"china,DC=com"]
    return data

def forteacher(data, file_name):
    # 创建一个Workbook对象
    wb2 = xlwt.Workbook()
    # 创建一个Worksheet对象
    ws2 = wb2.add_sheet('Sheet1')
    # 写入表头
    header_row = ['Grade', 'StudentID', 'Name', 'Email', 'Password']
    for i, cell in enumerate(header_row):
        ws2.write(0, i, cell)
    # 写入数据
    for i, row in enumerate(data):
        ws2.write(i+1, 0, row[6])
        ws2.write(i+1, 1, row[9])
        ws2.write(i+1, 2, row[1])
        ws2.write(i+1, 3, row[0])
        ws2.write(i+1, 4, row[7])

    wb2.save('.' + os.sep + 'Email' + os.sep + str(current_date) + os.sep + file_name)

def formac(data, file_name):
    # 创建一个Workbook对象
    wb2 = xlwt.Workbook()
    # 创建一个Worksheet对象
    ws2 = wb2.add_sheet('Sheet1')
    # 写入表头
    header_row = ['Grade', 'StudentID', 'Name', 'Email']
    for i, cell in enumerate(header_row):
        ws2.write(0, i, cell)
    # 写入数据
    for i, row in enumerate(data):
        ws2.write(i+1, 0, row[6])
        ws2.write(i+1, 1, row[9])
        ws2.write(i+1, 2, row[1])
        ws2.write(i+1, 3, row[0])


    wb2.save('.' + os.sep + 'Email' + os.sep + str(current_date) + os.sep + file_name)

def forIthelp(data, filename):
    wb = xlwt.Workbook()
    ws = wb.add_sheet('Sheet1')
    # 写入表头
    header_row = ['UserPrincipalName', 'DisplayName', 'firstName', 'Initials', 'lastname', 'username', 'email', 'StreetAddress', 'City', 'ZipCode', 'State', 'Country', 'Department', 'Password', 'Telephone', 'Jobtitle', 'Company', 'Description', 'EmployeeID', 'EmployeeNumber', 'OU']
    for i, cell in enumerate(header_row):
        ws.write(0, i, cell)

    # 写入数据
    for i, row in enumerate(data):
        ws.write(i+1, 0, row[0])
        ws.write(i+1, 1, row[1])
        ws.write(i+1, 2, row[2])
        ws.write(i+1, 4, row[3])
        ws.write(i+1, 5, row[4])
        ws.write(i+1, 6, row[5])
        ws.write(i+1, 12, row[6])
        ws.write(i+1, 13, row[7])
        ws.write(i+1, 17, row[8])
        ws.write(i+1, 18, row[9])
        ws.write(i+1, 19, row[10])
        ws.write(i+1, 20, row[11])

    # 保存文件
    wb.save('.' + os.sep + 'Email' + os.sep + str(current_date) + os.sep + filename)

# 获取当前时间20220101格式
current_date = get_current_year_month_day()

# 询问用户输入
user_input = int(input("1.手动输入数据\n2.选择excel文件\n"))
if user_input == 1:
    def get_student_info():
        # 初始化空表格
        inputData = []

        # 循环添加学生信息
        while True:
            # 询问并获取学生的姓名
            first_name = input('请输入学生名: ')
            last_name = input('请输入学生姓: ')

            # 询问并获取学生的年级
            while True:
                try:
                    grade = int(input('请输入学生的年级（-2到12）: '))
                    if -2 <= grade <= 12:
                        break
                    else:
                        print('年级必须是大于-2小于12的整数，请新输入。')
                        continue
                except ValueError:
                    print('年级必须是大于-2小于12的整数。')
                    continue

            # 询问并获取学生的学号
            while True:
                try:
                    student_id = input('请输入学生的学号(7位或8位数字): ')
                    if student_id.isdigit() and (len(student_id) == 7 or len(student_id) == 8):
                        break
                    else:
                        print('请输入学生的学号(7位或8位数字): ')
                        continue
                except ValueError:
                    print('学号必须是7位或8位数字，请新输入。')
                    continue

            # 添加信息到列表中
            inputData.append([last_name + ", " + first_name, grade, student_id])

            # 判断是否继续添加学生信息
            add_another = input("是否还要继续添加学生？(y/n) ")
            if add_another.lower() == "y":
                continue
            elif add_another.lower() != "y":
                break

        # 返回所有学生信息列表
        return inputData

    #读取输入的信息列表
    data = get_student_info()

    user_input = input(
        f"你输入的数据是：{data}，这段数据你准备用来做什么？\n1.创建Email所需的数据格式。\n2.ID卡所需数据格式\n")
    if user_input == "1":
        def get_email_info():
            while True:
                domian = input("请输入你的学校域名: ")
                school = input("请输入你的学校名称: ")
                company = input("请输入你的公司名字: ")
                # 手动输入的数据
                # 修改数据成为指定格式
                inputData = modify_data(data, domian)
                # 添加数据成为指定格式
                inputData = add_data(inputData, domian, school, company)
                print("转换后的数据是：\n", inputData)

                # 标准数据（Grade1-12）
                # 筛选数据
                standard_list = [item for item in data if item[1] > 0]
                # 修改数据成为指定格式
                standardData = modify_data(standard_list, domian)
                # 添加数据成为指定格式
                standardData = add_data(standardData, domian,school, company)
                # 打印修改后的标准数据
                if standardData:
                    print("standardData的数据是:\n", standardData[0])
                else:
                    print("standardData的数据是空的")


                # Secondary数据（Grade6-12）
                # 筛选数据
                secondary_list = [item for item in data if item[1] >= 6]
                # 修改数据成为指定格式
                secondaryData = modify_data(secondary_list, domian)
                # 添加数据成为指定格式
                secondaryData = add_data(secondaryData, domian,school,company)

                # secondaryData
                if secondaryData:
                    print("secondaryData的数据是:\n", secondaryData[0])
                else:
                    print("secondaryData的数据是空的")

                # Elementary数据（Grade1-5）
                # 筛选数据
                elementary_list = [item for item in data if 6 > item[1] > 0]
                # 修改数据成为指定格式
                elementaryData = modify_data(elementary_list, domian)
                # 添加数据成为指定格式
                elementaryData = add_data(elementaryData, domian,school,company)
                # 打印elementaryData
                if elementaryData:
                    print("elementaryData的数据是:\n", elementaryData[0])
                else:
                    print("elementaryData的数据是空的")

                # ecc数据（Grade0--2）
                # 筛选数据
                ecc_list = [item for item in data if 0 >= item[1]]
                # 修改数据成为指定格式
                eccData = modify_data(ecc_list, domian)
                # 添加数据成为指定格式
                eccData = add_data(eccData, domian, school, company)
                # 打印ecc数据
                if eccData:
                    print("eccData的数据是:\n", eccData[0])
                else:
                    print("eccData的数据是空的")
                # 创建Emails文件夹
                folder_name = "./Email"
                if not os.path.exists(folder_name):
                    os.makedirs(folder_name)
                # 在Emails文件夹下创建以时间命名的文件名
                folder_name = "./Email/" + current_date
                if not os.path.exists(folder_name):
                    os.makedirs(folder_name)
                # 保存标准数据数据给ithlep
                forIthelp(inputData, 'Ithelp.xls')
                # 保存elementary文件给macteam
                formac(inputData, 'Mac.xls')
                # 保存standardData文件给IT
                forteacher(inputData, 'IT.xls')
                # 保存secondaryForTeacher.xls
                forteacher(secondaryData, 'secondary.xls')
                # 保存elementaryForTeacher
                forteacher(elementaryData, 'elementary.xls')
                # 保存 eccForTeacher
                forteacher(eccData, 'ecc.xls')
                print(f"数据保存成功,请在本程序目录下的Email/{current_date}查看文件！")
                break
            return inputData


        data = get_email_info()
    if user_input == "2":
        def get_id_info():
            while True:
                print(f"你选择了ID卡所需数据格式是：{data}")
                break
        data = get_id_info()
    else:
        print("输入错误，程序退出！")
        sys.exit()

        data = get_id_info()

elif user_input == 2:
    # 打开文件
    root = tk.Tk()
    root.withdraw()
    file_path = tk.filedialog.askopenfilename()


    # 读入xls文件
    df = pd.read_excel(file_path)

    # 打印读取的excel
    # print(df)
    # 打印第一行名字
    #lie = df.columns
    #print(lie)

    # 转换所有列名为小写
    df.columns = [col.lower() for col in df.columns]
    # 选择列名为 "lastfirst"、"grade" 和 "student_number" 的列
    columns_to_keep = ['lastfirst', 'grade', 'student_number']
    df = df.loc[:, columns_to_keep]
    # 将 "Student_number" 列和 "Grade" 列的数据转换为整数型
    df['student_number'] = df['student_number'].astype(int)
    df['grade'] = df['grade'].astype(int)

    # 打印抓取后的数据
    # print(df)
    data = df.values.tolist()


    # 删除excel李的空格，也就是列表中的nan（float）数据
    def remove_nan(data):
        new_list = []
        for sublist in data:
            new_sublist = []
            for item in sublist:
                if isinstance(item, float) and math.isnan(item):
                    continue
                new_sublist.append(item)
            if new_sublist:
                new_list.append(new_sublist)
        return new_list

    # 删除excel李的空格，也就是列表中的nan（float）数据
    filtered_list = remove_nan(data)
    data = filtered_list

    # 打印原始数据
    #print("原始数据",data)


    # 自动判断所需数据
    result = []
    for entry in data:
        new_entry = []
        for element in entry:
            if isinstance(element, str) and ',' in element:
                new_entry.insert(0, element)  # 插入第一位
            elif isinstance(element, int) and element >= -2 and element <= 12:
                new_entry.insert(1, element)  # 插入第二位
            elif isinstance(element, int) and element > 1000000 and element < 9000000:
                new_entry.insert(2, element)  # 插入最后一位
        result.append(new_entry[:3])
    data = result
    user_input = input(f"抓取到的数据是：{data}，这段数据你准备用来做什么？\n1.创建Email所需的数据格式。\n2.ID卡所需数据格式\n")
    if user_input == "1":
        def get_email_info():
            while True:

                # 定义Domian
                domian = input("请输入你的学校域名: ")
                school = input("请输入你的学校名称: ")
                company = input("请输入你的公司名字: ")

                # 所有年级的数据
                # 修改数据成为指定格式
                all_data = modify_data(data, domian)
                # 添加数据成为指定格式
                all_data = add_data(all_data, domian, school,company)
                # 打印修改后的所有数据
                print("所有的数据是:\n", all_data[0])

                # 标准数据（Grade1-12）
                # 筛选数据
                standard_list = [item for item in data if item[1] > 0]
                # 修改数据成为指定格式
                standardData = modify_data(standard_list, domian)
                # 添加数据成为指定格式
                standardData = add_data(standardData, domian,school, company)
                # 打印修改后的标准数据
                if standardData:
                    print("standardData的数据是:\n", standardData[0])
                else:
                    print("standardData的数据是空的")


                # Secondary数据（Grade6-12）
                # 筛选数据
                secondary_list = [item for item in data if item[1] >= 6]
                # 修改数据成为指定格式
                secondaryData = modify_data(secondary_list, domian)
                # 添加数据成为指定格式
                secondaryData = add_data(secondaryData, domian,school,company)

                # secondaryData
                if secondaryData:
                    print("secondaryData的数据是:\n", secondaryData[0])
                else:
                    print("secondaryData的数据是空的")

                # Elementary数据（Grade1-5）
                # 筛选数据
                elementary_list = [item for item in data if 6 > item[1] > 0]
                # 修改数据成为指定格式
                elementaryData = modify_data(elementary_list, domian)
                # 添加数据成为指定格式
                elementaryData = add_data(elementaryData, domian,school,company)
                # 打印elementaryData
                if elementaryData:
                    print("elementaryData的数据是:\n", elementaryData[0])
                else:
                    print("elementaryData的数据是空的")

                # ecc数据（Grade0--2）
                # 筛选数据
                ecc_list = [item for item in data if 0 >= item[1]]
                # 修改数据成为指定格式
                eccData = modify_data(ecc_list, domian)
                # 添加数据成为指定格式
                eccData = add_data(eccData, domian, school, company)
                # 打印ecc数据
                if eccData:
                    print("eccData的数据是:\n", eccData[0])
                else:
                    print("eccData的数据是空的")

                choice = input("请选择你要保存的数据：\n1.保存Grade1-12的数据\n2.保存ECC数据\n3.保存所有年级的数据\n4.退出\n")

                while True:
                    if choice == "1":
                        # 创建Emails文件夹
                        folder_name = "./Email"
                        if not os.path.exists(folder_name):
                            os.makedirs(folder_name)
                        # 创建以时间命名的文件名
                        folder_name = "./Email/"+ current_date
                        if not os.path.exists(folder_name):
                            os.makedirs(folder_name)
                        # 保存标准数据数据给ithlep
                        forIthelp(standardData, 'standard4Ithelp.xls')
                        # 保存elementary文件给macteam
                        formac(standardData, 'standard4Mac.xls')
                        # 保存standardData文件给IT
                        forteacher(standardData, 'standard4IT.xls')
                        # 保存elementary文件secondaryForTeacher.xls
                        forteacher(secondaryData, 'sec4Teacher.xls')
                        # 保存elementary文件elementaryForTeacher.xls
                        forteacher(elementaryData, 'ele4Teacher.xls')
                        print(f"数据保存成功,请在本程序目录下的Email/{current_date}查看文件！")
                        sys.exit()
                        break
                    if choice == "2":
                        # 创建Emails文件夹
                        folder_name = "./Email"
                        if not os.path.exists(folder_name):
                            os.makedirs(folder_name)
                        # 创建以时间命名的文件名
                        folder_name = "./Email/"+ current_date
                        if not os.path.exists(folder_name):
                            os.makedirs(folder_name)
                        #保存ECC数据数据给ithlep
                        forIthelp(eccData, 'ecc4Ithelp.xls')
                        # 保存ecc文件eccForTeacher.xls
                        forteacher(eccData, 'ecc4Teacher.xls')
                        #保存ecc文件给macteam
                        formac(eccData, 'ecc4Mac.xls')

                        print(f"数据保存成功,请在本程序目录下的Email/{current_date}查看文件！")
                        sys.exit()
                        break

                    if choice == "3":
                        # 创建Emails文件夹
                        folder_name = "./Email"
                        if not os.path.exists(folder_name):
                            os.makedirs(folder_name)
                        # 创建以时间命名的文件名
                        folder_name = "./Email/"+ current_date
                        if not os.path.exists(folder_name):
                            os.makedirs(folder_name)
                        # 保存标准数据数据给ithlep
                        forIthelp(all_data, 'allDate4Ithelp.xls')
                        # 保存elementary文件给macteam
                        formac(all_data, 'allDate4Mac.xls')
                        # 保存所有文件给IT
                        forteacher(all_data, 'allDate4IT.xls')
                        # 保存elementary文件secondaryForTeacher.xls
                        forteacher(secondaryData, 'sec4Teacher.xls')
                        # 保存elementary文件elementaryForTeacher.xls
                        forteacher(elementaryData, 'ele4Teacher.xls')
                        # 保存ecc文件eccForTeacher.xls
                        forteacher(eccData, 'ecc4Teacher.xls')
                        print(f"数据保存成功,请在本程序目录下的Email/{current_date}查看文件！")
                        sys.exit()
                        break

                    if choice == "4":
                        print("程序退出！")
                        sys.exit()
                        break
                    else:
                        print("输入错误，程序退出！")
                        sys.exit()
                        break
        data = get_email_info()
    if user_input == "2":
        def get_id_info():
            while True:
                print(f"你选择了ID卡所需数据格式是：{data}")
                break
        data = get_id_info()
else:
    print("输入错误，程序退出！")
    sys.exit()

