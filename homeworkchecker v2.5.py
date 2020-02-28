from openpyxl import load_workbook
from openpyxl import Workbook


import os
import json
import types


def check_address(address):  # 检查路径是否标准，文件类型是否支持
    if os.path.isfile(address):
        if address.endswith(".xlsx"):
            print("文件检查正确")
            return True
        else:
            print("文件不支持请重新输入路径")
            return False
    else:
        print("找不到目标文件，请重新输入")
        return False


def set_homework_address():  # 获取作业文件路径
    fileaddress = input("请输入作业路径: ")
    return fileaddress


def set_answer_address():  # 获取答案文件路径
    fileaddress = input("请输入答案路径: ")
    return fileaddress


def get_homework_list(homework_sheet):  # 将作业转换成列表
    homework_list = [["名次"]["姓名"][["答案"],[]]]
    row_homework_sheet = 1#初始化行索引
    homework_row_max = homework_sheet.max_row
    print("作业最大行数为："+str(homework_row_max))
    homework_column_max=homework_sheet.max_column
    print("作业最大列数为："+str(homework_column_max))
    while(row_homework_sheet<=homework_row_max):
        column_homework_sheet=1#初始化列索引
        while(column_homework_sheet<=homework_column_max):
            value=homework_sheet.cell(row_homework_sheet,column_homework_sheet).value
            homework_list.append(value)#装载数据
            column_homework_sheet+=1#指向下一列
        row_homework_sheet+=1#指向下一行
    return homework_list


def get_anser_list(answer_sheet):  # 将答案转换成列表
    answer_list = []
    row_answer_sheet = 1
    answer_row_max = answer_sheet.max_row
    print("答案表的最大长度为："+str(answer_row_max))
    while(row_answer_sheet <= answer_row_max):
        cell_value = answer_sheet.cell(row_answer_sheet, 1).value
        cell_value = rectify_vlaue_type(cell_value)
        answer_list.append(cell_value)
        row_answer_sheet += 1
    print("答案列表为：")
    print(answer_list)
    return answer_list


def rectify_vlaue_type(value):  # 纠正数据类型为string
    if(not type(value) == type("a")):
        value = str(value)
    return value


def load_homework_sheet(address):  # 加载目标文件
    myworkbook = load_workbook(address)
    return myworkbook


def get_work_sheet(workbook):  # 获取默认工作表
    sheet = workbook.active
    return sheet


def delete_duplication_data(grade_list):  # 处理重复的数据

    return grade_list


def compute_grade(homework_sheet, answer):  # 计算成绩

    return grade_list


def load_json(address):  # 加载json默认设置文件
    with open("config.json") as json_file:
        config_dict = json.load(json_file)
    return config_dict


def create_json():  # 生成json文件
    config_dict = {"defualt start row": 2,
                   "default start column": 7, "kill_list": []}
    with open("config.json", "w") as json_file:
        json.dump(config_dict, json_file)
    json_file.flush()
    json_file.close()


def is_first_setup():  # 判断是否第一次启动
    if not os.path.isfile("config.json"):
        return True
    else:
        return False


def applicate_setting(config_dict, kill_list, row_start, column_start):  # 配置设置
    kill_list = config_dict["kill_list"]
    row_start = config_dict["defualt start row"]
    column_start = config_dict["default start column"]


def check_answer_row(homework_sheet, row_start, answer_list):  # 检查答案与作业是否匹配
    if(((homework_sheet.max_row-row_start)+1) == len(answer_list)):
        return True
    else:
        return False


if __name__ == "__main__":
    print("欢迎使用")

    kill_list = []  # 剔除人的名单，全局变量
    row_start = None
    column_start = None

    if is_first_setup():
        print("检测到您是第一次使用，正在生成相应配置文件")
        create_json()
    config_dict = load_json("config.json")
    kill_list = config_dict["kill_list"]
    row_start = config_dict["defualt start row"]
    column_start = config_dict["default start column"]
    print("去除关键词名单:")
    print(kill_list)
    print("起始行为："+str(row_start))
    print("起始列为："+str(column_start))

    #输入作业地址并检测
    while(True):
        homework_address = set_homework_address()
        if check_address(homework_address):
            break
    homework_workbook = load_workbook(homework_address)
    homework_sheet = get_work_sheet(homework_workbook)

    #输入答案地址并检测
    while(True):
        answer_address = set_answer_address()
        if check_address(answer_address):
            break
    answer_workbook = load_workbook(answer_address)
    answer_sheet = get_work_sheet(answer_workbook)
    answer_list = get_anser_list(answer_sheet)

    #循环验证答案与作业是否匹配
    while(True):
        if(not check_answer_row(homework_sheet, row_start, answer_list)):
            print("答案位数错误")
            #输入答案地址并检测
            while(True):
                answer_address = set_answer_address()
                if check_address(answer_address):
                    break
            answer_workbook = load_workbook(answer_address)
            answer_sheet = get_work_sheet(answer_workbook)
            answer_list = get_anser_list(answer_sheet)
        else:
            break

    pass
