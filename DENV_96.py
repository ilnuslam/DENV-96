from email.mime import image
import glob
from http.client import PRECONDITION_FAILED
from importlib.abc import PathEntryFinder
from operator import le
import re
import threading
import cv2
import numpy as np
import xlrd
import tkinter as tk
from tkinter import N, filedialog
import time
import sys
import os

import sqlio


def choose_xls_doc():
    #选择文件
    print("please choose .xls doc")
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', 1)
    excel_path = filedialog.askopenfilename(
        title="选择qPCR结果表格",
        filetypes=[("excel", "*.xls"), ("database", "*.db"), ("所有文件", "*.*")]
    )
    if not excel_path:
        print("未选择文件，程序退出")
        sys.exit(0)
    print(excel_path)
    file_type = os.path.splitext(excel_path)
    return excel_path, file_type 

def get_integer_input(prompt):
    while True:
        user_input = input(prompt)
        try:
            integer_value = int(user_input)
            #print(f"您输入的整数是：{integer_value}")
            return integer_value
        except ValueError:
            print("输入无效，请输入一个整数。")

def get_nonempty_list(nc_num):
    if '' in nc_num:
        print("The list is empty, cannot calculate the average.")
        b_list = []
        for element in nc_num:
            if element:
                b_list.append(element)
        print(b_list)
        level = np.mean(b_list)
        return level
    else:
        # 使用numpy.mean()函数计算平均值
        level = np.mean(nc_num)
        return level

def convert_letters_to_numbers(letter_list):
    letters_to_numbers = {
    'A': 1,
    'B': 2,
    'C': 3,
    'D': 4,
    'E': 5,
    'F': 6,
    'G': 7,
    'H': 8
}
    return [letters_to_numbers[letter] for letter in letter_list]

def extract_last_one_or_two_digits(s):
    # 尝试提取末尾的两位数字
    if len(s) >= 2 and s[-2:].isdigit():
        return int(s[-2:])
    # 如果末尾没有两位数字，则提取末尾的一位数字
    elif len(s) >= 1 and s[-1].isdigit():
        return int(s[-1])
    else:
        return None  # 如果没有找到数字，则返回None

def compare_list_with_number(lst, num):
    # 创建一个新列表，用于存储比较结果
    result = []
    zero_count = 0
    small_count = 0
    red_count = 0
    # 遍历列表中的每个元素
    for item in lst:
        # 如果元素小于给定的数，则添加0到新列表
        if item == -1:
            result.append(-1)
        elif item == 0:
            result.append(11)
            zero_count = zero_count + 1
        elif item < num/2:
            result.append(5)
            small_count = small_count + 1
        elif item < num:
            result.append(0)
        elif item > num and item < 2*num:
            result.append(3)
        elif item >= 2*num:
            result.append(1)
            red_count = red_count + 1
        else:
            result.append(10)
    print("警戒值:", red_count)
    print("零值:", zero_count)
    print("小于1/2NC:", small_count)

    return result

img_share = []
save_prog = 2
save_data = []

def image_show():
    global save_prog
    global save_data
    # 创建一个黑色图像，大小为960x720
    image = np.zeros((680, 920, 3), np.uint8)
    while True:
        show, result = img_share

        # 定义矩形的左上角和右下角坐标
        pt1 = (100, 100)  # 左上角坐标
        pt2 = (820, 580)  # 右下角坐标，这样矩形的宽度为720，高度为480
        step = (pt2[0] - pt1[0]) // 12

        # 计算每个方格的宽度和高度
        grid_width = (pt2[0] - pt1[0]) // step
        grid_height = (pt2[1] - pt1[1]) // step

        # 在每个方格内绘制圆
        cv2.rectangle(image, (0,0), (920,680), (0, 0, 0), -1)
        co = 0
        for i in range(8):
            for j in range(12):
                # 计算方格的左上角坐标
                grid_pt1 = (pt1[0] + j * step, pt1[1] + i * step)
                # 计算圆心的坐标
                center = (grid_pt1[0] + step // 2, grid_pt1[1] + step // 2)
                # 绘制圆
                if show == 0 :
                    cv2.putText(image, f'96-well DENV test', (10, 30), cv2.FONT_HERSHEY_SIMPLEX, 1, (255, 255, 255), 2)
                    cv2.circle(image, center, step // 2, (255, 0, 0), -1)  # 蓝色圆，填充
                else:
                    color = show[co]
                    if co in result:
                        colors = (255, 255, 255)
                        full = 1
                        cv2.circle(image, center, step // 2, (0, 0, 0), -1, cv2.LINE_AA)
                        cv2.putText(image, 'NC', (center[0] - step // 4, center[1] + step // 8), cv2.FONT_HERSHEY_COMPLEX, 0.7, (255,255,255), 1, cv2.LINE_AA)
                    elif color == 0:
                        colors = (0, 255, 0)
                        full = -1
                    elif color == 1:
                        colors = (0, 0, 255)
                        full = -1
                    elif color == -1 or color == '-1':
                        colors = (0, 0, 0)
                        full = -1
                    elif color == 3:
                        colors = (0, 255, 255)
                        full = -1
                    elif color == '2':
                        cv2.putText(image, f'go to cmd for next step', (10, 30), cv2.FONT_HERSHEY_SIMPLEX, 1, (255, 255, 255), 2)
                        colors = (255, 255, 255)
                        full = -1
                    elif color == 10:
                        colors = (255, 255, 255)
                        full = 1
                        cv2.putText(image, '?', (center[0] - step // 6, center[1] + step // 8), cv2.FONT_HERSHEY_COMPLEX, 0.7, (255,255,255), 1, cv2.LINE_AA)
                    elif color == 11:
                        colors = (255, 255, 255)
                        full = 1
                        cv2.putText(image, '0', (center[0] - step // 6, center[1] + step // 8), cv2.FONT_HERSHEY_COMPLEX, 0.7, (255,255,255), 1, cv2.LINE_AA)
                    elif color == 5:
                        colors = (255, 255, 255)
                        full = 1
                        cv2.putText(image, '<NC/2', (center[0] - step // 3, center[1] + step // 8), cv2.FONT_HERSHEY_COMPLEX, 0.4, (255,255,255), 1, cv2.LINE_AA)
                    
                    cv2.circle(image, center, step // 2, colors, full, cv2.LINE_AA)
                    co = co + 1

        # 绘制横向直线
        for y in range(pt1[1], pt2[1] + 1, step):
            cv2.line(image, (pt1[0], y), (pt2[0], y), (255, 255, 255), 1)  # 红色直线，线条宽度为1

        # 绘制纵向直线
        for x in range(pt1[0], pt2[0] + 1, step):
            cv2.line(image, (x, pt1[1]), (x, pt2[1]), (255, 255, 255), 1)  # 红色直线，线条宽度为1
        # 绘制矩形
        cv2.rectangle(image, pt1, pt2, (255, 255, 255), 2)  # 白色矩形，线条宽度为2
    
        #绘制列编号
        yletters = "ABCDEFGH"
        for i, letter in enumerate(yletters):
            y_start = i * step  # 计算每个字母的起始y坐标
            org = (pt1[0] - step // 2, y_start + pt1[1] + step // 2)
            cv2.putText(image, letter, org, cv2.FONT_HERSHEY_SIMPLEX, 0.5, (255, 255, 255), 1, cv2.LINE_AA)

        #绘制行编号
        for i in range(1, 13):
            x_start = i * step  # 计算每个字母的起始x坐标
            org = (x_start + pt1[0] - step // 2 , pt1[1] - step // 4)
            cv2.putText(image, str(i), org, cv2.FONT_HERSHEY_SIMPLEX, 0.5, (255, 255, 255), 1, cv2.LINE_AA)

        # 显示图像
        cv2.imshow("96-well qPCQ", image)

        kee = cv2.waitKey(1)
        if show != 0 and '2' not in str(show):
            save_data = image
            key = cv2.waitKey(0)
            if key == ord('q'):
                print("图形程序退出")
                break
        
    cv2.destroyAllWindows()

def denv_96():
    image_show = 0, 0
    global img_share
    global save_prog
    global save_data

    img_share = image_show
    print("适用数据：sds7500 96孔板 第30个循环 xls/db格式")

    #选择Excel文件
    input_source, file_type = choose_xls_doc()
    path = os.path.dirname(input_source)
    file_name = os.path.basename(input_source)
    print(file_type[1])

    if file_type[1] == ".xls":
        workbook = xlrd.open_workbook(input_source)        # 打开Excel文件
        sheet = workbook.sheet_by_name('Multicomponent Data')        # 获取工作表
        data = []        # 初始化一个列表来存储D列的数据
        
        # 循环读取C2793到C2888单元格的内容
        for row in range(2792, 2888):  # 注意行索引从1开始，对应D1到D96
            cell_value = sheet.cell_value(row, 2)  # 列索引从0开始，D列是第4列，索引为3
            data.append(cell_value)

    elif file_type[1] == ".db":
        # 打开db文件
        data = sqlio.get_sql_list(input_source, "imported_data", ["id", "original_file_name", "updated_time"])
    else:
        print("document error!")

    #print(data)
    # 遍历列表中的每个元素，用-1替换空字符串
    cleaned_data = data
    cleaned_data = ['-1' if item == '' else item for item in data]
    
    #预览有数据的孔
    preview_data = data
    preview_data = ['2' if item != '' else item for item in data]
    preview_data = ['-1' if item == '' else item for item in preview_data]
    img_share = (preview_data, [-1])

    nc_count = get_integer_input("请输入阴性对照个数：")
    nc_count = int(nc_count)

    nc_exc = []
    if nc_count == 0:
        level = input("输入阴性均值：").strip()
        result = [100]
    else:
        for i in range(nc_count):
            element = input(f"请输入第{i+1}个阴性对照格号(如A1)：")
            nc_exc.append(element)
        #根据输入的单元格格式nc_exc转换成列表排序
        letters = [item[0] for item in nc_exc if item[0].isalpha()]
        numbers = [extract_last_one_or_two_digits(s) for s in nc_exc]
        letter_tr = convert_letters_to_numbers(letters)
        result = [12 * (int(x) - 1) + (int(y) - 1) for x, y in zip(letter_tr, numbers)]

        #从data检索nc对应数值
        nc_num = [data[i] for i in result]

        #计算均值，返回均值（如果有空字符串则删除后再计算平均值）
        level = get_nonempty_list(nc_num)
        print("阴性均值：",level)

    show = compare_list_with_number(np.array(np.asarray(cleaned_data, dtype=float)), np.asarray(level, dtype=float))
    img_share = (show, result)
    
    print("是否保存？[y/n]")
    save_prog = input()
    if save_prog == 'y':
        root = tk.Tk()
        root.withdraw()
        #root.attributes('-topmost', 1)
        save_path = filedialog.asksaveasfilename(initialfile = file_name[:-4], defaultextension = ".jpg",
            filetypes=[("JPEG files", "*.jpg"),  ("PNG files", "*.png"), ("所有文件", "*.*")]
        )
        if save_path:
            print(save_path)
            ext = save_path[-4:]
            cv2.imencode(ext, save_data)[1].tofile(save_path)
            print(f"Image saved to {save_path}")
        else:
            print("No file path selected.")

        print(os.path.dirname(save_path) + "/test.db")
        sqlio.create_table(os.path.dirname(save_path) + "/test.db")
        sqlio.insert_from_list(list(data), os.path.dirname(save_path) + "/test.db", file_name)

    elif save_prog == 'n':
        pass
    print("计算程序退出，在图形界面按'q'退出图形程序")

def main():
    # 启动参数线程
    param_thread = threading.Thread(target=denv_96)
    display_thread = threading.Thread(target=image_show)
    
    param_thread.start()
    display_thread.start()

    # 等待所有线程完成
    display_thread.join()
    param_thread.join()

if __name__ == "__main__":
    main()