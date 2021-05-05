import os

import xlwings

from data.read_data import read_data
from output.render_data import render_data
from rule.read_rule import read_rule

if __name__ == "__main__":
    xlapp = xlwings.App(visible=True, add_book=False)
    try:
        # 配置允差表表的地址
        setting_workbook_path = "./setting/允差表-202105042117.xlsx"
        # 配置比对结果记录表的地址
        data_workbook_path = "./data/比对结果记录表-202105042117.xlsx"
        output_path = r"C:\Users\admin\Desktop\比对数据判定表.xls"

        sheet_list = ["铁", "钢", "炉渣", "烧结矿", "球团矿"]
        rule_list = read_rule(xlapp, setting_workbook_path, sheet_list)

        data_list = read_data(xlapp, data_workbook_path, sheet_list, rule_list)
        # 如果输出文件已存在，则删除它
        if os.path.exists(output_path):
            os.remove(output_path)

        render_data(xlapp, sheet_list, data_list, output_path)
    finally:
        xlapp.quit()
