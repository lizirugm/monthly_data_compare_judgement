import xlwings as xw

from rule.read_rule import read_rule

if __name__ == "__main__":
    xlapp = xw.App(visible=True, add_book=False)

    # 配置允差表表的地址
    setting_workbook_path = "./setting/允差表-202105042117.xlsx"
    # 配置比对结果记录表的地址
    data_workbook_path = "./data/比对结果记录表-202105042117.xlsx"

    rule_list = read_rule(xlapp, setting_workbook_path, ["铁", "钢", "炉渣", "烧结矿", "球团矿"])

    data_list=read_data()
    xlapp.quit()

