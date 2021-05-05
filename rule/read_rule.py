# 规则：
#     下划线：read_rule
#     大驼峰:ReadRule
#     小驼峰:readRule
from rule.judgement_rule import JudgementRule


def read_rule_from_start_point_vertical(sheet_entity, start_point) -> list:
    """
    从固定的子表的固定起始点开始读取规则\
    这个函数使用纵向遍历的方式读取规则
    :param sheet_entity: 子表实体
    :param start_point: 起始点
    :return: 这个纵向的所有规则的列表
    """
    res_list = []
    start_range = sheet_entity.range(start_point)
    while start_range.merge_cells:
        merge_area = start_range.merge_area
        next_start_range = merge_area.end("down")
        # next_start_range = down_range.offset(column_offset=1)
        rule_range = merge_area.expand(mode='right')
        jr=JudgementRule.create_by_sheet_range(rule_range)
        res_list.append(jr)

        # print(rule_range.value)
        start_range = next_start_range
    return res_list


def read_rule(xlapp, file_path, sheet_list) -> list:
    """
    读取规则表，找到样品元素的规则（规则用一个特殊格式表示）\

    :param xlapp:  xlwings实体
    :param file_path:  规则表的文件的路径
    :param sheet_list: 规则表的哪些sheet含有规则，且代表规则表的第一个键:list
    :param start_point:起始点 TODO 现在默认的起始点是A2和E2，以后如果需要更改再说
    :return: 一个规则列表，list，里面包含所有的规则，规则用特殊结构存储
    """
    setting_workbook = xlapp.books.open(file_path)

    return_list = []
    for sheet_name in sheet_list:
        # sheet_entity:sheet实体，表示现在操作的是哪个子表
        sheet_entity = setting_workbook.sheets[sheet_name]
        # start_point:起始识别的格子
        start_point_list = ["A2", "E2"]
        for start_point in start_point_list:
            # 一个起始点的规则列，是总规则列的一部分
            site_rule_list = read_rule_from_start_point_vertical(sheet_entity, start_point)
            return_list.extend(site_rule_list)
    setting_workbook.close()
    return return_list
