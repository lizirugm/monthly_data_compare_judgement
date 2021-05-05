from data.data_entity import DataEntity
from data.data_unity_entity import DataUnitedEntity


def read_data_from_start_point_vertical(sheet_entity, start_point, element_list,rule_list):
    """
    从指定的sheet里面读取数据，结合元素列表，返回数据结果的格式列表
    :param sheet_entity:sheet实体
    :param start_point: 起始点
    :param element_list: 元素列表
    :return: 元素数据列表
    """
    res_list = []
    start_range = sheet_entity.range(start_point)
    while start_range.merge_cells:
        merge_area = start_range.merge_area
        next_start_range = merge_area.end("down")
        # next_start_range = down_range.offset(column_offset=1)
        data_range = merge_area.expand(mode='right')
        site_data_list = DataUnitedEntity.create(data_range, element_list,rule_list)
        res_list.append(site_data_list)

        # print(data_range.value)
        start_range = next_start_range
    return res_list


def read_data(xlapp, file_path, sheet_list,rule_list):
    """
    读取这些sheet里面的数据，然后返回结果列表
    :param file_path: 数据表的位置
    :param xlapp: xlwings实体
    :param sheet_list:  sheet列表
    :return:
    """

    data_workbook = xlapp.books.open(file_path)
    return_list = []
    for sheet_name in sheet_list:
        # sheet_entity:sheet实体，表示现在操作的是哪个子表
        try:
            sheet_entity = data_workbook.sheets[sheet_name]
        except :
            raise Exception(f"没有在这个workbook里面找到这个名称:{sheet_name}的子表，检查文件中sheet的名字是否符合:{sheet_list}这些名字")
        # start_point:起始识别的格子
        element_start = "C1"
        element_range = sheet_entity.range(element_start).expand('right')
        element_list = element_range.value
        start_point = "A2"

        site_data_list = read_data_from_start_point_vertical(sheet_entity, start_point, element_list,rule_list)
        return_list.extend(site_data_list)
    data_workbook.close()
    return return_list
