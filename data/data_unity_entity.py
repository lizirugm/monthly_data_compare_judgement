from decimal import Decimal

from data.data_entity import DataEntity
from setting.quantize_setting import quantize_setting


class DataUnitedEntity:
    """
    根据式样号组合好的数据格式
    """
    # 式样号
    mat_sample_no = None
    # 该式样号的元素类别
    element_list = []
    # 该式样号的大式样名称
    sample_name = None
    # 属于该式样号的数据列表
    data_list = []
    # 属于该式样号的规则子集
    site_rule_list_inherit = []

    data_range1_value = None

    def __init__(self):
        # 式样号
        self.mat_sample_no = None
        # 该式样号的元素类别
        self.element_list = []
        # 该式样号的大式样名称
        self.sample_name = None
        # 属于该式样号的数据列表
        self.data_list = []
        # 属于该式样号的规则子集
        self.site_rule_list_inherit = []

    @staticmethod
    def create(data_range, element_list, rule_list):
        """
       从range里面读数据
       :param data_range:
       :param element_list:
       :return:
       """
        new_due = DataUnitedEntity()

        range_list = data_range.value
        new_due.mat_sample_no = range_list[0][0]
        new_due.element_list = element_list

        new_due.sample_name = data_range.sheet.name
        data_list = DataUnitedEntity.get_data_list_from_range(element_list, new_due, range_list)

        new_due.data_list = data_list
        new_due.site_rule_list_inherit = [x for x in rule_list if x.sample_name == new_due.sample_name]
        return new_due

    @staticmethod
    def get_data_list_from_range(element_list, new_due, range_list):
        data_list = []
        data_inherit_list = []
        for item in range_list:
            new_list = item[2:]
            data_inherit_list.append(new_list)
        for index, origin_value in enumerate(data_inherit_list[0]):
            element_name = element_list[index]
            current_value = data_inherit_list[1][index]
            de = DataEntity.create(origin_value, current_value, element_name, new_due.sample_name, new_due.mat_sample_no)
            data_list.append(de)
        return data_list

    def render_range(self, start_range):
        """
        渲染这组数据到sheet里面

        :param start_point:开始range
        :return:下一个开始点的range
        """

        # 合并单元格：式样号下面四格（有允差）
        start_range.value = self.mat_sample_no
        msn_merge_rage = start_range.resize(row_size=4, column_size=None)
        # 下面一句是合并单元格
        msn_merge_rage.api.Merge()
        # 式样号右边一列是中文
        # 第一个中文位置
        origin_chn_descript_rng = start_range.offset(row_offset=0, column_offset=1)
        # sht.range('A1').options(transpose=True).value = [1, 2, 3]
        origin_chn_descript_rng.options(transpose=True).value = ["原结果", "现结果", "偏差", "允差"]
        data_start_range = origin_chn_descript_rng.offset(row_offset=0, column_offset=1)
        for element_item in self.element_list:
            fit_data_list = [x for x in self.data_list if x.element_name == element_item]
            if len(fit_data_list) == 0:
                data_start_range.options(transpose=True).value = [None, None, None, None]
            else:
                fit_data:DataEntity = fit_data_list[0]
                quantize_str = quantize_setting.get(self.sample_name, '0.000')
                this_data_vertical_4_range = data_start_range.resize(row_size=4, column_size=None)
                this_data_vertical_4_range.api.NumberFormat = "@"
                data_start_range.options(transpose=True).value = [fit_data.get_origin_value_str(),
                                                                       fit_data.get_current_value_str(),
                                                                       fit_data.get_diff_value_str(),
                                                                       fit_data.get_franchise_str(self.site_rule_list_inherit)]
                qualified = fit_data.get_qualified(self.site_rule_list_inherit)
                # 占据此次数据渲染的竖排四个格子

                # 允差判断结果：非空且不合格才染色
                if qualified is not None and qualified == False:
                    this_data_vertical_4_range = data_start_range.resize(row_size=4, column_size=None)
                    this_data_vertical_4_range.api.Font.Color = 0x0000ff
            data_start_range = data_start_range.offset(row_offset=0, column_offset=1)
        return start_range.offset(row_offset=4, column_offset=0)
