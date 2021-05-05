from decimal import Decimal

from setting.quantize_setting import quantize_setting


class DataEntity:
    """
    数据存储的特殊格式，附带计算结果
    """
    # 原始值
    origin_value = 0
    # 现在值
    current_value = 0

    # 比对样号
    mat_sample_no = None

    # 元素名称
    element_name = None

    # 样品名称
    sample_name = None

    def __init__(self):
        # 原始值
        self.origin_value = None
        # 现在值
        self.current_value = None
        # 元素名称
        self.element_name = None

        # 样品名称
        self.sample_name = None

        # 比对样号
        self.mat_sample_no = None

    @staticmethod
    def create(origin_value, current_value, element_name, sample_name, mat_sample_no):
        de = DataEntity()
        try:
            quantize_str = quantize_setting.get(sample_name, '0.000')
            de.mat_sample_no = mat_sample_no
            de.origin_value = Decimal(origin_value).quantize(Decimal(quantize_str))
            de.current_value = Decimal(current_value).quantize(Decimal(quantize_str))

        except:
            pass
        de.sample_name = sample_name
        de.element_name = element_name
        return de

    def get_origin_value_str(self):
        return format(self.origin_value, '.4f')

    def get_current_value_str(self):
        return format(self.current_value, '.4f')

    def get_diff_value(self):
        """
        偏差
        :return:
        """
        if self.origin_value is None:
            return None
        return self.origin_value - self.current_value

    def get_diff_value_str(self):
        """
        获得字串偏差
        :return:
        """
        if self.origin_value is None:
            return None
        return format(self.get_diff_value(), '.4f')

    def get_franchise(self, rule_list):
        """
        获得允差
        :param rule_list: 规则列表
        :return: 允差的值或者None，取决于是否有匹配上该原始值的规则
        """
        if self.origin_value is None:
            return None
        filter_list = [x for x in rule_list if x.sample_name == self.sample_name and x.element_name == self.element_name]
        if len(filter_list) > 0:
            fit_rule = filter_list[0]
            franchise = fit_rule.judge_standard_value(self.origin_value)
            return franchise
        else:
            return None

    def get_franchise_str(self,rule_list):
        self_fr=self.get_franchise(rule_list)
        if self_fr is None:
            return None
        else:
            return format(self_fr, '.4f')

    def get_qualified(self, rule_list):
        """
        获得是否合格
        :return:
        """
        if self.origin_value is None:
            return None
        # 首先获得允差，如果没有找到规则，则回调一个其他结果(None)
        fr = self.get_franchise(rule_list)
        if fr is None:
            return None
        else:
            fr=Decimal(fr)
        # 如果有允差，回调合格True或不合格False

        diff_abs = abs(self.get_diff_value())
        if diff_abs > abs(fr):
            return False
        else:
            return True
