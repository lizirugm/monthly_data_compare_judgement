from decimal import Decimal
from typing import Any, Optional


class JudgementRule:
    """
    表示一个样品的元素的规则
    """
    # 样品名称/类型
    sample_name = None
    # 元素名称
    element_name = None
    # 子规则列表
    site_rule_list = []

    def __init__(self):
        # 样品名称/类型
        sample_name = None
        # 元素名称
        element_name = None
        # 子规则列表
        site_rule_list = []
        pass

    def judge_standard_value(self, standard_value) -> Optional[Decimal]:
        """
        传入标准值/原结果，判断规则内是否有允差并返回
        :param standard_value:  标准值/原结果
        :return:  如果有允差返回Decimal允差，如果不符合任何一个规则，则返回None
        """
        result_list = [x for x in self.site_rule_list if x.judge_value_in_rule(standard_value) == True]
        if len(result_list) > 0:
            # 如果筛选结果有数，则筛选成功，否则失败，代表这个传入的数字不符合任何一个子规则
            return result_list[0].franchise
        else:
            return None

    @staticmethod
    def create_by_sheet_range(xlrange):
        """
        通过xlwings的range来创建这个类\
        举例格式\
        [['Si', '0.050-0.100', 0.012],\
         [None, '>0.100-0.250', 0.02],\
          [None, '>0.250-0.500', 0.03], \
          [None, '>0.500-1.00', 0.04],\
           [None, '>1.0-2.50', 0.065], \
           [None, '>2.50-4.00', 0.13],
           [None,'>4.00',None]]\

        :param xlrange: xl的range，格式比较固定
        :return: 创建好的类
        """
        xl_value = xlrange.value
        element_name = xl_value[0][0]
        sheet_name = xlrange.sheet.name
        judge_rule = JudgementRule()
        judge_rule.element_name = element_name
        judge_rule.sample_name = sheet_name
        range_list = xlrange.value
        for site_list in range_list:
            if len(site_list) == 3 and site_list[1] is not None and site_list[2] is not None:
                sr1 = SiteRule.create(site_list[1], site_list[2])
                judge_rule.site_rule_list.append(sr1)
        return judge_rule


class SiteRule:
    """
    子规则
    只包含返回的允差结果，和判断函数
    """
    min_value = None
    max_value = None
    min_type = ">"
    max_type = "<="
    # 允差
    franchise = 0

    def __init__(self):
        min_value = 0
        max_value = 0
        min_type = ">"
        max_type = "<="
        # 允差
        franchise = 0
        pass

    @staticmethod
    def create(range_str, frachise):
        site_rule = SiteRule()
        site_rule.franchise = frachise
        list1 = str(range_str).split('-')
        if (len(list1) == 1):
            try:
                unknow_value = list1[0]
                if unknow_value.__contains__('>'):
                    site_rule.max_value = 99999
                    site_rule.min_value = Decimal(unknow_value.replace('>', ''))
                elif unknow_value.__contains__('<'):
                    site_rule.min_value = 0
                    site_rule.max_value = Decimal(unknow_value.replace('<', ''))
                else:
                    site_rule.min_value = site_rule.max_value = Decimal(unknow_value)
                    site_rule.min_type = ">="
            except:
                return site_rule
        elif (len(list1) == 2):
            try:
                min_str = list1[0]
                max_str = list1[1]
                if not min_str.__contains__('>'):
                    site_rule.min_type = '>='
                    site_rule.min_value = Decimal(min_str)
                else:
                    min_str = min_str.replace('>', '')
                    site_rule.min_value = Decimal(min_str)
                if max_str.__contains__('<'):
                    site_rule.max_type = '<'
                    max_str = max_str.replace('<', '')
                    site_rule.max_value = Decimal(max_str)
                else:
                    site_rule.max_value = Decimal(max_str)
            except:
                return site_rule
        else:
            return site_rule
        return site_rule


    def judge_value_in_rule(self, val) -> bool:
        """
        判断这个值是否在这个规则里面
        :param val: 传入的值
        :return: 可用于表达式，返回bool
        """
        if val == self.min_value and self.min_type == '>=':
            return True
        elif val == self.max_value and self.max_type == '<=':
            return True
        elif self.max_value > val > self.min_value:
            return True
        else:
            return False
