from data.data_unity_entity import DataUnitedEntity


def render_data(xlapp, sheet_list,data_list,output_path):

    output_wb = xlapp.books.add()
    # 设置颜色的语句：
    # sht.range("A1").api.Font.Color=0x0000ff
    for sheet_name in sheet_list:
        this_sheet = output_wb.sheets.add(sheet_name)
        fit_data_unity_list = [x for x in data_list if x.sample_name == sheet_name]
        if len(fit_data_unity_list)>0:
            first_data_unity_entity:DataUnitedEntity=fit_data_unity_list[0]
            element_list=first_data_unity_entity.element_list
            title_list=["比对样号","比对项目"]
            title_list.extend(element_list)
            title_start_range=this_sheet.range("A1")
            title_start_range.value=title_list

            start_range = this_sheet.range("A2")
            for data_unity in fit_data_unity_list:
                next_start_range = data_unity.render_range(start_range)
                start_range = next_start_range

    output_wb.save(output_path)