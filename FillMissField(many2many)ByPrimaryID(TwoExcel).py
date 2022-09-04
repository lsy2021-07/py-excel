"""
相较于上个版本 可以将表二的多个字段映射到表1的多个字段
change：
    ex1_record_field and ex2_miss_field 都是list类型，表示对应的多个字段
"""

import os
import pandas as pd

currentPath = os.path.dirname(__file__)  #返回当前文件所在的路径

def fillField(ex1_name, ex1_primary_field, ex1_record_field, ex2_name, ex2_primary_field, ex2_miss_field, ex1_sheet_name='Sheet1', ex2_sheet_name='Sheet1'):
    '''
    ex1_primary_field 对应 ex2_primary_field
    @param ex1_name:  名称(包含路径)
    @param ex1_sheet_name:  名称(不包含路径)
    @param ex1_primary_field: 主键名称
    @param ex1_record_field:  填充字段名称
    @param ex2_name:
    @param ex2_sheet_name:
    @param ex2_primary_field:
    @param ex2_miss_field: 缺失字段名称
    @return:
    '''
    path1 = os.path.join(currentPath, ex1_name)
    path2 = os.path.join(currentPath, ex2_name)

    ex1 = pd.read_excel(path1,index_col=0,sheet_name=ex1_sheet_name)
    Dict = dict()

    for a,(index_ex1,b) in zip(ex1[ex1_primary_field], ex1[ex1_record_field].iterrows()):
        if not b.isna().all():
            # print(b)
            Dict[a] = b

    ex2 = pd.read_excel(path2,index_col=0,sheet_name=ex2_sheet_name)

    miss_field_column_index = [ex2.columns.get_loc(column) for column in ex2_miss_field]# 根据列字段获得列字段对应的是第几列 例如：'联系电话'是第五列 column_index=5
    num_miss = 0  #统计缺失值
    num_supplement = 0  # 统计补充条数

    for index, (a, (index_ex2, b)) in enumerate(zip(ex2[ex2_primary_field], ex2[ex2_miss_field].iterrows())):
        if b.isna().any():
            num_miss += 1
            # print(ex2.iloc[index, miss_field_column_index])
            if Dict.get(a) is not None:
                ex2.iloc[index, miss_field_column_index] = Dict.get(a)
                num_supplement += 1
                # print(f"查看替换的key{a}与value{Dict.get(a)}")

    print(f"补充{num_supplement},总缺失{num_miss}条,还剩{num_miss-num_supplement}条")

    ex2.to_excel(save_ex_name, sheet_name=save_ex_sheet_name)

    """
    存到指定excel:
    # ex2.to_excel(save_ex_name, sheet_name=save_ex_sheet_name)  
    """

ex1_name = 'resource/test_前山防疫名单.xlsx'
ex1_sheet_name = 'Sheet1'
ex1_primary_field = '身份证号码'
ex1_record_field = ['电话号码','序号']

ex2_name = 'resource/test_前山2022年广丰区劳动力人员信息登记统计表.xlsx'
ex2_sheet_name = '2022年广丰区劳动力人员信息登记统计表'
ex2_primary_field = '身份证号码'
ex2_miss_field = ['电话号码','号码序号']

save_ex_name = "output/example.xlsx"
save_ex_sheet_name = 'new_sheet_name'
fillField(ex1_name, ex1_primary_field, ex1_record_field, ex2_name, ex2_primary_field, ex2_miss_field, ex1_sheet_name=ex1_sheet_name, ex2_sheet_name=ex2_sheet_name)