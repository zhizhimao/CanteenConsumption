#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
创建时间：Fri Aug 17 16:01:26 2018
作者: 星空飘飘
平台：Anaconda 3-5.1.0
语言版本：Python 3.6.4
编辑器：Spyder
分析器：Pandas: 0.22.0
解析器：lxml: 4.1.1
数据库：MongoDB 2.6.12
程序名：Canteen.py
"""

import pandas as pd
import random


class Canteen(object):
    def __init__(self):
        self.df = pd.read_excel('ic.xls')
        # 获取1号消费机数据 川人厨匠
        self.num1 = self.df.loc[self.df['机号'] == 1]
        # 获取2号消费机数据 龙文
        self.num3 = self.df.loc[self.df['机号'] == 3]
        # 获取4号消费机数据 坝坝面
        self.num4 = self.df.loc[self.df['机号'] == 4]

    def department(self):
        # 部门费用
        department_name = []
        department_sum = []
        department_list = self.df['部门信息']
        department_list = department_list.drop_duplicates()  # 过滤重复部门
        for name in department_list:
            consumption = self.df.loc[(self.df['部门信息'] == name)]['消费金额'].sum()  # 消费合计
            if consumption != 0:
                department_name.append(name)
                department_sum.append(consumption)
        department_name.append('熊易华')
        department_name.append('邹永红')
        xyh = self.df.loc[(self.df['姓名'] == '熊易华')]['消费金额'].sum()
        department_sum.append(xyh)
        zyh = self.df.loc[(self.df['姓名'] == '邹永红')]['消费金额'].sum()
        department_sum.append(zyh)
        df_department = pd.DataFrame([dict(zip(department_name, department_sum))])
        df_department.rename(columns={'一所': '设计一院', '三所': '设计三院', '二所': '设计二院', '九所': '第二分公司', '技术质量部': '科技质量部'}, inplace=True)
        df_department['总经办']= df_department['总经办'] - xyh -zyh
        df_department = df_department.T
        df_department = df_department.reindex(['设计一院', '设计二院', '设计三院', '水环境院', '经营部', '综合部', '科技质量部', '第二分公司', '投资', '大明行政', '大明一所', '总经办', '熊易华', '邹永红', '财务部', '监审部', '正恒泰'])
        df_department.loc['合计'] = df_department.sum()[0]  # 最后一行添加合计
        df_department['部门签字'] = ''  # 添加新的一列
        df_department.rename(columns={0: '部门费用'}, inplace=True)
        return df_department

    def plate(self):
        # 板块费用
        # 集团
        group = self.df.loc[(self.df['部门信息'] == '总经办') | (self.df['部门信息'] == '财务部') | (self.df['部门信息'] == '监审部')]['消费金额'].sum()
        # 投资
    #    investment = df.loc[(df['部门信息'] == '投资')]['消费金额'].sum()
        # 造价
        manufacturing_cost = self.df.loc[(self.df['部门信息'] == '大明行政') | (self.df['部门信息'] == '大明一所')]['消费金额'].sum()
        # 正恒泰
        zht = self.df.loc[(self.df['部门信息'] == '正恒泰')]['消费金额'].sum()
        # 设计板块
        # plate = ['总经办','财务部','监审部','投资','大明行政','大明一所','正恒泰'] #板块公司名称
        design = self.df.loc[(self.df['部门信息'] == '三所') | (self.df['部门信息'] == '技术质量部') | (self.df['部门信息'] == '水环境院') | (self.df['部门信息'] == '投资') | (self.df['部门信息'] == '经营部') | (self.df['部门信息'] == '二所') | (self.df['部门信息'] == '一所') | (self.df['部门信息'] == '综合部') | (self.df['部门信息'] == '九所')]['消费金额'].sum()

        total_plate = {'集团': group, '造价': manufacturing_cost, '正恒泰': zht, '设计': design}
        df_plate = pd.DataFrame([total_plate])
        df_plate = df_plate.T
        df_plate.rename(columns={0: '板块费用'}, inplace=True)
        return df_plate

    def plate_canteen(self):
        '''板块费用明细'''
        # 集团
        num1_group = self.num1.loc[(self.num1['部门信息'] == '总经办') | (self.num1['部门信息'] == '财务部') | (self.num1['部门信息'] == '监审部')]['消费金额'].sum()
        num3_group = self.num3.loc[(self.num3['部门信息'] == '总经办') | (self.num3['部门信息'] == '财务部') | (self.num3['部门信息'] == '监审部')]['消费金额'].sum()
        num4_group = self.num4.loc[(self.num4['部门信息'] == '总经办') | (self.num4['部门信息'] == '财务部') | (self.num4['部门信息'] == '监审部')]['消费金额'].sum()
        group = {'川人厨匠': num1_group, '龙文': num3_group, '坝坝面': num4_group}
        # 造价
        num1_manufacturing_cost = self.num1.loc[(self.num1['部门信息'] == '大明行政') | (self.num1['部门信息'] == '大明一所')]['消费金额'].sum()
        num3_manufacturing_cost = self.num3.loc[(self.num3['部门信息'] == '大明行政') | (self.num3['部门信息'] == '大明一所')]['消费金额'].sum()
        num4_manufacturing_cost = self.num4.loc[(self.num4['部门信息'] == '大明行政') | (self.num4['部门信息'] == '大明一所')]['消费金额'].sum()
        manufacturing_cost = {'川人厨匠': num1_manufacturing_cost, '龙文': num3_manufacturing_cost, '坝坝面': num4_manufacturing_cost}
        # 正恒泰
        num1_zht = self.num1.loc[(self.num1['部门信息'] == '正恒泰')]['消费金额'].sum()
        num3_zht = self.num3.loc[(self.num3['部门信息'] == '正恒泰')]['消费金额'].sum()
        num4_zht = self.num4.loc[(self.num4['部门信息'] == '正恒泰')]['消费金额'].sum()
        zht = {'川人厨匠': num1_zht, '龙文': num3_zht, '坝坝面': num4_zht}
        # 设计
        num1_design = self.num1.loc[(self.num1['部门信息'] == '三所') | (self.num1['部门信息'] == '技术质量部') | (self.num1['部门信息'] == '水环境院') | (self.num1['部门信息'] == '投资') | (self.num1['部门信息'] == '经营部') | (self.num1['部门信息'] == '二所') | (self.num1['部门信息'] == '一所') | (self.num1['部门信息'] == '综合部') | (self.num1['部门信息'] == '九所')]['消费金额'].sum()
        num3_design = self.num3.loc[(self.num3['部门信息'] == '三所') | (self.num3['部门信息'] == '技术质量部') | (self.num3['部门信息'] == '水环境院') | (self.num3['部门信息'] == '投资') | (self.num3['部门信息'] == '经营部') | (self.num3['部门信息'] == '二所') | (self.num3['部门信息'] == '一所') | (self.num3['部门信息'] == '综合部') | (self.num3['部门信息'] == '九所')]['消费金额'].sum()
        num4_design = self.num4.loc[(self.num4['部门信息'] == '三所') | (self.num4['部门信息'] == '技术质量部') | (self.num4['部门信息'] == '水环境院') | (self.num4['部门信息'] == '投资') | (self.num4['部门信息'] == '经营部') | (self.num4['部门信息'] == '二所') | (self.num4['部门信息'] == '一所') | (self.num4['部门信息'] == '综合部') | (self.num4['部门信息'] == '九所')]['消费金额'].sum()
        design = {'川人厨匠': num1_design, '龙文': num3_design, '坝坝面': num4_design}
        # plate_canteen = pd.DataFrame([group,investment,manufacturing_cost,zht,design],index=('集团','投资','造价','正恒泰','设计')) #设置index
        plate_canteen = pd.DataFrame([group, manufacturing_cost, zht, design])
        plate_canteen['合计'] = plate_canteen.apply(lambda x: x.sum(), axis=1)  # 行求和
        # plate_canteen = plate_canteen.append(plate_canteen.apply(lambda x: x.sum()),ignore_index=True) #行求和
        plate_canteen = plate_canteen.append(plate_canteen.sum(),ignore_index=True)  # 行求和同上一样
        plate_canteen = plate_canteen.rename({0: '集团', 1: '造价', 2: '正恒泰', 3: '设计', 4: '合计'})
        return plate_canteen


def generate(n, sum_v):
    '''随机生成食堂费用'''
    Vector = [random.random() for i in range(n)]
    Vector = [int(i / sum(Vector) * sum_v) for i in Vector]
    if sum(Vector) < sum_v:
        Vector[0] += sum_v-sum(Vector)
    return Vector


def detailed_list(money):
    '''食堂明细清单'''
    vegetable_list = {'菜品': ['米', '猪肉', '油', '青菜', '四季豆', '鸡蛋', '冬瓜', '土豆', '藕', '西红柿', '木耳', '平菇', '白萝卜', '黄瓜', '猪肝', '鸡肉', '豆芽', '青椒', '红椒', '豌豆尖', '紫菜', '胡萝卜', '南瓜', '山药', '青笋']}
    price_list = {'单价': [3.2, 15.0, 8.0, 3.0, 4.0, 8.0, 1.0, 2.5, 4.5, 4.0, 8.0, 6.0, 2.0, 3.5, 10, 29.0, 2.0, 4.0, 4.0, 3.0, 8.0, 2.0, 1.0, 6.0, 3.5]}
    number_list = {'数量': []}
    total_list = {'金额': generate(25, money)}
    vegetable = pd.DataFrame(vegetable_list)
    price = pd.DataFrame(price_list)
    number = pd.DataFrame(number_list)
    total = pd.DataFrame(total_list)
    df_list = [vegetable, price, number, total]
    df = pd.concat(df_list, axis=1)
    df['数量'] = df['金额'] / df['单价']
    df = df.round(1)  # 保留1位小数
    return df


if __name__ == '__main__':
    run = Canteen()
    sign = run.department()  # 部门签字单
    plate_money = run.plate()  # 板块费用
    total = run.plate_canteen()  # 食堂、板块费用合计
    # 获取各板块消费金额
    group = plate_money['集团':'集团']['板块费用'][0]
    manufacturing_cost = plate_money['造价':'造价']['板块费用'][0]
    design = plate_money['设计':'设计']['板块费用'][0]
    zht = plate_money['正恒泰':'正恒泰']['板块费用'][0]
    # 生成菜品清单
    group_detailed = detailed_list(group)  # 集团
    manufacturing_cost_detailed = detailed_list(manufacturing_cost)  # 造价
    design_detailed = detailed_list(design)  # 设计
    zht_detailed = detailed_list(zht)  # 正恒泰

    if sign.iloc[-1:, 0][0] == plate_money.sum()[0] and total.iloc[-1:, -1][0]:
        writer = pd.ExcelWriter('食堂费用分摊.xlsx')
        sign.to_excel(writer, sheet_name='部门分摊')
        total.to_excel(writer, sheet_name='板块费用')
        group_detailed.to_excel(writer, sheet_name='集团清单', index=False)
        manufacturing_cost_detailed.to_excel(writer, sheet_name='造价清单', index=False)
        design_detailed.to_excel(writer, sheet_name='设计清单', index=False)
        zht_detailed.to_excel(writer, sheet_name='正恒泰清单', index=False)
        writer.save()
