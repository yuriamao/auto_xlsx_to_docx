# -*- coding: utf-8 -*-
"""
Created on Tue Aug 30 08:21:10 2022

@author: lxl
"""
import sys
import os

from auto_figure_func import *
from auto_figure_docx import *
import pandas as pd


# 读取配置文件
config_df = pd.read_excel(r'/Users/harvin/code/自动报告产品开发-产业链@20220830/data/pmi_config.xlsx',header=0)
# 读取数据
data_dim,data_week_df = data_process(r'/Users/harvin/code/自动报告产品开发-产业链@20220830/data/指标数据20231120.xlsx')
# 获取时间序列
time_list = list(pd.read_excel(r'/Users/harvin/code/自动报告产品开发-产业链@20220830/data/time.xlsx').iloc[:,0])

# 批量执行
# for each_t in time_list:
#     create_file(str(each_t),data_week_df,config_df,data_dim)

for each_t in time_list:
    try:
        create_file(str(each_t), data_week_df, config_df, data_dim,title0_name='产业上中下游价格指数周报',file_pre='hhh',file_dir='2')
    except Exception as e:
        print(f"出现异常: {e}")
        continue  # 出现异常时跳过当前迭代，继续执行下一个迭代
