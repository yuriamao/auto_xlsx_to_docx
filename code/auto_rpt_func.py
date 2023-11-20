# -*- coding: utf-8 -*-
"""
Created on Mon Aug 29 15:41:44 2022

@author: lxl
"""

import numpy as np
import pandas as pd
import locale
import datetime
import math

# locale.setlocale(locale.LC_CTYPE, 'chinese')

import locale

try:
    # 设置区域设置为简体中文（中国大陆）
    locale.setlocale(locale.LC_CTYPE, 'zh_CN.UTF-8')
    print("成功设置区域设置为简体中文（中国大陆）")
except locale.Error as e:
    print(f"无法设置区域设置: {e}")
    # 在失败时可以使用默认的区域设置
    # locale.setlocale(locale.LC_CTYPE, '')  # 或其他适合的默认区域设置


# 数据处理
def data_process(filename):
    data_df = pd.read_excel(filename,header=0,skiprows = range(1,7))
    data_df = data_df[:-2]
    data_dim = pd.read_excel(filename,header=0,nrows=6)
    
    # 将需要计算的数值列选取出来，假设这些列名都以数字类型开头
    numerical_columns = [col for col in data_df.columns if pd.api.types.is_numeric_dtype(data_df[col])]

    # 日、周提取
    data_df['date'] = pd.to_datetime(data_df['指标名称'], errors='coerce')
    data_df['week'] = data_df['date'].dt.strftime('%Y-w%U')

    # 按周计算数值列的平均值
    data_week_df = data_df[numerical_columns + ['week']].groupby('week').mean()
    data_week_df = data_week_df.fillna(method='ffill')
    
    return data_dim,data_week_df

# 获取当周
def get_this_week(date_str):
    date_str = datetime.datetime.strptime(date_str,'%Y-%m-%d')
    this_week = str(date_str.year) + '-w' + str(date_str.strftime('%W'))
    return this_week

# 获取上周
def get_last_week(date_str):
    date_str = datetime.datetime.strptime(date_str,'%Y-%m-%d')
    last_week_begin_date = date_str-datetime.timedelta(days=date_str.weekday()+7)
    last_week = str(last_week_begin_date.year) + '-w' + str(last_week_begin_date.strftime('%W'))
    return last_week

# 获取年初第一周
def get_year_begin_week(date_str):
    date_str = datetime.datetime.strptime(date_str,'%Y-%m-%d') 
    year_begin_week = str(date_str.year) + '-w01'
    return year_begin_week


# 获取某指标及单位
def get_indicator_info(config_data,data_dim,first_level_para_id,second_level_para_id,third_level_para_id,indicator_id):
    para = config_data[(config_data['first_level_para_id'] == first_level_para_id) & (config_data['second_level_para_id'] == second_level_para_id) & (config_data['third_level_para_id'] == third_level_para_id)]
    indicator_name = para[para['indicator_id'] == indicator_id]['indicator_name'].tolist()[0]
    indicator_brief_name = para[para['indicator_id'] == indicator_id]['indicator_brief_name'].tolist()[0]
    unit = data_dim[indicator_name].loc[2]
    if isinstance(unit,str) == False:
        unit = "个点"
    return indicator_name,indicator_brief_name,unit

# 获取某周某指标数据
def get_num_index_one_time(config_data,data_dim,date_str,data,calc_baseline,first_level_para_id,second_level_para_id,third_level_para_id,indicator_id):
    indicator_name,indicator_brief_name,unit = get_indicator_info(config_data,data_dim,first_level_para_id,second_level_para_id,third_level_para_id,indicator_id)
    if calc_baseline == '当周':
        week = get_this_week(date_str)
    if calc_baseline == '上周':
        week = get_last_week(date_str)
    if calc_baseline == '年初':
        week = get_year_begin_week(date_str)
    res = data.loc[week,indicator_name]
    return res

# 获取某指标文本
def get_text_lvl_index_one_time(config_data,data_dim,date_str,data,calc_baseline,first_level_para_id,second_level_para_id,third_level_para_id,indicator_id):
    indicator_name,indicator_brief_name,unit = get_indicator_info(config_data,data_dim,first_level_para_id,second_level_para_id,third_level_para_id,indicator_id)
    num = get_num_index_one_time(config_data,data_dim,date_str,data,calc_baseline,first_level_para_id,second_level_para_id,third_level_para_id,indicator_id)
    
    if math.isnan(num) == True:
        round_num = '  '
    
    if math.isnan(num) == False:
        if len(str(int(num))) >=4:
            round_num = str(abs(int(num)))
        if len(str(int(num))) <4:
            round_num = str(abs(round(num,1)))

    if isinstance(unit,str) == True:
        text = round_num + unit
    return text

# 获取比较计算结果
def get_num_compare_index_two_time(config_data,data_dim,date_str,data,calc_baseline,calc_type,first_level_para_id,second_level_para_id,third_level_para_id,indicator_id ):
    num1 = get_num_index_one_time(config_data,data_dim,date_str,data,'当周',first_level_para_id,second_level_para_id,third_level_para_id,indicator_id)
    num2 = get_num_index_one_time(config_data,data_dim,date_str,data,calc_baseline,first_level_para_id,second_level_para_id,third_level_para_id,indicator_id)
    
    if calc_type == "pct": #计算变化的百分比
        res = num1 / num2 * 100 - 100
    if calc_type == "change": #计算数值变化
        res = num1 - num2
    return res

# 获取比较计算文本
def get_text_compare_index_two_time(config_data,data_dim,date_str,data,calc_baseline,calc_type,first_level_para_id,second_level_para_id,third_level_para_id,indicator_id):
    indicator_name,indicator_brief_name,unit = get_indicator_info(config_data,data_dim,first_level_para_id,second_level_para_id,third_level_para_id,indicator_id)
    num = get_num_compare_index_two_time(config_data,data_dim,date_str,data,calc_baseline,calc_type,first_level_para_id,second_level_para_id,third_level_para_id,indicator_id)
    
    if math.isnan(num) == True:
        round_num = '  '
    
    if math.isnan(num) == False:
        if len(str(int(num))) >= 4:
            round_num = str(abs(round(num,0)))
        if len(str(int(num)) )<4:
            round_num = str(abs(round(num,1)))
    
    if calc_type == "pct":
        if num > 0.0001:
            text = "上升"+round_num+"%"
        elif num < -0.0001:
            text = "下降"+round_num+"%"
        else:
            text = "持平"
    if calc_type == "change":
        if num > 0.0001:
            text = "上涨"+round_num+unit
        elif num < -0.0001:
            text = "下跌"+round_num+unit
        else:
            text = "持平"        
    return text

# 获取某指标完整文本
def get_sentence_of_indicator(config_data,data_dim,date_str,data,calc_baseline,calc_type,first_level_para_id,second_level_para_id,third_level_para_id,indicator_id):
    indicator_name,indicator_brief_name,unit = get_indicator_info(config_data,data_dim,first_level_para_id,second_level_para_id,third_level_para_id,indicator_id)
    text1 = get_text_lvl_index_one_time(config_data,data_dim,date_str,data,'当周',first_level_para_id,second_level_para_id,third_level_para_id,indicator_id)
    calc_baseline = calc_baseline.tolist()
    calc_type = calc_type.tolist()
    if len(calc_baseline) == 1:
        text2 = get_text_compare_index_two_time(config_data,data_dim,date_str,data,calc_baseline.tolist[0],calc_type[0],first_level_para_id,second_level_para_id,third_level_para_id,indicator_id)
        text = "本周"+indicator_brief_name+"为"+text1+"，较"+ calc_baseline + text2+"。"
    if len(calc_baseline) == 2:
        text2 = get_text_compare_index_two_time(config_data,data_dim,date_str,data,calc_baseline[0],calc_type[0],first_level_para_id,second_level_para_id,third_level_para_id,indicator_id)
        text3 = get_text_compare_index_two_time(config_data,data_dim,date_str,data,calc_baseline[1],calc_type[1],first_level_para_id,second_level_para_id,third_level_para_id,indicator_id)
        text = "本周"+indicator_brief_name+"为"+text1+"，较"+ calc_baseline[0] + text2+"，较" + calc_baseline[1]+text3+"。"
    return text

# 获取某段落完整文本
def get_paragraph(config_data,data_dim,date_str,data,first_level_para_id,second_level_para_id):
    para = config_data[(config_data['first_level_para_id'] == first_level_para_id) & (config_data['second_level_para_id'] == second_level_para_id)]
    third_level_para_id_list = list(dict.fromkeys(para['third_level_para_id'].tolist()))
    
    text = ''
    
    for i in range(0,len(third_level_para_id_list)):
        third_level_para_id = third_level_para_id_list[i]
        para = config_data[(config_data['first_level_para_id'] == first_level_para_id) & (config_data['second_level_para_id'] == second_level_para_id) & (config_data['third_level_para_id'] == third_level_para_id)]
        third_level_para_name = para['third_level_para_name'].tolist()[0]
        indicator_id_list = list(dict.fromkeys(para['indicator_id'].tolist()))
        
        text = text + third_level_para_name + "方面，"
        for j in range(0,len(indicator_id_list)):
            indicator_id = indicator_id_list[j]
            calc_baseline = para[para['indicator_id'] == indicator_id]['calc_baseline']
            calc_type = para[para['indicator_id'] == indicator_id]['calc_type']
            text = text + get_sentence_of_indicator(config_data,data_dim,date_str,data,calc_baseline,calc_type,first_level_para_id,second_level_para_id,third_level_para_id_list[i],indicator_id)
        text = text + "\n"
    
    text = text.rstrip()        
    return text


