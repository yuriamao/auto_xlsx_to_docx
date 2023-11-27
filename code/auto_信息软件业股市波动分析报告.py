import pandas as pd
from fuzzywuzzy import process
from docx import Document
from docx.shared import Pt
import matplotlib.pyplot as plt
from tqdm import tqdm
from utils import *
from io import BytesIO
import matplotlib.dates as mdates
import matplotlib.patches as mpatches
from wordcloud import WordCloud
from matplotlib.font_manager import FontProperties
import numpy as np
from matplotlib.font_manager import findfont, FontProperties
from datetime import datetime
# 创建一个椭圆形的图像


def generate_report(date, data_path, company_list):
    title_list=[]
    txt_list=[]
    pic_dir_list=[]
    little_list=[]
    # 将日期字符串转换为年月格式
    year = date[:4]
    month=date[4:6]
    month=int(month)
    month=str(month)
    time_year= f"（{year}年{month}月）"
    report_text = f"{year}年{month}月"

    # 转换日期为日期格式，便于后续处理
    date_formatted = pd.to_datetime(date, format='%Y%m')

    # 计算上个月的年份和月份
    last_month = date_formatted - pd.DateOffset(months=1)
    last_month_year = str(last_month.year)
    last_month_month = str(last_month.month).zfill(2)  # 格式化月份，保证是两位数

    # 上个月的起止日期
    start_of_last_month = f"{last_month_year}{last_month_month}"
    end_of_last_month = pd.to_datetime(start_of_last_month, format='%Y%m')
    end_of_last_month = end_of_last_month.strftime('%Y%m')


    # 读取数据
    selected_data = pd.read_excel(data_path, header=0, sheet_name='市值')

    # 计算当前月的平均市值及单位信息

    for company in company_list:
            # 获取上个月的市值信息（仅考虑年份和月份）


        selected_data_current_month = selected_data[
            (selected_data['日期'].astype(str).str[:6] == date[:6]) & 
            (selected_data['公司名称']==(company))
        ]
        last_month_values = selected_data[
            (selected_data['日期'].astype(str).str[:6] == f"{last_month_year}{last_month_month}") & 
            (selected_data['公司名称']==(company))
        ]


        average_current = selected_data_current_month.groupby('公司名称')['市值(亿)'].mean()[company]

    
        units = selected_data_current_month.groupby('公司名称')['单位'].first()


        average_last=last_month_values.groupby('公司名称')['市值(亿)'].mean()[company]

        print(average_last,average_current)

        max_value_current_month = selected_data_current_month.loc[selected_data_current_month['市值(亿)'].idxmax()]
        min_value_current_month = selected_data_current_month.loc[selected_data_current_month['市值(亿)'].idxmin()]


            # 计算月市值的变化
        diff_current = average_current-average_last
        day_diff_current = max_value_current_month ['市值(亿)'] - min_value_current_month['市值(亿)']


        report_text += f"{company}平均市值为{average_current:.1f}{units[company]}，"
        if diff_current > 0:
            report_text += f"较上月高{abs(diff_current):.1f}{units[company]}。"
        elif diff_current < 0:
            report_text += f"较上月低{abs(diff_current):.1f}{units[company]}。"
        else:
            report_text += "与上月持平。"

        # 假设 min_value_last_month 和 max_value_last_month 中 '日期' 列的数据格式是 %Y-%m-%d
        # 假设 min_value_last_month 和 max_value_last_month 中 '日期' 列的数据类型是整数
        min_date = str(min_value_current_month['日期'])  # 将整数转换为字符串
        min_date = datetime.strptime(min_date, '%Y%m%d').strftime('%m月%d日')  # 格式化日期

        max_date = str(max_value_current_month['日期'])  # 将整数转换为字符串
        max_date = datetime.strptime(max_date, '%Y%m%d').strftime('%m月%d日')  # 格式化日期


        report_text += f"市值最低为{min_value_current_month['市值(亿)']:.1f}{units[company]}（{min_date}），"
        report_text += f"最高为{max_value_current_month ['市值(亿)']:.1f}{units[company]}（{max_date}），"
        report_text += f"相差{day_diff_current:.1f}{units[company]}。"

        # print(report_text)
        txt_list.append(report_text)
        report_text=''

        # 画图
        plt.figure(figsize=(10, 6))

        plt.plot(selected_data_current_month['日期'].astype(str).str[:8],selected_data_current_month['市值(亿)'],label=f"市值（{units[company]}）", linewidth=4,)
        # 设置 x 轴刻度标签的旋转角度为45度
        # 获取最大值和最小值对应的数值及索引
        
        plt.xticks(rotation=45)
        plt.legend(loc='lower center', bbox_to_anchor=(0.5, -0.22), ncol=2, fontsize=12)
        # 调整图形布局
        plt.tight_layout()

        # 设置 x 轴以天为间隔
        file_name=f'{company}_股市条形图.png'
        file_path = f'/Users/harvin/code/自动报告产品开发-产业链@20220830/data/output/信息软件业热门岗位分析报告/{file_name}'
        os.makedirs(os.path.dirname(file_path), exist_ok=True)
        pic_dir=file_path
        plt.savefig(pic_dir,dpi=200)
        # 构建返回的句子
        little = f"图：{month}月{company}每日市值"
        little_list.append(little)
        pic_dir_list.append(pic_dir)
        title_list.append(company)
    add_list=['一、','二、','三、','四、','五、']
    for i,titie in enumerate(title_list):
        title_list[i]=add_list[i]+titie
    print(title_list,pic_dir_list)

    word_generator = WordReportGenerator()

    word_generator.fina_bodong_docx_file(
    # 必须
                        file_pre='信息软件业股市波动分析报告',
                        file_dir='test',
                        year_month=time_year,
    #开头 第一段
                        title0='信息软件业热门岗位分析报告',
                    
                        # par1=first_par,


    # 画各公司图          
                        title_list=title_list,
                        txt_par_list=txt_list,
                        pic_dir_list=pic_dir_list,
                        little_list=little_list
                            )
    return title_list,txt_list,pic_dir_list,little_list
                    

def main():
    plt.rcParams["font.sans-serif"] = ["STFangsong"]  # 设置字体
    data_path = '/Users/harvin/code/自动报告产品开发-产业链@20220830/data/input/指标数据20231124-v1.xlsx'
    start_m=3
    end_year,end_m=2023,10
    end_m+=1
    company_list = ['中国移动','美团点评','京东集团','百度集团','小米集团']
    
    for i in tqdm(range(start_m,end_m)):
        if i<10 and len(str(i))<2:
            m='0'+str(i)
        else:
            m=str(i)
        t=str(end_year)+m
        # generate_horizontal_bar_chart(t, data_path,output_path=output_image_path)
        generate_report(t,data_path,company_list)

if __name__ == "__main__":
    # 执行 main 函数
    main()