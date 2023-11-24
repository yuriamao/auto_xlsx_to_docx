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
from matplotlib.font_manager import FontProperties

from matplotlib.font_manager import findfont, FontProperties
findfont(FontProperties(family=FontProperties().get_family()))




# plt.rcParams['font.sans-serif'] = ['Arial Unicode MS']  # 设置一个支持中文的字体，比如 Arial Unicode MS
# 设置中文显示

def generate_horizontal_bar_chart(date, data_path,output_path):
    # plt.rcParams['font.sans-serif'] = ['Arial Unicode MS']

    job_df = pd.read_excel(data_path, header=0, sheet_name='招聘')

    selected_data = job_df[(job_df['指标名称'] == '平均最高薪资区分集团系')]
    selected_data.loc[:, '统计值'] = selected_data['统计值'].astype(float)
    aggregated_df = selected_data.groupby(['集团系名称', '报告期编码']).agg({'统计值': 'first'}).reset_index()
    selected_data = aggregated_df[aggregated_df['报告期编码'].astype(str).str.contains(date)]

    ordered_group_names = ['小米系', '抖音系', '神州数码（北京神码和神码中国）', '百度系', '京东系', '美团系', '快手系']

    group_values = []  # 用于存储各集团系的薪资数据
    for group_name in ordered_group_names:
        group_data = selected_data[selected_data['集团系名称'] == group_name]


        if not group_data.empty:
            salary_value = group_data['统计值'].iloc[0]
            group_values.append(salary_value)

    if group_values:
        plt.figure(figsize=(10, 6))
        # 定义所需的中文字体
        # custom_font = FontProperties(fname='/System/Library/Fonts/STFangsong.ttf')
        # font = FontProperties(fname = "/Users/harvin/code/自动报告产品开发-产业链@20220830/data/华文仿宋.ttf")
        # plt.rcParams['font.sans-serif'] = [font.get_name()]
        plt.rcParams['axes.unicode_minus'] = False


    # 创建字体属性对象
        # font_prop = FontProperties(fname=font_path)
        # plt.subplots_adjust(top=0.9,bottom=0.1,left=0.2,right=1.95)
        # 设置 x 轴范围
        max_value = max(group_values)
        plt.xlim(0, max_value + 10000)  # 设置 x 轴范围，确保最大刻度比最大值大 1w
         # 添加图例，调整位置并设置字号
        # plt.legend(loc='lower center', bbox_to_anchor=(0.5, -0.15), ncol=1, fontsize=12)


        plt.rcParams['axes.unicode_minus'] = False



    # 绘制条形图
        plt.barh(ordered_group_names, group_values, color = '#4472C4')
        # 添加图例
        plt.legend(['平均最高薪资（元）'], loc='lower center',bbox_to_anchor=(0.5, -0.15), ncol=1,frameon=False,fontsize=15)  # 图例文本和位置

        # 添加图例
        # plt.legend(loc='lower center', bbox_to_anchor=(0.5, -0.15), ncol=1,frameon=False,fontsize=12, title='平均最高薪资（元）')
                # 在每个条后面显示数值
        for i, value in enumerate(group_values):
            plt.text(value, i, f"{value:.2f}", ha='left', va='center',fontsize=15)
            plt.tight_layout()
            
        # plt.xlabel('平均最高薪资（元）')
        # plt.title(f'信息软件业招聘岗位{date[:4]}年{date[4:6]}月不同集团系平均最高薪资对比')
        plt.yticks(fontsize=15)
        plt.xticks(fontsize=15)

        plt.tight_layout()
        
        plt.savefig(output_path)

def generate_report(date, data_path, output_image_path):
    job_df = pd.read_excel(data_path, header=0, sheet_name='招聘')

    selected_data = job_df[(job_df['指标名称'] == '平均最高薪资区分集团系')]
    selected_data.loc[:, '统计值'] = selected_data['统计值'].astype(float)  # 更改为浮点数类型
    aggregated_df = selected_data.groupby(['集团系名称', '报告期编码']).agg({'统计值': 'first'}).reset_index()
    
    # 获取当前月份的数据
    current_data = aggregated_df[aggregated_df['报告期编码'].astype(str).str.contains(date)]

    # 计算上个月的报告期编码
    year = int(date[:4])
    month = int(date[4:])
    if month == 1:
        previous_date = f"{year - 1}12"
    else:
        previous_date = f"{year}{month - 1:02d}"

    # 获取上个月的数据
    previous_data = aggregated_df[aggregated_df['报告期编码'].astype(str).str.contains(previous_date)]

    # 如果找到了上个月和当前月的数据，计算涨跌百分比
    if not current_data.empty and not previous_data.empty:
        current_value = current_data['统计值'].mean()
        previous_value = previous_data['统计值'].mean()
        
        # 计算涨跌百分比
        percentage_change = ((current_value - previous_value) / previous_value) * 100

        # 生成报告文本
        month=date[4:6]
        month=int(month)
        month=str(month)
        output_text = f"{date[:4]}年{month}月，北京市信息软件业平均最高薪资为{current_value:.2f}元，较上月{'上涨' if percentage_change >= 0 else '下跌'}{abs(percentage_change):.1f}%。其中，"
        ordered_group_names = ['小米系', '抖音系', '神州数码（北京神码和神码中国）', '百度系', '京东系', '美团系', '快手系']
        for group_name in ordered_group_names:
            group_data = current_data[current_data['集团系名称'] == group_name]
            if not group_data.empty:
                salary_value = group_data['统计值'].iloc[0]
                output_text += f"{group_name}平均最高薪资为{salary_value:.2f}元，"

        output_text = output_text[:-1] + "。"


        month=date[4:6]
        month=int(month)
        month=str(month)
        time_year= f"（{date[:4]}年{month}月）"


    # 实例化 WordReportGenerator 类
    word_generator = WordReportGenerator()

    # 使用类方法创建文件
    # job_salary_docx_file(self,title0,pic_title,par1,file_pre,file_dir,pic_dir,year_month)
    word_generator.job_salary_docx_file(title0='信息软件业平均最高薪资分析报告',
                             par1=output_text,
                             pic_title='图：北京市信息软件业平均最高薪资',
                             file_pre='信息软件业平均最高薪资分析报告',
                             pic_dir=output_image_path,
                             file_dir='test',
                             year_month=time_year
                             )


    # 添加其他内容到文档...

    # 保存 Word 文档
    # doc.save(f'/Users/harvin/code/自动报告产品开发-产业链@20220830/data/output/招聘统计信息_{date}.docx')



def main():
    import matplotlib as mpl
    print(mpl.get_cachedir())
    plt.rcParams["font.sans-serif"] = ["STFangsong"]  # 设置字体
    data_path = '/Users/harvin/code/自动报告产品开发-产业链@20220830/data/input/指标数据20231124-v1.xlsx'
    output_image_path = '/Users/harvin/code/自动报告产品开发-产业链@20220830/data/chart_most_salary_2022_nov.png'
    start_m=3
    end_year,end_m=2023,10
    end_m+=1
    
    for i in tqdm(range(start_m,end_m)):
        if i<10 and len(str(i))<2:
            m='0'+str(i)
        else:
            m=str(i)
        t=str(end_year)+m
        generate_horizontal_bar_chart(t, data_path,output_path=output_image_path)
        generate_report(t,data_path,output_image_path)

if __name__ == "__main__":
    # 执行 main 函数
    main()