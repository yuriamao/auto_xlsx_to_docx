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


def top_n_jobs(data, company_name, top=3):
    # 假设数据中有 '集团系名称' 和 '岗位名称' 以及 '统计值' 列
    selected_data = data[data['集团系名称'] == company_name]
    selected_data['统计值'] = selected_data['统计值'].astype(int)
    selected_data['岗位名称']=selected_data['岗位名称'].astype(str)
    # 按岗位名称分组并对统计值求和
    aggregated_data = selected_data.groupby('岗位名称')['统计值'].sum().reset_index()
    # 按统计值降序排序
    aggregated_data = aggregated_data.sort_values(by='统计值', ascending=False)
    
    # 获取前 top 个岗位和对应的统计值
    top_jobs = aggregated_data.head(top)
    top_jobs.reset_index(drop=True, inplace=True)  # 重置索引，确保索引从零开始
    # 构建返回的句子
    result = f"{company_name}前三大热门招聘岗位分别为"
    for i in range(top):
        result += f"{top_jobs['岗位名称'][i]}"
        if i != top - 1:
            result += "、"
    
    result += "，招聘人数分别为"
    for i in range(top):
        result += f"{top_jobs['统计值'][i]}人"
        if i != top - 1:
            result += "、"
    
    result += "。"
    # 生成词云图
    # wordcloud_text = ' '.join(aggregated_data['岗位名称'].astype(str))  # 将岗位名称转换为字符串再组合为文本
    top_jobs = aggregated_data['岗位名称'].tolist()  # 获取 '岗位名称' 列的前三个值并转换为列表
    top_values = aggregated_data['统计值'].tolist() # 获取 '统计值' 列的前三个值并转换为列表
       # 将领域名称和融资金额转换为词云所需的格式
    MAX_WORDS = 50  # 想要限制的最大词汇数量

    # 如果 top_jobs 和 top_values 的长度超过了 MAX_WORDS，则只保留前 MAX_WORDS 个元素
    # word_freq = {name: value for name, value in zip(top_jobs[:MAX_WORDS], top_values[:MAX_WORDS])}
    word_freq = {name: value for name, value in zip(top_jobs, top_values)}
     # 生成词云图
    font=r'/Users/harvin/Library/Fonts/华文仿宋.ttf'

    wordcloud = WordCloud(width=2000,height=1200,font_path=font, background_color='white').generate_from_frequencies(word_freq)
    # 显示词云图
    plt.figure(figsize=(10, 6))
    plt.imshow(wordcloud, interpolation='bilinear')
    plt.axis('off')
    # plt.title(f"{company_name}系热门岗位词云图")
    file_name=f'{company_name}_词云图.png'
    file_path = f'/Users/harvin/code/自动报告产品开发-产业链@20220830/data/output/信息软件业热门岗位分析报告/{file_name}'
    os.makedirs(os.path.dirname(file_path), exist_ok=True)
    pic_dir=file_path
    plt.savefig(pic_dir,dpi=200)
    # 构建返回的句子
    little = f"图：{company_name}热门岗位词云图"
    return result,pic_dir,little



def generate_report(date, data_path):
    year = date[:4]
    month=date[4:6]
    month=int(month)
    month=str(month)
    time_year= f"（{date[:4]}年{month}月）"

    job_df = pd.read_excel(data_path, header=0, sheet_name='招聘')
    selected_data = job_df[(job_df['指标名称'] == '热门岗位频率')]
    selected_data = selected_data[selected_data['报告期编码'].astype(str) == date]
    selected_data['岗位名称']=selected_data['岗位名称'].astype(str)
    selected_data['统计值'] = selected_data['统计值'].astype(int)
    job_type = selected_data[selected_data['集团系名称'].notnull()]
    job_type=job_type.groupby(['岗位名称'], as_index=False).agg({'统计值':'sum'})

    # 假设筛选后的数据按统计值统计，并且该数据有 '岗位名称' 和 '统计值' 两列
    selected_data = selected_data.sort_values(by='统计值', ascending=False)  # 按统计值降序排序

    job_type = job_type.sort_values(by='统计值', ascending=False)  # 按统计值降序排序


    company_list=['美团系','快手系','京东系','抖音系','神州数码（北京神码和神码中国）','小米系','百度系']


    # 假设 job_type 包含 '岗位名称' 和 '统计值' 列
    top_jobs = job_type['岗位名称'].tolist()  # 获取 '岗位名称' 列的前三个值并转换为列表
    top_values = job_type['统计值'].tolist() # 获取 '统计值' 列的前三个值并转换为列表


       # 将领域名称和融资金额转换为词云所需的格式
    word_freq = {name: value for name, value in zip(top_jobs, top_values)}

     # 生成词云图
    font=r'/Users/harvin/Library/Fonts/华文仿宋.ttf'

    wordcloud = WordCloud(width=2000,height=1200,font_path=font, background_color='white').generate_from_frequencies(word_freq)
    # 显示词云图
    plt.figure(figsize=(10, 6))
    plt.imshow(wordcloud, interpolation='bilinear')
    plt.axis('off')
    # plt.title(f"{company_name}系热门岗位词云图")
    file_name=f'总体_词云图.png'
    file_path = f'/Users/harvin/code/自动报告产品开发-产业链@20220830/data/output/信息软件业热门岗位分析报告/{file_name}'
    os.makedirs(os.path.dirname(file_path), exist_ok=True)
    pic_dir=file_path
    plt.savefig(pic_dir,dpi=200)
    # 构建返回的句子
    little = f"图：所有集团系热门岗位词云图"


    word_generator = WordReportGenerator()
    first_par = f"{date[:4]}年{month}月,北京市信息软件业前三大热门招聘岗位分别为{top_jobs[0]}、{top_jobs[1]}、{top_jobs[2]}，招聘人数分别为{top_values[0]}人、{top_values[1]}人、{top_values[2]}人。"
    txt_list=[first_par]
    pic_dir_list=[pic_dir]
    little_list=[little]
    for company in company_list:
        txt,pic_dir,litte=top_n_jobs(selected_data,company,top=3)
        txt_list.append(txt)
        pic_dir_list.append(pic_dir)
        little_list.append(litte)
    
    word_generator.top_job_docx_file(
    # 必须
                        file_pre='信息软件业热门岗位分析报告',
                        file_dir='test',
                        year_month=time_year,
    #开头 第一段
                        title0='信息软件业热门岗位分析报告',
                        
                        # par1=first_par,
    # 画各公司图
                        txt_par_list=txt_list,
                        pic_dir_list=pic_dir_list,
                        little_list=little_list
                            )
    
def main():
    plt.rcParams["font.sans-serif"] = ["STFangsong"]  # 设置字体
    data_path = '/Users/harvin/code/自动报告产品开发-产业链@20220830/data/input/指标数据20231124-v1.xlsx'
    start_m=7
    end_year,end_m=2023,10
    end_m+=1
    
    for i in tqdm(range(start_m,end_m)):
        if i<10 and len(str(i))<2:
            m='0'+str(i)
        else:
            m=str(i)
        t=str(end_year)+m
        # generate_horizontal_bar_chart(t, data_path,output_path=output_image_path)
        generate_report(t,data_path)

if __name__ == "__main__":
    # 执行 main 函数
    main()