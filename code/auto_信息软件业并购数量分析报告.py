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
from matplotlib.font_manager import FontProperties

from matplotlib.font_manager import findfont, FontProperties

def f_generate_horizontal_bar_chart(words,frequencies,output_path):

    # words=words.reverse()
    # frequencies=frequencies.reverse()
    non_zero_words = [word for word, freq in zip(words, frequencies) if freq != 0]
    non_zero_freqs = [freq for freq in frequencies if freq != 0]
        # 反转两个列表
    non_zero_words_reversed = non_zero_words[::-1]
    non_zero_freqs_reversed = non_zero_freqs[::-1]
    # 绘制条形图
    plt.figure(figsize=(10, 6))  # 设置图形大小
    # plt.barh(words, frequencies, color='#4472C4')  # 创建条形图
    plt.barh(non_zero_words_reversed, non_zero_freqs_reversed , color='#4472C4')  # 创建条形图
    plt.legend(['并购交易金额（万元）'], loc='lower center',bbox_to_anchor=(0.5, -0.2), ncol=1,frameon=False,fontsize=18)  # 图例文本和位置
    for i, value in enumerate(non_zero_freqs_reversed):
        if(value>0):
            plt.text(value, i, f"{int(value)}", ha='left', va='center',fontsize=16)
    
    plt.yticks(fontsize=18)
    plt.xticks(fontsize=18)
    max_value = max(non_zero_freqs_reversed)
    plt.xlim(0, max_value * 1.2)
    plt.tight_layout()  # 自动调整布局，避免重叠
    plt.savefig(output_path)


def n_generate_horizontal_bar_chart(words,frequencies,output_path):

    # words=words.reverse()
    # frequencies=frequencies.reverse()
    non_zero_words = [word for word, freq in zip(words, frequencies) if freq != 0]
    non_zero_freqs = [freq for freq in frequencies if freq != 0]
        # 反转两个列表
    non_zero_words_reversed = non_zero_words[::-1]
    non_zero_freqs_reversed = non_zero_freqs[::-1]
    # 绘制条形图
    plt.figure(figsize=(10, 6))  # 设置图形大小
    # plt.barh(words, frequencies, color='#4472C4')  # 创建条形图
    plt.barh(non_zero_words_reversed, non_zero_freqs_reversed , color='#4472C4')  # 创建条形图
    plt.legend(['并购交易数量'], loc='lower center',bbox_to_anchor=(0.5, -0.2), ncol=1,frameon=False,fontsize=18)  # 图例文本和位置
    for i, value in enumerate(non_zero_freqs_reversed):
        if(value>0):
            plt.text(value, i, f"{int(value)}", ha='left', va='center',fontsize=16)
    
    plt.yticks(fontsize=18)
    plt.xticks(fontsize=18)
    max_value = max(non_zero_freqs_reversed)
    plt.xlim(0, max_value * 1.2)
    plt.tight_layout()  # 自动调整布局，避免重叠
    plt.savefig(output_path)
    

def generate_report_sentence(date,current_year_finance, year_on_year_growth):
    # print(current_year_finance,year_on_year_growth)
    month=date[4:6]
    month=int(month)
    month=str(month)
    time_year= f"{date[:4]}年1-{month}月"
    sentence=time_year
    sentence += f"，北京市信息软件业并购金额为{current_year_finance:.1f}万元，"
    sentence += f"同比{'下降' if year_on_year_growth < 0 else '增长'}{abs(year_on_year_growth):.1f}%。"
    return sentence

def generate_report(date, data_path, pic_par1_path,pic_par2_path):
    month=date[4:6]
    month=int(month)
    month=str(month)
    time_year= f"（{date[:4]}年{month}月）"
    # time_year= f"（{date[:4]}年{date[4:6]}月）"

    df = pd.read_excel(data_path, header=0, sheet_name='投融资')
    df = df[df['报告期编码'].astype(str) == date]

    total_value=df[(df['指标名称'] == '并购交易金额_累计值')]['统计值'].iloc[0]
    total_rate=df[(df['指标名称'] == '并购交易金额_累计增速')]['统计值'].iloc[0]

    fd_df = df[(df['指标名称'] == '并购交易金额_累计值_分领域')]
    nd_df = df[(df['指标名称'] == '并购案例数量_累计值_分领域')]
    fd_df = fd_df.sort_values(by='统计值', ascending=False)
    nd_df = nd_df.sort_values(by='统计值', ascending=False)
    fd_top_names = list(fd_df['拓展维度'])
    nd_top_names = list(nd_df['拓展维度'])

    
    # 按照顺序获取并购交易金额
    finance_values = []
    for name in fd_top_names:
        finance_value = int(fd_df[fd_df['拓展维度'] == name]['统计值'])
        finance_values.append(finance_value)
    
    num_values=[]
    for name in nd_top_names:
        num_value=int(nd_df[nd_df['拓展维度'] == name]['统计值'])
        num_values.append(num_value)
    
    print(num_values)

    
    first_par=generate_report_sentence(date,total_value,total_rate)
    n_generate_horizontal_bar_chart(nd_top_names,num_values,pic_par1_path)
    f_generate_horizontal_bar_chart(fd_top_names,finance_values,pic_par2_path)

    second_par = f"分领域看，1-{month}月，北京市信息软件业并购案例数前五大领域分别为"
    second_par += "、".join([f"{name}" for name in nd_top_names[:5]])
    second_par+= f"，案例数量分别为{num_values[0]}、{num_values[1]}、{num_values[2]}、{num_values[3]}、{num_values[4]}个；"
    # 构建句子
    second_par += f"并购交易金额前五大领域分别为"
    second_par += "、".join([f"{name}" for name in fd_top_names[:5]])
    second_par += f"并购交易金额分别为{finance_values[0]}、{finance_values[1]}、{finance_values[2]}、{finance_values[3]}、{finance_values[4]}万元。"

    
    
    word_generator = WordReportGenerator()

    # 使用类方法创建文件
    # job_salary_docx_file(self,title0,pic_title,par1,file_pre,file_dir,pic_dir,year_month)
    word_generator.bingo_docx_file(
        # 必须
                            file_pre='信息软件业并购分析报告',
                            file_dir='test',
                            year_month=time_year,
        #开头 第一段
                            title0='信息软件业并购分析报告',
                            par1=first_par,

        #第二段
                            par2=second_par,
                            pic_dir1=pic_par1_path,
                            lt1='图：北京市信息软件业并购案例数量',
                            pic_dir2=pic_par2_path,
                            lt2='图：北京市信息软件业并购交易金额',
        #第三段
                             )
    



def main():
    plt.rcParams["font.sans-serif"] = ["STFangsong"]  # 设置字体
    data_path = '/Users/harvin/code/自动报告产品开发-产业链@20220830/data/input/指标数据20231124-v1.xlsx'
    pic_par1_path = '/Users/harvin/code/自动报告产品开发-产业链@20220830/data/chart_并购条形图1_2022_nov.png'
    pic_par2_path='/Users/harvin/code/自动报告产品开发-产业链@20220830/data/chart_并购条形图2_2022_nov.png'
    start_m=3
    end_year,end_m=2023,10
    end_m+=1
    
    for i in tqdm(range(start_m,end_m)):
        if i<10 and len(str(i))<2:
            m='0'+str(i)
        else:
            m=str(i)
        t=str(end_year)+m
        generate_report(t,data_path,pic_par1_path,pic_par2_path)

if __name__ == "__main__":
    # 执行 main 函数
    main()