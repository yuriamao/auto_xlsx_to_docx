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

def generate_horizontal_bar_chart(words,frequencies,output_path):
    # 绘制条形图
    plt.figure(figsize=(10, 6))  # 设置图形大小
    plt.barh(words, frequencies, color='#4472C4')  # 创建条形图
    plt.legend(['融资金额（万元）'], loc='lower center',bbox_to_anchor=(0.5, -0.2), ncol=1,frameon=False,fontsize=18)  # 图例文本和位置
    for i, value in enumerate(frequencies):
        plt.text(value, i, f"{int(value)}", ha='left', va='center',fontsize=16)
        plt.tight_layout()
        
    plt.yticks(fontsize=18)
    plt.xticks(fontsize=18)
    max_value = max(frequencies)

    plt.xlim(0, max_value * 1.2)

    # plt.xlabel('融资轮次')  # x 轴标签
    # plt.ylabel('融资金额（万元）')  # y 轴标签
    # plt.title('融资金额分布')  # 图表标题
    # plt.xticks(rotation=45)  # 旋转 x 轴标签，使其更易读

    plt.tight_layout()  # 自动调整布局，避免重叠
    # plt.show()  # 显示图形
    plt.savefig(output_path)
    

def generate_filed_bar_chart(words,frequencies,output_path):
    words = list(reversed(words))
    frequencies = list(reversed(frequencies))
    # 绘制条形图
    plt.figure(figsize=(10, 6))  # 设置图形大小
    plt.barh(words, frequencies, color='#4472C4')  # 创建条形图
    plt.legend(['融资金额（万元）'], loc='lower center',bbox_to_anchor=(0.5, -0.2), ncol=1,frameon=False,fontsize=18)  # 图例文本和位置
    for i, value in enumerate(frequencies):
        plt.text(value, i, f"{int(value)}", ha='left', va='center',fontsize=16)
        plt.tight_layout()
        
    plt.yticks(fontsize=18)
    plt.xticks(fontsize=18)
    max_value = max(frequencies)

    plt.xlim(0, max_value * 1.2)

    # plt.xlabel('融资轮次')  # x 轴标签
    # plt.ylabel('融资金额（万元）')  # y 轴标签
    # plt.title('融资金额分布')  # 图表标题
    # plt.xticks(rotation=45)  # 旋转 x 轴标签，使其更易读

    plt.tight_layout()  # 自动调整布局，避免重叠
    # plt.show()  # 显示图形
    plt.savefig(output_path)
    

def generate_report_sentence(date,current_year_finance, year_on_year_growth):
    # print(current_year_finance,year_on_year_growth)
    month=date[4:6]
    month=int(month)
    month=str(month)
    time_year= f"{date[:4]}年1-{month}月"
    sentence=time_year
    sentence += f"，北京市信息软件业融资金额为{current_year_finance:.1f}万元，"
    sentence += f"同比{'下降' if year_on_year_growth < 0 else '增长'}{abs(year_on_year_growth):.1f}%。"
    return sentence

def generate_report(date, data_path, output_image_path,pic_par2_path,pic_par3_path):
    month=date[4:6]
    month=int(month)
    month=str(month)
    time_year= f"（{date[:4]}年{month}月）"
    # time_year= f"（{date[:4]}年{date[4:6]}月）"

    job_df = pd.read_excel(data_path, header=0, sheet_name='投融资')

    selected_data = job_df[(job_df['指标名称'] == '融资金额_累计值_分轮次')]
    current_df=job_df[(job_df['指标名称'] == '融资金额_累计增速')]
    selected__cur_data = current_df[current_df['报告期编码'].astype(str).str.contains(date)]
    c_rate=selected__cur_data['统计值'].iloc[0]
    selected_data = selected_data[selected_data['报告期编码'].astype(str).str.contains(date)]
    # print(selected_data.head())

    # 根据数据生成段落
    output_text = f"分轮次看，{date[:4]}年{month}月，北京市"
    rounds = {
        'A轮': 'A轮融资金额',
        'B轮': 'B轮融资金额',
        'C轮': 'C轮融资金额',
        'D轮': 'D轮融资金额',
        'F轮': 'F轮融资金额',
        'IPO上市后': 'IPO上市后融资金额',
        'Pre-IPO': 'Pre-IPO融资金额',
        '天使轮': '天使轮融资金额',
        '种子轮': '种子轮融资金额'
    }

    total_v=0
    # 初始化词频字典
    word_freq = {}
    for round_name, column_name in rounds.items():
        # print(round_name)
        round_data = selected_data[selected_data['拓展维度'] == round_name]
        # print(round_data)
        if not round_data.empty:
            financing_value = float(round_data['统计值'].iloc[0])
            total_v+=financing_value
            output_text += f"{round_name}融资金额为{financing_value:.2f}万元，"
            word_freq[round_name] = financing_value

    # 对词频字典按值（频数）进行排序
    sorted_word_freq = dict(sorted(word_freq.items(), key=lambda item: item[1], reverse=False))

    # 提取经排序后的词汇和频率
    words = list(sorted_word_freq.keys())
    frequencies = list(sorted_word_freq.values())

    generate_horizontal_bar_chart(words,frequencies,output_path=pic_par2_path)


    current_year_finance=total_v
    year_on_year_growth=c_rate
    first_par=generate_report_sentence(date, float(current_year_finance), float(year_on_year_growth))
    output_text = output_text[:-1] + "。"
    second_par=output_text

    third_par = f"分领域看，1-{month}月，北京市信息软件业融资金额前五大领域分别为"

    d_selected_data = job_df[(job_df['指标名称'] == '融资金额_累计值_分领域')]
    selected__d_data = d_selected_data[d_selected_data['报告期编码'].astype(str).str.contains(date)]

    sorted_d_data = selected__d_data.sort_values(by='统计值', ascending=False)
    names_list = list(sorted_d_data['拓展维度'])





    top_names = names_list

    # 按照顺序获取融资金额
    finance_values = []
    for name in top_names:
        finance_value = sorted_d_data[sorted_d_data['拓展维度'] == name]['统计值'].values[0]
        finance_values.append(finance_value)

        # 获取前 10 个最大值的索引
    top_indices = sorted(range(len(finance_values)), key=lambda i: finance_values[i], reverse=True)[:10]

    # 仅保留前 10 个最大值的数据
    top_10_names = [top_names[i] for i in top_indices]
    top_10_finance_values = [finance_values[i] for i in top_indices]

    generate_filed_bar_chart(top_10_names,top_10_finance_values,output_path=pic_par3_path)

    # 将领域名称和融资金额转换为词云所需的格式
    word_freq = {name: value for name, value in zip(top_names, finance_values)}

    # 构建句子
    third_par += "、".join([f"{name}" for name in top_names[:5]])
    third_par += f"，融资金额分别为{finance_values[0]}、{finance_values[1]}、{finance_values[2]}、{finance_values[3]}、{finance_values[4]}万元。"

    # 生成词云图

    font=r'/Users/harvin/Library/Fonts/华文仿宋.ttf'
    # 生成词云图
    wordcloud = WordCloud(width=2000,height=1200,font_path=font, background_color='white').generate_from_frequencies(word_freq)

    # 显示词云图
    plt.figure(figsize=(10, 6))
    # custom_font = FontProperties(fname='/System/Library/Fonts/STFangsong.ttf')
    
    plt.imshow(wordcloud, interpolation='bilinear')
    plt.axis('off')
    plt.savefig(output_image_path,dpi=200)  # 保存图形

    # plt.show()

    word_generator = WordReportGenerator()

    # 使用类方法创建文件
    # job_salary_docx_file(self,title0,pic_title,par1,file_pre,file_dir,pic_dir,year_month)
    word_generator.ie_fina_docx_file(
        # 必须
                            file_pre='信息软件业投融资分析报告_231125_01',
                            file_dir='test',
                            year_month=time_year,
        #开头 第一段
                            title0='信息软件业投融资分析报告',
                            par1=first_par,


        #第二段
                            par2=second_par,
                            pic_dir2=pic_par2_path,
                            lt2='图：分轮次信息软件业融资金额',
        #第三段
                            par3=third_par,
                            pic_dir=output_image_path,
                            lt='图：信息软件业融资领域词云图',
                            pic_dir3=pic_par3_path,
                            lt3='图：信息软件业融资金额前十大领域',
                             )
    



def main():
    plt.rcParams["font.sans-serif"] = ["STFangsong"]  # 设置字体
    data_path = '/Users/harvin/code/自动报告产品开发-产业链@20220830/data/input/指标数据20231124-v1.xlsx'
    output_image_path = '/Users/harvin/code/自动报告产品开发-产业链@20220830/data/chart_融资词云图_2022_nov.png'
    pic_par2_path='/Users/harvin/code/自动报告产品开发-产业链@20220830/data/chart_融资条形图_2022_nov.png'
    pic_par3_path='/Users/harvin/code/自动报告产品开发-产业链@20220830/data/chart_融资十大领域条形图_2022_nov.png'
    start_m=1
    end_year,end_m=2023,10
    end_m+=1
    
    for i in tqdm(range(start_m,end_m)):
        if i<10 and len(str(i))<2:
            m='0'+str(i)
        else:
            m=str(i)
        t=str(end_year)+m
        # generate_horizontal_bar_chart(t, data_path,output_path=output_image_path)
        generate_report(t,data_path,output_image_path,pic_par2_path,pic_par3_path)

if __name__ == "__main__":
    # 执行 main 函数
    main()