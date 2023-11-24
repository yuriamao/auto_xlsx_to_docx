import pandas as pd
from fuzzywuzzy import process
from docx import Document
from docx.shared import Pt
import matplotlib.pyplot as plt
from tqdm import tqdm
from utils import *
from io import BytesIO
import matplotlib.dates as mdates

plt.rcParams['font.sans-serif'] = ['Arial Unicode MS']  # 设置一个支持中文的字体，比如 Arial Unicode MS

def longest_common_substring(s1, s2):
    # 基于最大重复子串匹配
    m = [[0] * (1 + len(s2)) for _ in range(1 + len(s1))]
    longest, x_longest = 0, 0
    for x in range(1, 1 + len(s1)):
        for y in range(1, 1 + len(s2)):
            if s1[x - 1] == s2[y - 1]:
                m[x][y] = m[x - 1][y - 1] + 1
                if m[x][y] > longest:
                    longest = m[x][y]
                    x_longest = x
            else:
                m[x][y] = 0
    return s1[x_longest - longest: x_longest]

def get_best_match(name, ordered_names):
    max_common_substring = 0
    best_match = None
    for ordered_name in ordered_names:
        common_substring = longest_common_substring(name, ordered_name)
        if len(common_substring) > max_common_substring:
            max_common_substring = len(common_substring)
            best_match = ordered_name
    return best_match


def generate_monthly_chart(data_path, output_path, start_year, start_month, end_year, end_month):
    # 加载数据
    job_df = pd.read_excel(data_path, header=0, sheet_name='招聘')
    # 选择报告期编码为 '202307' 的数据
    selected_data = job_df.loc[job_df['指标名称'] == '地区分布岗位数量'].copy()
    selected_data['统计值'] = selected_data['统计值'].astype(int)
    # 按照集团系名称、报告期编码和招聘地区进行分组聚合
    aggregated_df = selected_data.groupby(['集团系名称', '报告期编码', '招聘地区']).agg({'统计值': 'sum'}).reset_index()
    # 将报告期编码转换为字符串格式
    aggregated_df['报告期编码'] = aggregated_df['报告期编码'].astype(str)
    # 添加年月份列，提取年份和月份前六个字符作为年月份
    aggregated_df['年月份'] = pd.to_datetime(aggregated_df['报告期编码'].str[:6], format='%Y%m')
    # 选择特定时间范围的数据
    filtered_data = aggregated_df[(aggregated_df['年月份'].dt.year > start_year) |
                                  ((aggregated_df['年月份'].dt.year == start_year) &
                                   (aggregated_df['年月份'].dt.month >= start_month))]

    # 按照年月份进行分组并对统计值进行求和
    grouped_data = filtered_data.groupby(['年月份', '招聘地区'])['统计值'].sum().reset_index()
    # 按年月份进行排序
    grouped_data = grouped_data.sort_values(by='年月份')

    # 找到最早的年月份作为 x 轴起始点
    start_date = grouped_data['年月份'].min()

    # 创建一个新的 Figure 对象
    plt.figure(figsize=(10, 6))
    # 循环遍历不同的招聘地区
    for area in grouped_data['招聘地区'].unique():
        # 按照招聘地区筛选数据
        area_data = grouped_data[grouped_data['招聘地区'] == area]
            # 对结束时间进行筛选
        end_date =(f"{end_year}-{end_month}")
        area_data = area_data[area_data['年月份'] <= end_date]
        if area == '本地':
            label = '本地招聘岗位数'
        elif area == '外地':
            label = '外地招聘岗位数'
        else:
            label = area
        # 绘制折线图，x轴为年月份，y轴为统计值，并指定标签
        plt.plot(area_data['年月份'], area_data['统计值'], label=label, linewidth=4)#加粗

 
    # 设置 x 轴日期格式
    plt.gca().xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m'))
    
    # 设置 x 轴的日期刻度间隔为每 2 个月
    plt.gca().xaxis.set_major_locator(mdates.MonthLocator(interval=2))
    plt.gca().xaxis.set_minor_locator(mdates.MonthLocator())

    # 设定 x 轴的显示范围，从最早日期开始
    end_date = datetime.datetime.strptime(end_date, '%Y-%m')
    end_date=mdates.date2num(end_date)
    plt.gca().set_xlim(start_date, end_date)

    
    # 添加图例，调整位置并设置字号
    plt.legend(loc='lower center', bbox_to_anchor=(0.5, -0.15), ncol=2, fontsize=12)
    # 添加标签和标题，设置字号
    # plt.xlabel('年月份', fontsize=8)
    plt.ylabel('单位：个', fontsize=12)
    plt.xticks(fontsize=12)  # 设置x轴标签字号
    plt.yticks(fontsize=12)  # 设置y轴标签字号
    plt.tight_layout()  # 自动调整布局，避免标签重叠
    plt.savefig(output_path)  # 保存图形
    plt.close()  # 关闭图形绘制，释放资源
        # 创建一个 BytesIO 对象
    image_stream = BytesIO()

    # 保存图表到 BytesIO 对象
    # plt.savefig(image_stream, format='png', bbox_inches='tight', dpi=200)
    plt.savefig(image_stream)
    image_stream.seek(0)
    return image_stream


def generate_report(date,data_path,img):
    job_df = pd.read_excel(data_path, header=0, sheet_name='招聘')
    selected_data = job_df[(job_df['指标名称'] == '地区分布岗位数量')]
    selected_data.loc[:, '统计值'] = selected_data['统计值'].astype(int)
    # 按照集团系名称和招聘地区进行分组聚合统计值
    aggregated_df = selected_data.groupby(['集团系名称', '报告期编码', '招聘地区']).agg({'统计值': 'sum'}).reset_index()
    # 筛选出指定日期的数据
    selected_data = aggregated_df[aggregated_df['报告期编码'].astype(str).str.contains(date)]
    total_data = selected_data.groupby(['招聘地区']).agg({'统计值': 'sum'}).reset_index()
    local_val = total_data[total_data['招聘地区'] == '本地']['统计值'].values[0]
    foreign_val = total_data[total_data['招聘地区'] == '外地']['统计值'].values[0]
    total_val = local_val + foreign_val
    # 输出一段文字
    if not total_data.empty:
        output_text = f"{date[:4]}年{date[4:6]}月，北京市信息软件业招聘岗位数{total_val}个，其中本地{local_val}个，外地{foreign_val}个。"    
        time_year= f"（{date[:4]}年{date[4:6]}月）"
    else:
        output_text = "暂无数据。"
    # 按集团系名称进行聚合
    grouped_data = selected_data.groupby(['招聘地区', '集团系名称']).agg({'统计值': 'sum'}).reset_index()
    # 指定的 group_name 列表
    ordered_group_names = ['阿里系', '百度系', '抖音系', '京东系', '快手系', '美团系', '神州数码', '小米系']
    # 获取 unique 集团系名称
    unique_group_names = selected_data['集团系名称'].unique()
    # 创建一个空字典，将模糊匹配的结果存储在其中
    matched_names = {}
    # 遍历每个 unique 集团系名称并找到最匹配的名称
    for name in ordered_group_names:
        # best_match = process.extractOne(name, ordered_group_names)[0]
        best_match = get_best_match(name, unique_group_names)
        matched_names[best_match] = name
    print('原始数据：word文档集团名映射',matched_names)
    
    # 根据模糊匹配的结果，对集团系名称进行重新排序
    selected_data = selected_data.copy()
    selected_data['集团系名称'] = selected_data['集团系名称'].map(matched_names)
    # 按集团系名称进行聚合
    grouped_data=selected_data.groupby(['招聘地区','集团系名称']).agg({'统计值': 'sum'}).reset_index()
    # 输出格式
    output = ""
    report_data = []
    for group_name in ordered_group_names:
        group_data = grouped_data[grouped_data['集团系名称'] == group_name]
        if not group_data.empty:
            row = group_data.iloc[0]
            local_filtered_data = grouped_data[(grouped_data['集团系名称'] == row['集团系名称']) & (grouped_data['招聘地区'] == '本地')]
            foreign_filtered_data = grouped_data[(grouped_data['集团系名称'] == row['集团系名称']) & (grouped_data['招聘地区'] == '外地')]
            if not local_filtered_data.empty:
                local_value = int(local_filtered_data['统计值'].iloc[0])
            else:
                local_value = 0
            if not foreign_filtered_data.empty:
                foreign_value = int(foreign_filtered_data['统计值'].iloc[0])
            else:
                foreign_value = 0
            output += f"{group_name}招聘{local_value+foreign_value}人，其中本地{local_value}人，外地{foreign_value}人；"

            report_data.append({
            '': group_name,
            '本地': str(local_value),
            '外地': str(foreign_value),
            '合计': str(local_value+foreign_value)
            })
    # Convert the list to a Pandas DataFrame
    report_dataframe = pd.DataFrame(report_data)
    prefix='分集团看，'
    output = output[:-1] + "。"
    par=prefix+output

    # 在 Word 文档中添加段落，包含输出文字
    # 插入图片
    img  # 图片路径

    # 实例化 WordReportGenerator 类
    word_generator = WordReportGenerator()

    # 使用类方法创建文件
    word_generator.job_docx_file(title0='信息软件业招聘岗位数分析报告',
                             title1='',
                             par1=output_text,
                             pic_title='图：北京市信息软件业招聘岗位数',
                             table_title='表：分集团系北京市信息软件业招聘岗位数',
                             pars=par,
                             file_pre='信息软件业招聘岗位数分析报告',
                             pic_dir=img,
                             df_table=report_dataframe,
                             file_dir='test',
                             year_month=time_year
                             )


    # 添加其他内容到文档...

def main():
    data_path = '/Users/harvin/code/自动报告产品开发-产业链@20220830/data/input/指标数据20231124-v1.xlsx'
    output_image_path = '/Users/harvin/code/自动报告产品开发-产业链@20220830/data/chart_2022_nov.png'
    start_m=3
    end_year,end_m=2023,10
    end_m+=1
    
    for i in tqdm(range(start_m,end_m)):
        if i<10 and len(str(i))<2:
            m='0'+str(i)
        else:
            m=str(i)
        t=str(end_year)+m
        generate_monthly_chart(data_path, output_image_path, start_year=2021, start_month=9,end_year=end_year,end_month=m)
        generate_report(t,data_path,output_image_path)

if __name__ == "__main__":
    # 执行 main 函数
    main()