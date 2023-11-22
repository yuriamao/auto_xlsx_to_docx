import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.font_manager import FontProperties
from matplotlib.dates import DateFormatter, MonthLocator  # Add this line
from docx import Document
from io import BytesIO

# 设置中文显示
plt.rcParams['font.sans-serif'] = ['Arial Unicode MS']  # 设置中文字体为系统自带的中文字体，如 macOS 的 Arial Unicode MS

# 读取配置文件
config_df = pd.read_excel(r'/Users/harvin/code/自动报告产品开发-产业链@20220830/data/pmi_config.xlsx', header=0, sheet_name='指标数据')

import pandas as pd

def compare_indices(selected_month_1, selected_month_2):
    # 假设 config_df 包含的字段为：'地区编码'、'产业名称'、'数据值'、'报告期编码'
    # 假设数据已经加载到 config_df 中

    # 选择特定产业名称的数据
    desired_industries = ['生产指数', '新订单指数', '原材料库存指数', '主要原材料购进价格指数', '供应商配送时间指数']
    selected_data = config_df[config_df['产业名称'].isin(desired_industries)]

    # 筛选出 2021 年及之后的数据
    selected_data['报告期编码'] = pd.to_datetime(selected_data['报告期编码'], format='%Y%m', errors='coerce')
    selected_data = selected_data[selected_data['报告期编码'].dt.year >= 2021]

    # 去除非日期值的行
    selected_data.dropna(subset=['报告期编码'], inplace=True)

    # 转换数据值列为数值型，跳过无法转换的值
    selected_data['数据值'] = pd.to_numeric(selected_data['数据值'], errors='coerce')

    # 去除非数值型数据行
    selected_data.dropna(subset=['数据值'], inplace=True)

    # 按指标名称和月份进行数据分组，计算每个月的平均值
    # 按指标名称、地区编码和月份进行数据分组，计算每个月的平均值，并保留地区编码列
    grouped_data = selected_data.groupby(['地区编码', '产业名称', pd.Grouper(key='报告期编码', freq='1M')])['数据值'].mean().reset_index()
    pivot_data = grouped_data 
    pivot_data['报告期编码'] = grouped_data['报告期编码'].dt.to_period('M')

    # 重新命名列名，确保列名符合 'YYYY-MM' 的格式
    # pivot_data.columns = pivot_data.columns.strftime('%Y-%m') if isinstance(pivot_data.columns, pd.DatetimeIndex) else pivot_data.columns

    # 获取你希望对比的两个月份的数据
    # 从数据中筛选出全国和北京的数据，并进行对比
    national_data = pivot_data[pivot_data['地区编码'] == '000000']
    beijing_data = pivot_data[pivot_data['地区编码'] == '110000']

    # 选择特定月份的数据
    national_month_1_data = national_data[national_data['报告期编码'] == selected_month_1]
    national_month_2_data = national_data[national_data['报告期编码'] == selected_month_2]
    beijing_month_1_data = beijing_data[beijing_data['报告期编码'] == selected_month_1]
    beijing_month_2_data = beijing_data[beijing_data['报告期编码'] == selected_month_2]

    text = ""

    for index_name in desired_industries:
        national_index_month_1 = national_month_1_data[national_month_1_data['产业名称'] == index_name]
        national_index_month_2 = national_month_2_data[national_month_2_data['产业名称'] == index_name]
        beijing_index_month_1 = beijing_month_1_data[beijing_month_1_data['产业名称'] == index_name]
        beijing_index_month_2 = beijing_month_2_data[beijing_month_2_data['产业名称'] == index_name]

        if not national_index_month_1.empty and not national_index_month_2.empty:
            nation_value1 = national_index_month_1['数据值'].iloc[0]
            nation_value2 = national_index_month_2['数据值'].iloc[0]
            
            diff_national = nation_value1 - nation_value2
            nation_percentage_change = (diff_national / nation_value2) * 100
            
            text += f"{index_name}为{nation_value1:.1f}%，较上月{'提高' if diff_national > 0 else '下降'}{abs(diff_national):.1f}个百分点，"
            
            if not beijing_index_month_1.empty and not beijing_index_month_2.empty:
                beijing_value1 = beijing_index_month_1['数据值'].iloc[0]
                beijing_value2 = beijing_index_month_2['数据值'].iloc[0]

                diff_beijing = beijing_value1 - beijing_value2
                beijing_percentage_change = (diff_beijing / beijing_value2) * 100

                text += f"较全国{'低' if diff_beijing > 0 else '高'}{abs(diff_beijing):.1f}个百分点；\n"
            else:
                text += "北京缺少数据；\n"
        else:
            text += "全国缺少数据；\n"

    return text

# 使用函数并将结果保存在变量中
selected_month_1 = '2022-07'
selected_month_2 = '2022-06'
result_text = compare_indices(selected_month_1, selected_month_2)

# 打印结果
print(result_text)
