import pandas as pd
from wordcloud import WordCloud
from pandas import ExcelWriter
import matplotlib.pyplot as plt

# 读取景区评论数据
scenic_comments = pd.read_excel(r'C:\Python\TEST\DATA\景区评论.xlsx')

# 将评论内容拼接成一个长文本
scenic_text = ' '.join(scenic_comments['评论内容'])

# 创建词云对象
scenic_wordcloud = WordCloud(font_path='simhei.ttf', width=800, height=400, background_color='white').generate(scenic_text)

# 绘制词云
# plt.figure(figsize=(10, 5))
# plt.imshow(scenic_wordcloud, interpolation='bilinear')
# plt.title('景区评论词云')
# plt.axis('off')
# plt.show()

# 读取酒店评论数据
hotel_comments = pd.read_excel(r'C:\Python\TEST\DATA\酒店评论.xlsx')

# 将评论内容拼接成一个长文本
hotel_text = ' '.join(hotel_comments['评论内容'])

# 创建词云对象
hotel_wordcloud = WordCloud(font_path='simhei.ttf', width=800, height=400, background_color='white').generate(hotel_text)

# 绘制词云
# plt.figure(figsize=(10, 5))
# plt.imshow(hotel_wordcloud, interpolation='bilinear')
# plt.title('酒店评论词云')
# plt.axis('off')
# plt.show()

# 选择综合评价高、中、低三个层次的各 3 家景点和 3家酒店，结合模型的结果，分析他们各自的特
# 读取评分数据
hotel_scores = pd.read_excel(r'C:\Python\TEST\文本与图像挖掘\游客目的地印象分析\酒店评分.xlsx')
scenic_scores = pd.read_excel(r'C:\Python\TEST\文本与图像挖掘\游客目的地印象分析\景区评分.xlsx')

# 按总得分对景区和酒店进行排序
sorted_hotels = hotel_scores.sort_values(by='总得分', ascending=False)
sorted_scenics = scenic_scores.sort_values(by='总得分', ascending=False)

# 选择总得分排名前三、中间三和末尾三的景区和酒店
top_hotels = sorted_hotels.head(3)
middle_hotels = sorted_hotels.iloc[len(sorted_hotels)//2-1:len(sorted_hotels)//2+2]
bottom_hotels = sorted_hotels.tail(3)

top_scenics = sorted_scenics.head(3)
middle_scenics = sorted_scenics.iloc[len(sorted_scenics)//2-1:len(sorted_scenics)//2+2]
bottom_scenics = sorted_scenics.tail(3)
# 定义函数用于打印景区或酒店的特色分析
def analyze_entity(entity_name, entity_data):
    print(f"--- {entity_name} ---")
    for index, row in entity_data.iterrows():
        print(f"名称：{row['景区名称'] if '景区名称' in row else row['酒店名称']}")
        print(f"总得分：{row['总得分']}")
        print(f"服务得分：{row['服务得分'] if '服务得分' in row else row['服务得分']}")
        print(f"位置得分：{row['位置得分'] if '位置得分' in row else row['位置得分']}")
        print(f"设施得分：{row['设施得分'] if '设施得分' in row else row['设施得分']}")
        print(f"卫生得分：{row['卫生得分'] if '卫生得分' in row else row['卫生得分']}")
        print(f"性价比得分：{row['性价比得分'] if '性价比得分' in row else row['性价比得分']}")
        print()

# 分析景区
analyze_entity("景区", top_scenics)
analyze_entity("景区", middle_scenics)
analyze_entity("景区", bottom_scenics)

# 分析酒店
analyze_entity("酒店", top_hotels)
analyze_entity("酒店", middle_hotels)
analyze_entity("酒店", bottom_hotels)

# 创建一个ExcelWriter对象
writer = ExcelWriter(r'C:\Python\TEST\文本与图像挖掘\游客目的地印象分析\各层次特色.xlsx')

# 将数据写入Excel文件的不同sheet
top_scenics.to_excel(writer, sheet_name='景区-总得分高')
middle_scenics.to_excel(writer, sheet_name='景区-总得分中')
bottom_scenics.to_excel(writer, sheet_name='景区-总得分低')

top_hotels.to_excel(writer, sheet_name='酒店-总得分高')
middle_hotels.to_excel(writer, sheet_name='酒店-总得分中')
bottom_hotels.to_excel(writer, sheet_name='酒店-总得分低')

# 将数据写入Excel文件的不同sheet
with pd.ExcelWriter(r'C:\Python\TEST\文本与图像挖掘\游客目的地印象分析\各层次特色.xlsx') as writer:
    top_scenics.to_excel(writer, sheet_name='景区-总得分高', index=False)
    middle_scenics.to_excel(writer, sheet_name='景区-总得分中', index=False)
    bottom_scenics.to_excel(writer, sheet_name='景区-总得分低', index=False)
    top_hotels.to_excel(writer, sheet_name='酒店-总得分高', index=False)
    middle_hotels.to_excel(writer, sheet_name='酒店-总得分中', index=False)
    bottom_hotels.to_excel(writer, sheet_name='酒店-总得分低', index=False)
