import pandas as pd
import jieba
from collections import Counter
import xlwt
from sklearn.linear_model import LinearRegression
from sklearn.metrics import mean_squared_error

#step 1
# 读取停用词表
with open(r'C:\Python\TEST\DATA\cn_stopwords.txt', 'r', encoding='utf-8') as f:
    stopwords = [line.strip() for line in f]

# 读取景区评论和酒店评论数据
scenic_comments = pd.read_excel(r'C:\Python\TEST\DATA\景区评论.xlsx')['评论内容']
hotel_comments = pd.read_excel(r'C:\Python\TEST\DATA\酒店评论.xlsx')['评论内容']

# 分词函数
def segment(text):
    words = jieba.cut(text)
    words = [word for word in words if word not in stopwords and len(word) > 1]  # 过滤停用词和长度为1的词
    return words

# 处理景区评论
scenic_words = []
for comment in scenic_comments:
    scenic_words.extend(segment(comment))

# 处理酒店评论
hotel_words = []
for comment in hotel_comments:
    hotel_words.extend(segment(comment))

# 统计词频
scenic_word_counts = Counter(scenic_words)
hotel_word_counts = Counter(hotel_words)

# 合并词频并计算总的热度
total_word_counts = scenic_word_counts + hotel_word_counts
total_word_counts_sorted = sorted(total_word_counts.items(), key=lambda x: x[1], reverse=True)

# 选择热度最高的前20个词作为目的地TOP20热门词
top20_hot_words = total_word_counts_sorted[:20]

# 保存结果到Excel文件
output_file = r'印象词云表.xls'
workbook = xlwt.Workbook(encoding='utf-8')
worksheet = workbook.add_sheet('Sheet1')

worksheet.write(0, 0, '热词')
worksheet.write(0, 1, '热度')

for i, (word, count) in enumerate(top20_hot_words):
    worksheet.write(i + 1, 0, word)
    worksheet.write(i + 1, 1, count)

# workbook.save(output_file)


#step2
# 读取景区评分和酒店评分数据
scenic_scores = pd.read_excel(r'C:\Python\TEST\DATA\景区评分.xlsx')
hotel_scores = pd.read_excel(r'C:\Python\TEST\DATA\酒店评分.xlsx')

# 构建特征矩阵和目标变量
X_scenic = scenic_scores[['服务得分', '位置得分', '设施得分', '卫生得分', '性价比得分']]
y_scenic = scenic_scores['总得分']

X_hotel = hotel_scores[['服务得分', '位置得分', '设施得分', '卫生得分', '性价比得分']]
y_hotel = hotel_scores['总得分']

# 建立线性回归模型
model_scenic = LinearRegression()
model_hotel = LinearRegression()

# 拟合模型
model_scenic.fit(X_scenic, y_scenic)
model_hotel.fit(X_hotel, y_hotel)

# 预测总得分
y_pred_scenic = model_scenic.predict(X_scenic)
y_pred_hotel = model_hotel.predict(X_hotel)

# 计算MSE
mse_scenic = mean_squared_error(y_scenic, y_pred_scenic)
mse_hotel = mean_squared_error(y_hotel, y_pred_hotel)

print("景区评分模型的MSE:", mse_scenic)
print("酒店评分模型的MSE:", mse_hotel)
# 定义权重
weights = {
    '服务得分': 0.2,
    '位置得分': 0.2,
    '设施得分': 0.2,
    '卫生得分': 0.2,
    '性价比得分': 0.2
}

# 计算加权平均得分函数
def weighted_average(row):
    total_score = 0
    for dimension, weight in weights.items():
        total_score += row[dimension] * weight
    return total_score

# 应用加权平均函数得到总得分预测值
scenic_scores['总得分预测'] = scenic_scores.apply(weighted_average, axis=1)
hotel_scores['总得分预测'] = hotel_scores.apply(weighted_average, axis=1)

# 输出结果
print("景区评分表:")
print(scenic_scores)
print("\n酒店评分表:")
print(hotel_scores)

#step3
