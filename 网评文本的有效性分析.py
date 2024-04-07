import pandas as pd
import re
import jieba
from snownlp import SnowNLP

# 加载停用词
with open(r'C:\Python\TEST\DATA\cn_stopwords.txt', 'r', encoding='utf-8') as f:
    stopwords = [line.strip() for line in f.readlines()]

# 文本预处理函数
def preprocess_text(text):
    text = re.sub(r'[^\w\s]', '', text)  # 去除标点符号
    words = jieba.cut(text)  # 分词
    filtered_words = [word for word in words if word not in stopwords]  # 去除停用词
    return ' '.join(filtered_words)

# 情感分析函数
# 情感分析函数
def sentiment_analysis(text):
    if not text.strip():  # 检查文本是否为空或只包含空格
        return 0.5  # 返回中立的情感得分
    else:
        return SnowNLP(text).sentiments

# 情感标签函数
def label_sentiment(score):
    if score < 0.4:
        return '消极'
    elif score > 0.6:
        return '积极'
    else:
        return '中性'

# 读取数据
scenic_comments = pd.read_excel(r'C:\Python\TEST\DATA\景区评论.xlsx')
hotel_comments = pd.read_excel(r'C:\Python\TEST\DATA\酒店评论.xlsx')

# 预处理评论
scenic_comments['评论内容'] = scenic_comments['评论内容'].apply(preprocess_text)
hotel_comments['评论内容'] = hotel_comments['评论内容'].apply(preprocess_text)

# 进行情感分析
scenic_comments['情感倾向'] = scenic_comments['评论内容'].apply(sentiment_analysis)
hotel_comments['情感倾向'] = hotel_comments['评论内容'].apply(sentiment_analysis)

# 标记情感
scenic_comments['情感标签'] = scenic_comments['情感倾向'].apply(label_sentiment)
hotel_comments['情感标签'] = hotel_comments['情感倾向'].apply(label_sentiment)

# 输出情感分析结果
print("景区评论情感分析结果：")
print(scenic_comments[['景区名称', '评论日期', '评论内容', '情感倾向', '情感标签']])
print("\n酒店评论情感分析结果：")
print(hotel_comments[['酒店名称', '评论日期', '评论内容', '情感倾向', '情感标签']])
