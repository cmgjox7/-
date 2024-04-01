import pandas as pd
import jieba
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity
from difflib import SequenceMatcher
import re

# 读取景区和酒店的网络评论数据
scenic_comments = pd.read_excel(r'C:\Python\TEST\DATA\景区评论.xlsx')['评论内容']
hotel_comments = pd.read_excel(r'C:\Python\TEST\DATA\酒店评论.xlsx')['评论内容']

# 文本预处理
def preprocess_text(text):
    # 去除特殊字符
    text = re.sub(r'[^\w\s]', '', text)
    # 分词
    words = jieba.cut(text)
    return " ".join(words)

# 计算文本相似度
def text_similarity(text1, text2):
    vectorizer = TfidfVectorizer()
    X = vectorizer.fit_transform([text1, text2])
    similarity = cosine_similarity(X)[0, 1]
    return similarity

# 检测重复评论
def detect_duplicate_comments(comments):
    duplicate_comments = set()
    unique_comments = set()
    for comment in comments:
        if comment in unique_comments:
            duplicate_comments.add(comment)
        else:
            unique_comments.add(comment)
    return duplicate_comments

# 检测内容不相关评论
def detect_irrelevant_comments(comments, description, threshold=0.6):
    irrelevant_comments = []
    for comment in comments:
        similarity = text_similarity(preprocess_text(comment), preprocess_text(description))
        if similarity < threshold:
            irrelevant_comments.append(comment)
    return irrelevant_comments

# 检测简单复制修改评论
def detect_copied_comments(comments, threshold=0.8):
    copied_comments = []
    for i in range(len(comments)):
        for j in range(i+1, len(comments)):
            similarity = SequenceMatcher(None, comments[i], comments[j]).ratio()
            if similarity > threshold:
                copied_comments.append((comments[i], comments[j]))
    return copied_comments

# 景区评论分析
print("景区评论分析：")
print("----------------------")
# 检测重复评论
print("重复评论检测：")
duplicate_scenic_comments = detect_duplicate_comments(scenic_comments)
print("重复评论数量:", len(duplicate_scenic_comments))
print("----------------------")
# 检测内容不相关评论
print("内容不相关评论检测：")
description = "景区风景优美，设施完善，服务周到"
irrelevant_scenic_comments = detect_irrelevant_comments(scenic_comments, description)
print("内容不相关评论数量:", len(irrelevant_scenic_comments))
print("----------------------")
# 检测简单复制修改评论
print("简单复制修改评论检测：")
copied_scenic_comments = detect_copied_comments(scenic_comments)
print("简单复制修改评论数量:", len(copied_scenic_comments))
print("----------------------")

# 酒店评论分析
print("酒店评论分析：")
print("----------------------")
# 检测重复评论
print("重复评论检测：")
duplicate_hotel_comments = detect_duplicate_comments(hotel_comments)
print("重复评论数量:", len(duplicate_hotel_comments))
print("----------------------")
# 检测内容不相关评论
print("内容不相关评论检测：")
description = "酒店环境优美，服务一流，交通便利"
irrelevant_hotel_comments = detect_irrelevant_comments(hotel_comments, description)
print("内容不相关评论数量:", len(irrelevant_hotel_comments))
print("----------------------")
# 检测简单复制修改评论
print("简单复制修改评论检测：")
copied_hotel_comments = detect_copied_comments(hotel_comments)
print("简单复制修改评论数量:", len(copied_hotel_comments))
print("----------------------")
