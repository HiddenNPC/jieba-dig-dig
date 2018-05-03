# -*- coding: utf-8 -*-
import jieba
import xlrd
import os
import jieba.posseg as psg

workbook = xlrd.open_workbook(os.path.abspath("./raw_data/excel/1.xlsx"))
default_sheet = workbook.sheet_by_index(0)
# 第一个循环找到相对靠谱分词，存储在counts中
counts = {}
# =["业务技能","知识背景","人际沟通","团队协作","组织协调","外语英语","语言表达","写作学习","思维学习","","","","","","","","","","",]
for row in range(0, default_sheet.nrows):
    if row > 0:
        this_value = default_sheet.cell(row, 5).value
        this_value.encode("utf-8")
        # ex = {'有庆', '我们', '知道', '看到', '自己', '起来'}
        ls = []
        words = jieba.lcut(this_value, cut_all=True)
        for word in words:
            if len(word) < 3:
                continue
            else:
                for this_word in psg.cut(word):
                    if this_word.flag.startswith('n'):
                        sub_words = jieba.lcut(this_value, cut_all=True)
                        # for sub_word in sub_words:
                        #     if len(sub_word)>1:
                        counts[word] = counts.get(word, 0) + 1
                        # items = list(counts.items())
                        # items.sort(key=lambda x: x[1], reverse=True)

# 第二个循环统计所有靠谱分词在excel各列中出现的次数
refered_count_dict={}
for key in counts:
    for row in range(0, default_sheet.nrows):
        if row > 0:
            this_value = default_sheet.cell(row, 5).value
            this_value.encode("utf-8")
            if key in this_value:
                refered_count_dict[key]=refered_count_dict.get(key, 0) + 1
# 第三个循环算出每个词的出现频率
percentage_refered_dict={}
for key in refered_count_dict:
    percentage_refered=float(refered_count_dict[key])/float(len(refered_count_dict.keys()))*100
    # if percentage_refered>float(10):
    percentage_refered_dict[key]= percentage_refered

                        # for word in ex:
                        #     del (counts[word])


                        # for i in range(10):
                        #     word, count = items[i]
                        #     print ("{:<10}{:>5}".format(word, count))
                        #
                        # wz = open('ms.txt', 'w+')
                        # wz.write(str(ls))

cccccc = "xxx"

# import matplotlib.pyplot as plt
# from wordcloud import WordCloud
#
# wzhz = WordCloud().generate(txt)
# plt.imshow(wzhz)
# plt.show()
