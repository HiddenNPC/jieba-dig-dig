# -*- coding: utf-8 -*-
from __future__ import print_function
import jieba
import xlrd
import os
import jieba.posseg as psg
import jieba.analyse as analyse


def get_all_word_count_dict():
    workbook = xlrd.open_workbook(os.path.abspath("./raw_data/excel/1.xlsx"))
    default_sheet = workbook.sheet_by_index(0)
    # 第一个循环找到相对靠谱分词，存储在counts中
    counts = {}
    # =["业务技能","知识背景","人际沟通","团队协作","组织协调","外语英语","语言表达","写作学习","思维学习","抗压能力","执行力","适应与变通能力","独立工作能力","责任心","主动性","职业操守","工作热情","细节"]
    for row in range(0, default_sheet.nrows):
        if row > 0:
            this_value = default_sheet.cell(row, 5).value
            # ex = {'有庆', '我们', '知道', '看到', '自己', '起来'}
            words = jieba.lcut(this_value)
            # words.append(u"执行力")
            for word in words:
                if len(word) < 2:
                    continue
                else:
                    counts[word] = counts.get(word, 0) + 1
            if u"细心" in this_value:
                counts[u"细心"] = counts.get(u"细心", 0) + 1
            if u"细节" in this_value:
                counts[u"细节"] = counts.get(u"细节", 0) + 1
    return counts, default_sheet


def get_synonyms_dict(counts):
    key_list1 = counts.keys()
    key_list2 = counts.keys()
    synonyms_dict = {}
    traversed_key_list = []
    traversed_key_list2 = []
    occupation_skills_list = [u"业务技能"]
    knowledge_list = [u"知识背景"]
    lists_dict = {}
    standard_list = [u"人际沟通能力", u"团队协作能力", u"组织协调能力", u"外语能力",
                     u"语言表达能力", u"写作能力", u"学习能力",
                     u"思维学习能力", u"抗压能力", u"执行力", u"适应与变通能力",
                     u"独立工作能力", u"责任心", u"主动性",
                     u"职业操守", u"工作热情", u"注重细节"]
    for key in standard_list:
        lists_dict[key] = [key]

    for key1 in key_list1:
        # 找出key所有同义词list
        # 业务技能
        if u"工程" in key1 or u"技能" in key1 or u"编程" in key1 or u"C语言" in key1 or u"C++" in key1 or u"脚本" in key1 \
                or u"汇编" in key1:
            occupation_skills_list.append(key1)
            if key1 in key_list2:
                key_list2.remove(key1)
        # 知识背景
        elif u"处理" in key1 or u"设计" in key1 or u"系统" in key1:
            knowledge_list.append(key1)
            if key1 in key_list2:
                key_list2.remove(key1)
        elif u"人际" in key1 or u"沟通" in key1 or u"公关" in key1 or u"形象" in key1 or u"交流" in key1 or u"理解" in key1:
            lists_dict[u"人际沟通能力"].append(key1)
            if key1 in key_list2:
                key_list2.remove(key1)
        elif u"团队" in key1 or u"协作" in key1 or u"纪律" in key1:
            lists_dict[u"团队协作能力"].append(key1)
            if key1 in key_list2:
                key_list2.remove(key1)
        elif u"组织" in key1 or u"协调" in key1:
            lists_dict[u"组织协调能力"].append(key1)
            if key1 in key_list2:
                key_list2.remove(key1)
        elif u"外语" in key1 or u"中英文" in key1 or u"英文" in key1 or u"英语" in key1 or u"日语" in key1 or u"美语" in key1 or u"意大利语" in key1 or u"法语" in key1 or u"俄语" in key1 or u"韩语" in key1 or u"朝鲜语" in key1 or u"多语种" in key1:
            lists_dict[u"外语能力"].append(key1)
            if key1 in key_list2:
                key_list2.remove(key1)
        elif u"语言" in key1 or u"表达" in key1:
            lists_dict[u"语言表达能力"].append(key1)
            if key1 in key_list2:
                key_list2.remove(key1)
        elif u"写作" in key1 or u"书写" in key1 or u"读写" in key1 or u"文笔" in key1:
            lists_dict[u"写作能力"].append(key1)
            if key1 in key_list2:
                key_list2.remove(key1)
        elif u"学习" in key1 or u"好学" in key1 or u"奖学金" in key1:
            lists_dict[u"学习能力"].append(key1)
            if key1 in key_list2:
                key_list2.remove(key1)
        elif u"思维" in key1 or u"思路" in key1 or u"活跃" in key1 \
                or u"分析" in key1 or u"判断" in key1 or u"思考" in key1 or u"领悟" in key1:
            lists_dict[u"思维学习能力"].append(key1)
            if key1 in key_list2:
                key_list2.remove(key1)
        elif u"抗压" in key1 or u"艰苦" in key1 or u"奋斗" in key1 or u"健康" in key1 or u"坚持" in key1 or u"坚韧" in key1 or u"承受" in key1:
            lists_dict[u"抗压能力"].append(key1)
            if key1 in key_list2:
                key_list2.remove(key1)
        elif u"执行" in key1 or u"行动" in key1 or u"做事" in key1 or u"完成" in key1:
            lists_dict[u"执行力"].append(key1)
            if key1 in key_list2:
                key_list2.remove(key1)
        elif u"适应" in key1 or u"变通" in key1 or u"应变" in key1 or u"能动性" in key1:
            lists_dict[u"适应与变通能力"].append(key1)
            if key1 in key_list2:
                key_list2.remove(key1)
        elif u"独立" in key1:
            lists_dict[u"独立工作能力"].append(key1)
            if key1 in key_list2:
                key_list2.remove(key1)
        elif u"责任" in key1 or u"认真" in key1 or u"党员" in key1:
            lists_dict[u"责任心"].append(key1)
            if key1 in key_list2:
                key_list2.remove(key1)
        elif u"主动" in key1 or u"积极" in key1:
            lists_dict[u"主动性"].append(key1)
            if key1 in key_list2:
                key_list2.remove(key1)
        elif u"操守" in key1 or u"道德" in key1 or u"爱岗" in key1 or u"敬业" in key1 or u"正直" in key1 or u"诚信" in key1:
            lists_dict[u"职业操守"].append(key1)
            if key1 in key_list2:
                key_list2.remove(key1)
        elif u"热情" in key1 or u"热爱" in key1:
            lists_dict[u"工作热情"].append(key1)
            if key1 in key_list2:
                key_list2.remove(key1)
        elif u"细节" in key1 or u"细心" in key1 or u"细致" in key1:
            lists_dict[u"注重细节"].append(key1)
            if key1 in key_list2:
                key_list2.remove(key1)
        elif u"能力" in key1:
            synonyms_dict[key1] = [key1]
            if key1 in key_list2:
                key_list2.remove(key1)
        else:
            pass

    for key1 in key_list1:
        traversed_key_list.append(key1)
        if u"工程" in key1 or u"技能" in key1 or u"编程" in key1 or u"C语言" in key1 or u"C++" in key1 or u"脚本" in key1 \
                or u"汇编" in key1 or u"处理" in key1 or u"设计" in key1 or u"系统" in key1 or u"人际" in key1 \
                or u"沟通" in key1 or u"公关" in key1 or u"形象" in key1 or u"交流" in key1 or u"交流" in key1 \
                or u"团队" in key1 or u"协作" in key1 or u"纪律" in key1 or u"组织" in key1 or u"协调" in key1 \
                or u"外语" in key1 or u"中英文" in key1 or u"英文" in key1 or u"英语" in key1 or u"日语" in key1 \
                or u"美语" in key1 or u"意大利语" in key1 or u"法语" in key1 or u"俄语" in key1 or u"韩语" in key1 \
                or u"朝鲜语" in key1 or u"多语种" or u"语言" in key1 or u"表达" in key1 or u"写作" in key1 or u"学习" \
                in key1 or u"好学" in key1 or u"奖学金" in key1 or u"书写" in key1 or u"读写" in key1 or u"文笔" \
                in key1 or u"思维" in key1 or u"思路" in key1 or u"活跃" in key1 or u"分析" in key1 or u"判断" \
                or u"思考" in key1 in key1 or u"领悟" in key1 or u"抗压" in key1 or u"艰苦" in key1 or u"奋斗" in key1 \
                or u"健康" in key1 or u"坚持" in key1 or u"坚韧" in key1 or u"承受" in key1 or u"执行" in key1 \
                or u"行动" in key1 or u"做事" in key1 or u"完成" in key1 or u"适应" in key1 or u"变通" in key1 \
                or u"应变" in key1 or u"能动性" in key1 or u"独立" in key1 or u"责任" in key1 or u"认真" in key1 \
                or u"党员" in key1 or u"主动" in key1 or u"积极" in key1 or u"道德" in key1 or u"操守" in key1 \
                or u"爱岗" in key1 or u"敬业" in key1 or u"诚信" in key1 or u"敬业" in key1 or u"正直" in key1 \
                or u"热情" in key1 or u"热爱" in key1 or u"细节" in key1 or u"细心" in key1 or u"细致" in key1:
            continue
        synonyms_list = [key1]
        for index, key2 in enumerate(key_list2):
            if key2 in traversed_key_list:
                continue
            # if key2 in traversed_key_list2:
            #     continue
            if key1 == key2:
                continue
            union_result = set(key1).union(set(key2))
            intersection_result = set(key1).intersection(set(key2))
            if len(union_result) > len(key1) and len(union_result) > len(key2) and len(
                    intersection_result) > 1:
                synonyms_list.append(key2)
        synonyms_dict[key1] = synonyms_list
    synonyms_dict[u"业务技能"] = occupation_skills_list
    synonyms_dict[u"知识背景"] = knowledge_list
    for list_to_be_appended in lists_dict:
        synonyms_dict[list_to_be_appended] = lists_dict[list_to_be_appended]
    return synonyms_dict, lists_dict


# def join_synonyms_lists(synonyms_dict):
#     for key in synonyms_dict:
#
# synonyms_dict = {}
# for synonyms_pair in synonyms_list:
#     synonyms_dict[synonyms_pair[0]] = counts.get(synonyms_pair[0]) + counts.get(synonyms_pair[1])
#     counts[synonyms_pair[0]] = 0


def get_refered_count_dict(counts, synonyms_dict, lists_dict):
    # 统计所有靠谱分词在excel各列中出现的次数
    refered_count_dict = {}
    for key in synonyms_dict:
        for synonyms in synonyms_dict[key]:
            if synonyms == u"业务技能" or synonyms == u"知识背景" or synonyms in lists_dict.keys():
                continue
            refered_count_dict[key] = refered_count_dict.get(key, 0) + counts.get(synonyms)

    return refered_count_dict


def get_percentage_refered_dict(refered_count_dict, default_sheet):
    # 算出每个词的出现频率
    percentage_refered_dict = {}
    percentage_lt10_dict = {}
    for key in refered_count_dict:
        percentage_refered = float(refered_count_dict[key]) / float(default_sheet.nrows) * 100
        percentage_refered_dict[key] = percentage_refered
        if percentage_refered > float(10):
            percentage_lt10_dict[key] = percentage_refered
    return percentage_refered_dict, percentage_lt10_dict


def show_percentage_lt10_dict(refered_count_dict, percentage_lt10_dict, synonyms_dict, default_sheet):
    for key in percentage_lt10_dict:
        print("--------------------------------------")
        if type(key) == "str":
            ukey = key
            ukey.decode("utf-8")
            if not synonyms_dict[key][0] == key:
                print(key, end=" ")
        for synonyms in synonyms_dict[key]:
            print("%s" % synonyms, end=" ")
        print("|%s|(%s/%s)" % (percentage_lt10_dict[key], refered_count_dict[key], default_sheet.nrows))


def show_percentage_refered_dict(refered_count_dict, percentage_refered_dict, synonyms_dict, default_sheet):
    for key in percentage_refered_dict:
        print("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
        if type(key) == "str":
            ukey = key
            ukey.decode("utf-8")
            if not synonyms_dict[key][0] == key:
                print(key, end=" ")
        if not synonyms_dict[key][0] == key:
            print(key, end=" ")
        for synonyms in synonyms_dict[key]:
            print("%s" % synonyms, end=" ")
        print("|%s|(%s/%s)" % (percentage_refered_dict[key], refered_count_dict[key], default_sheet.nrows))

        # for word in ex:
        #     del (counts[word])

        # for i in range(10):
        #     word, count = items[i]
        #     print ("{:<10}{:>5}".format(word, count))
        #
        # wz = open('ms.txt', 'w+')
        # wz.write(str(ls))


if __name__ == '__main__':
    counts, default_sheet = get_all_word_count_dict()
    synonyms_dict, lists_dict = get_synonyms_dict(counts)
    refered_count_dict = get_refered_count_dict(counts, synonyms_dict, lists_dict)
    percentage_refered_dict, percentage_lt10_dict = get_percentage_refered_dict(refered_count_dict, default_sheet)
    show_percentage_lt10_dict(refered_count_dict, percentage_lt10_dict, synonyms_dict, default_sheet)
    show_percentage_refered_dict(refered_count_dict, percentage_refered_dict, synonyms_dict, default_sheet)

cccccc = "xxx"

# import matplotlib.pyplot as plt
# from wordcloud import WordCloud
#
# wzhz = WordCloud().generate(txt)
# plt.imshow(wzhz)
# plt.show()
