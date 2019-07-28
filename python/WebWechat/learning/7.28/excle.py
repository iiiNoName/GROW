import xlrd,xlwt
import jieba
from jieba import analyse

def searchAnswer(text):
    if text == ['exit']:
        answer = '很高兴为您解答拜拜 ฅ( ̳• ◡ • ̳)ฅ\n'
    else:
        res = ''
        p = 0
        num = 1
        pre = ["英文：", "类别：", "含义：", "说明："]
        List = []
        wordBook = xlrd.open_workbook('AiKeywords.xlsx')
        table = wordBook.sheet_by_name('EngWords')
        nrows = table.nrows
        for word in text:
            for row in range(1, nrows):
                cell1 = table.cell(row, 3).value  # 英文
                cell2 = table.cell(row, 4).value  # 类别
                if word in cell1 or word in cell2:
                    rows = table.row_values(row)
                    if rows not in List:
                        List.append(rows)
                    else:
                        continue
        if len(List) != 0:
            if len(List) == 1:
                for i in range(0, 4):
                    res += pre[p] + rows[i + 3] + '\n'
                    p += 1
                res = "第%s条\n" % num + res
            else:
                for list in List:
                    p = 0
                    res = res + "第%s条\n" % num
                    for i in range(0, 4):
                        res += pre[p] + list[i + 3] + '\n'
                        p += 1
                    num += 1
            answer = '为你找到%s条内容O(∩_∩)O\n' % len(List) + res
            # print(answer)
        else:
            answer = '哎呀，我好像没有明白呢\n请换个姿势提问呦(﹡ˆᴗˆ﹡)♡\n'
    # print(answer)
    return answer

def keyWord(text):
    exitWords = []          #退出词汇列表
    noUseWords = []         #无效词汇列表
    wordBook = xlrd.open_workbook("AiKeywords.xlsx")
    table = wordBook.sheet_by_name('CnWords')
    # nrows = table.nrows
    cell1 = table.cell(2, 3).value.split(';')  # 无效词汇   .split(';')变成列表
    cell2 = table.cell(3, 3).value.split(';')  # 退出词汇   .split(';')变成列表
    cell3 = table.cell(1, 3).value.split(';')  # 专业词汇   .split(';')变成列表
    for ProfessionalVocabulary in cell3:
        jieba.add_word(ProfessionalVocabulary)      #往jieba库里添加专业词汇
    keywords = analyse.extract_tags(text,topK=9,withWeight=False,allowPOS=())
    # print(keywords)
    for words in keywords:      # 注意：for遍历循环中的remove操作出现奇怪现象
        if words in cell1:
            noUseWords.append(words)
        elif words in cell2:
            exitWords.append(words)
    for noUse in noUseWords:    #解决办法-不在for操作keywords的同时又remove
        keywords.remove(noUse)  #如此
    # print(exitWords)
    if exitWords and keywords is not None:      #判断退出词汇的占比
        result = len(exitWords)/len(keywords)
        # print(result)
        if result >= 1/2:
            keywords = ['exit']
        else:
            for no in exitWords:
                keywords.remove(no)
    print(keywords)
    return keywords

if __name__ == "__main__":
    while True:
        searchAnswer(keyWord(input('输入：')))
