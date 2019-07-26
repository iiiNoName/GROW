import xlrd,xlwt
from jieba import analyse

def searchAnswer(text):
    res = ''
    p = 0
    num = 1
    pre = ["英文：", "类别：", "含义：", "说明："]
    List = []
    wordBook = xlrd.open_workbook('AiKeywords.xlsx')
    table = wordBook.sheet_by_name('EngWords')
    nrows = table.nrows
    for word in text:
        for row in range(1,nrows):
            cell1 = table.cell(row,3).value         #英文
            cell2 = table.cell(row,4).value         #类别
            if word in cell1 or word in cell2:
                rows = table.row_values(row)
                if rows not in List:
                    List.append(rows)
                else:
                    continue
    if len(List) != 0:
        if len(List) == 1:
            for i in range(0,4):
                res += pre[p] + rows[i + 3] + '\n'
                p += 1
            res = "第%s条\n" % num + res
        else:
            for list in List:
                p = 0
                res = res + "第%s条\n"%num
                for i in range(0, 4):
                    res += pre[p] + list[i + 3] + '\n'
                    p += 1
                num += 1
        answer = '为你找到%s条内容O(∩_∩)O\n'%len(List) + res
        # print(answer)
    else:
        answer = '哎呀，我好像没有明白呢\n请换个姿势提问呦(﹡ˆᴗˆ﹡)♡\n'
    # print(answer)
    return answer

def keyWord(text):
    keywords = analyse.extract_tags(text,topK=5,withWeight=False,allowPOS=())
    wordBook = xlrd.open_workbook("AiKeywords.xlsx")
    table = wordBook.sheet_by_name('CnWords')
    # nrows = table.nrows
    cell1 = table.cell(2, 3).value  # 无效词汇
    for words in keywords:
        if words in cell1:
            keywords.remove(words)
    # print(keywords)
    return keywords


if __name__ == "__main__":
    while True:
        searchAnswer(keyWord(input('输入：')))

'''
怎么用import导入库函数呢
怎么用import导入模块呢
import库函数
请问变量是什么呢？


'''