import xlrd,xlwt
from jieba import analyse
from xlutils.copy import copy
import datetime

sum = 1

def searchAnswer(text):
    for word in text:
        res = ''
        p = 0
        num = 1
        pre = ["英文：","类别：", "含义：", "说明："]
        List=[]
        wordBook = xlrd.open_workbook('SearchPythonInfo(1).xlsx')
        table = wordBook.sheet_by_name('words')
        nrows = table.nrows
        for row in range(1,nrows):
            cell1 = table.cell(row,3).value         #英文
            cell2 = table.cell(row,4).value         #类别
            if word in cell1 or word in cell2:
                rows = table.row_values(row)
                List.append(rows)
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
        return answer

def keyWord(text):
    keywords = analyse.extract_tags(text,topK=5,withWeight=False,allowPOS=())
    wordBook = xlrd.open_workbook("NoUseWords.xlsx")
    table = wordBook.sheet_by_name('nousewords')
    nrows = table.nrows
    for words in keywords:
        for row in range(1,nrows):
            cell1 = table.cell(row, 1).value  # 中文词汇
            cell2 = table.cell(row, 2).value  # 英文词汇
            if words in cell1 or words in cell2:
                keywords.remove(words)
    return keywords

def chatHistory(num,now,FromName,ToName,type,content):
    List = [num,now,FromName,ToName,type,content]
    print(List)
    reexcle = xlrd.open_workbook("History.xls")
    file = copy(reexcle)
    table = file.get_sheet(0)
    # ncols = table.ncols
    for row in range(num,num+1):
        for col in range(0,6):
            table.write(row,col,List[col])
    file.save("History.xls")




if __name__ == "__main__":
    while True:
        # searchAnswer(keyWord(input('输入：')))
        chatHistory(sum,datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'),input(1),input(2),1,input(3))
        sum += 1