import xlrd
import xlwt
import random
punction=['，','。','？','！','-','（','）',',','：','；','：','0','1','2','3','4','5','6','7','8','9']

# coding=utf-8
workbook = xlrd.open_workbook("data/100.xlsx")
names = workbook.sheet_names()
name = workbook.sheet_names()[0]
worksheet = workbook.sheet_by_index(0)
row = worksheet.nrows
rows=[]
for i in range(row):
    rows.append(worksheet.row_values(i))
want=[]
want_word=[]
# want_line=[]
for i in range(len(rows)):
    line=rows[i]
    line=line[0]
    # print(len(temp1))
    word_id=random.randint(0,len(line)-1)
    # temp=line[word_id]
    # print(word_id)
    word=line[word_id]
    while(str(word) in punction or  not str(word).isalnum()):
        word_id =random.randint(0,len(line)-1)
        word=line[word_id]
    # print(word)
    flag=random.randint(0,2)

    if flag==0 and word_id+1<len(line)-1 and str(line[word_id+1]) not in punction:
        new_line=str(line[:word_id+1])+str(word)+str(line[word_id+1])+str(line[word_id+1:])
        new_word=str(word)+str(word)+str(line[word_id+1])+str(line[word_id+1])+'#'+str(word)+str(line[word_id+1])
    else :
        if flag==1 and word_id+1<len(line)-1 and str(line[word_id+1]) not in punction:
            new_line = str(line[:word_id + 1]) + str(line[word_id + 1]) + str(word) + str(line[word_id + 1:])
            new_word = str(word)+ str(line[word_id + 1])  + str(word) + str(line[word_id + 1]) + '#' + str(word) + str(
                line[word_id + 1])
        else:
            new_line=str(line[:word_id+1])+str(word)+str(line[word_id+1:])
            new_word=str(word)+str(word)+'#'+str(word)
    want.append(new_line)
    want_word.append(new_word)
#
# for i in range(len(want)):
#     print(want_word[i])
#     print(want[i])
#     print(rows[i][0])
#     print('\n')
book = xlwt.Workbook(encoding="utf-8", style_compression=0)
sheet = book.add_sheet("test01", cell_overwrite_ok=True)
# # sheet.write(0,0,"各省市")
# #
#
#
for i in range(len(want)):
    sheet.write(i+1,0,want[i])
for i in range(len(want_word)):
    sheet.write(i+1,1,want_word[i])

Project = ['例句', '断言', '等级', '类型']
for i in range(len(Project)):
    sheet.write(0,i,Project[i])
#
#
book.save("data/repeat100.xlsx")