from openpyxl import Workbook
from openpyxl import load_workbook
import difflib
import heapq

if __name__ == "__main__":
    # read from the terminal:
    command = (list(input("Please enter k and the string you want to compare").split()))
    k = command[0]  # the first is topk
    lens = len(command)


    # getExcel:
    # 实例化
    wb = Workbook()
    # 激活 worksheet
    ws = wb.active

    wb2 = load_workbook('测试.xlsx')

    sheet = wb2["关联关系"]
    max_row = sheet.max_row
    max_col = sheet.max_column

    # creat Array:
    boardC = ['s' for i in range(max_row)]
    boardA = ['s' for i in range(max_row)]
    boardB = ['s' for i in range(max_row)]
    similarities_of_data_element = [0 for i in range(max_row)]

    # read the data element:
    count = 0
    for row in sheet.rows:
        boardC[count] = row[2].value
        count += 1

    # only have : topk \ B \ D
    if lens == 3:
        for i in range(0, max_row):
            similarities_of_data_element[i] = difflib.SequenceMatcher(None, command[2], boardC[i]).ratio()
    # only have : topk \ A \ B \ D
    if lens == 4:
        for i in range(0, max_row):
            similarities_of_data_element[i] = difflib.SequenceMatcher(None, command[3], boardC[i]).ratio()

    # choose the top50percent:
    num_of_chosen = int(k)
    num_dict = {}
    for i in range(len(similarities_of_data_element)):
        num_dict[i] = similarities_of_data_element[i]
    res_list = sorted(num_dict.items(), key=lambda e: e[1])
    while similarities_of_data_element[max_row - num_of_chosen - 1] == similarities_of_data_element[max_row - num_of_chosen] :
        num_of_chosen = num_of_chosen + 1
    data_element_largestTopk_index = list([one[0] for one in res_list[::-1][:num_of_chosen]])

    print("num of chosen is ", num_of_chosen)

    ###

    for i in range(0, num_of_chosen):
        print(data_element_largestTopk_index[i], end=' ')
        print(similarities_of_data_element[data_element_largestTopk_index[i]])
    print()
    print()

    ###

    # compare the largest similarities' topk:
    similarities_of_institute_name = [0 for i in range(0, num_of_chosen)]
    similarities_of_table_Chinese_name = [0 for i in range(0, num_of_chosen)]

    # read the institute name/A:
    count = 0
    for i in range(0, max_row):
        boardA[i] = None
    for row in sheet.rows:
        if row[0].value is not None:
            boardA[count] = row[0].value
        count += 1
    # only have : topk \ B \ D
    if lens == 3:
        for i in range(0, num_of_chosen):
            if boardA[data_element_largestTopk_index[i]] is not None:
                similarities_of_institute_name[i] = 0
            else:
                similarities_of_institute_name[i] = 1
    # only have : topk \ A \ B \ D
    elif lens == 4:
        for i in range(0, num_of_chosen):
            if boardA[data_element_largestTopk_index[i]] is not None:
                similarities_of_institute_name[i] = difflib.SequenceMatcher(None, command[1], boardA[
                    data_element_largestTopk_index[i]]).ratio()
            else:
                similarities_of_institute_name[i] = 0

    # read the table Chinese name/B:
    count = 0
    for i in range(0, max_row):
        boardB[i] = None
    for row in sheet.rows:
        boardB[count] = row[1].value
        count += 1
    for i in range(0, num_of_chosen):
        similarities_of_table_Chinese_name[i] = difflib.SequenceMatcher(None, command[2], boardB[
            data_element_largestTopk_index[i]]).ratio()

    # calculate the similarities:
    similarities = [0 for i in range(0, num_of_chosen)]
    for i in range(0, num_of_chosen):
        similarities[i] = (similarities_of_table_Chinese_name[i] + similarities_of_institute_name[i] + similarities_of_data_element[data_element_largestTopk_index[i]]) / 3

    ###

    for i in range(0, num_of_chosen):
        print(data_element_largestTopk_index[i], end=' ')
        print(similarities[i])
    print()
    print()
    ###

    # rank the topk:
    num_of_topk = int(k)
    num_dict = {}
    for i in range(len(similarities)):
        num_dict[i] = similarities[i]
    res_list = sorted(num_dict.items(), key=lambda e: e[1])
    largestTopk_index = list([one[0] for one in res_list[::-1][:num_of_chosen]])

    # print:
    if num_of_topk > num_of_chosen:
        print("Please enter a smaller k! NO MORE THAN", num_of_chosen)
    else:
        for i in range(0, num_of_topk):
            print(largestTopk_index[i] + 1, boardA[largestTopk_index[i]], boardB[largestTopk_index[i]], boardC[largestTopk_index[i]], similarities[largestTopk_index[i]])


