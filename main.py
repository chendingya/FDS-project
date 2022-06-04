from openpyxl import Workbook
from openpyxl import load_workbook
import difflib
import heapq

if __name__ == "__main__":
    # getExcel:
    # 实例化
    wb = Workbook()
    # 激活 worksheet
    ws = wb.active

    wb2 = load_workbook('测试.xlsx')
    print(wb2.sheetnames)

    sheet = wb2["关联关系"]
    max_row = sheet.max_row
    max_col = sheet.max_column
    print(max_row, " ", max_col)

    # creat Array:
    board = ['s' for i in range(max_row)]
    similarities = [0 for i in range(max_row)]
    count = 0
    for row in sheet.rows:
        board[count] = row[3].value
        count += 1

    # read from the terminal:
    command = (list(input("Please enter k and the string you want to compare").split()))
    k = command[0]  # the first is topk
    lens = len(command)
    # only have : topk \ B \ D
    if lens == 3:
        for i in range(0, max_row):
            similarities[i] = difflib.SequenceMatcher(None, command[2], board[i]).ratio()
    # only have : topk \ A \ B \ D
    if lens == 4:
        for i in range(0, max_row):
            similarities[i] = difflib.SequenceMatcher(None, command[3], board[i]).ratio()


    # choose the top10k:
    num_of_chosen = 10 * int(k)
    if num_of_chosen > max_row:
        num_of_chosen = max_row
    largestTopk = heapq.nlargest(num_of_chosen, similarities)

    for i in range(0, num_of_chosen):
        print(largestTopk[i])




