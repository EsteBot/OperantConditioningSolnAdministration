import xlwt
from xlwt import Workbook

wb = Workbook()

lever_press_list = []

fhand = open('2021-07-19_09h32m_Subject 45.txt')
count = 0
for line in fhand:
    line = line.rstrip()
    #if line.find('0:  ') == -1 and line.find('5:  ') == -1 and line.find('100:') != -1: continue
    print(line)

fhand = open('2021-07-19_09h32m_Subject 45.txt')
for line in fhand:
    line = line.rstrip()
    if line.startswith('A:'):
            print("Left  Lever Press Time Stamps")
    elif line.startswith('B:'):
            print("Right Lever Press Time Stamps")
    elif line.startswith('     0:')\
        or line.startswith('     5:') or line.startswith('    10:')\
        or line.startswith('    15:') or line.startswith('    20:')\
        or line.startswith('    25:') or line.startswith('    30:')\
        or line.startswith('    35:') or line.startswith('    40:')\
        or line.startswith('    45:') or line.startswith('    50:')\
        or line.startswith('    55:') or line.startswith('    60:')\
        or line.startswith('    65:') or line.startswith('    70:')\
        or line.startswith('    75:') or line.startswith('    80:')\
        or line.startswith('    85:') or line.startswith('    90:')\
        or line.startswith('    95:'):
        #print(line)
        slicedp1 = line [12:18]
        slicedp2 = line [25:31]
        slicedp3 = line [38:44]
        slicedp4 = line [51:57]
        slicedp5 = line [64:70]
        
        if (float(slicedp1) > 0) or (float(slicedp2) > 0)\
           or (float(slicedp3) > 0) or (float(slicedp4) > 0)\
               or (float(slicedp5) > 0):
         print(slicedp1)
         print(slicedp2)
         print(slicedp3)
         print(slicedp4)
         print(slicedp5)
         lever_press_list.append(slicedp1)
         lever_press_list.append(slicedp2)
         lever_press_list.append(slicedp3)
         lever_press_list.append(slicedp4)
         lever_press_list.append(slicedp5)
         #print(lever_press_list)
        #sheet1 = wb.add_sheet('Sheet 1')
        #sheet1.write(row,col, data, style)
        #sheet1.write(1, 0, slicedp1)
        #sheet1.write(2, 0, slicedp2)
        #sheet1.write(3, 0, slicedp3)
        #sheet1.write(4, 0, slicedp4)

        #wb.save('lever_sample_data1.xls')

