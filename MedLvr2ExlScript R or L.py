import xlwt
import sys
from tempfile import TemporaryFile

def LOCOLvr2ExlScript():

    try:
        TextFile = input("Enter the name of the Med Associates text file to be converted: ")

        output=""
        with open(TextFile+'.txt') as f:
            for line in f:
                if not line.isspace():
                    output+=line
                f = open("output.txt","w")
                f.write(output)

    except FileNotFoundError:
        print("\n"
        "!UNABLE TO FIND MED ASSOC. FILE!\n"
        "Make sure the Med Associates text file to be converted matches the inputted file name exactly.\n"
        "Copying and pasting the name is suggested. Ensure the file extension '.txt' is not in the file name.\n"
        "The Med Assoc. text file must be in the same location from which this python program is running.")

        k = input("\n"
                  "Enter 'e' to try re-entering inputs and run the MedTxt2Exl\n"
                  "program again or enter any other key to exit the program. ")
        if k == 'e':
            LOCOLvr2ExlScript()
        else: quit()

    ExcelFile = input("Enter the name of the Excel file to contain the converted data: ")

    book = xlwt.Workbook()
    sheet1 = book.add_sheet('sheet1')
    sheet1.write(0,0, ExcelFile)
    sheet1.write(1,0, 'Run Date')
    sheet1.write(1,1, 'Exp ID')
    sheet1.write(1,2, 'Rat ID')
    sheet1.write(1,3, 'Box ID')
    sheet1.write(1,4, 'Grp ID')
    sheet1.write(1,5, 'R Lvr')
    sheet1.write(1,6, 'L Lvr')
    sheet1.write(1,7, 'Tot Rwd')
    sheet1.write(1,8, 'Tot Time')

    medtxtfile = open('output.txt')
    content = medtxtfile.readlines()

    RunDate_line = content[1]
    RDS = slice(12,20)
    sheet1.write(2,0, RunDate_line[RDS])

    Exp_ID = content[9]
    EDS = slice(5,9)
    sheet1.write(2,1, Exp_ID[EDS])

    RatID_line = content[3]
    IDS = slice(9,12)
    sheet1.write(2,2, RatID_line[IDS]) 

    Box_line = content[6]
    BS = slice(5,8)
    sheet1.write(2,3, Box_line[BS])

    Grp_line = content[5]
    GPS = slice(7,9)
    sheet1.write(2,4, Grp_line[GPS])
    
    TotRpress_line = content[25]
    TPS = slice(5,13)
    sheet1.write(2,5,float(TotRpress_line[TPS]))

    TotLpress_line = content[19]
    TPS = slice(5,13)
    sheet1.write(2,6,float(TotLpress_line[TPS]))

    TotRwd_line = content[23]
    TPS = slice(5,13)
    sheet1.write(2,7,float(TotRwd_line[TPS]))

    TotrunTime_line = content[27]
    TTS = slice(5,13)
    sheet1.write(2,8,float(TotrunTime_line[TTS]))
        
    # Bin creation for 90min of lever pressing behavior. 18 5min Bins = 90min.
    for col in range(9,27,1):
        binnum = str(col - 8)
        sheet1.write(1,col,'Bin'+binnum)
    if (Exp_ID[EDS]) == "L_FR":  
        medtxtfile = open('output.txt')
        specific_linesA = list(range(35,94))
        L_Lever_Press_List = []

        print("L Lever Time Stamps (seconds)")
        for pos, A_num in enumerate(medtxtfile):
            A_num = A_num.rstrip()
            if pos in specific_linesA:
                #print(A_num)
                slicedp1 = A_num [12:18]
                slicedp2 = A_num [25:31]
                slicedp3 = A_num [38:44]
                slicedp4 = A_num [51:57]
                slicedp5 = A_num [64:70]
                
                if (float(slicedp1) > 0):
                    L_Lever_Press_List.append(float(slicedp1))
                    print(slicedp1)
                if (float(slicedp2) > 0):
                    L_Lever_Press_List.append(float(slicedp2))
                    print(slicedp2)
                if (float(slicedp3) > 0):
                    L_Lever_Press_List.append(float(slicedp3))
                    print(slicedp3)
                if (float(slicedp4) > 0):
                    L_Lever_Press_List.append(float(slicedp4))
                    print(slicedp4)
                if (float(slicedp5) > 0):
                    L_Lever_Press_List.append(float(slicedp5))
                    print(slicedp5)
                
        print(L_Lever_Press_List)  

        #for i, e in enumerate(L_Lever_Press_List, start=2):
            #sheet1.write(i,8,e)
        def L_binthere(L_Lever_Press_List):
            L_bin1  = 0
            L_bin2  = 0
            L_bin3  = 0
            L_bin4  = 0
            L_bin5  = 0
            L_bin6  = 0
            L_bin7  = 0
            L_bin8  = 0
            L_bin9  = 0
            L_bin10 = 0
            L_bin11 = 0
            L_bin12 = 0
            L_bin13 = 0
            L_bin14 = 0
            L_bin15 = 0
            L_bin16 = 0
            L_bin17 = 0
            L_bin18 = 0

            for timestamp in L_Lever_Press_List:
                if timestamp < 300:
                    L_bin1 +=1
                if 300 <= timestamp < 600:
                    L_bin2 +=1
                if 600 <= timestamp < 900:
                    L_bin3 +=1
                if 900 <= timestamp < 1200:
                    L_bin4 +=1
                if 1200 <= timestamp < 1500:
                    L_bin5 +=1
                if 1500 <= timestamp < 1800:
                    L_bin6 +=1
                if 1800 <= timestamp < 2100:
                    L_bin7 +=1
                if 2100 <= timestamp < 2400:
                    L_bin8 +=1
                if 2400 <= timestamp < 2700:
                    L_bin9 +=1
                if 2700 <= timestamp < 3200:
                    L_bin10 +=1
                if 3200 <= timestamp < 3500:
                    L_bin11 +=1
                if 3500 <= timestamp < 3800:
                    L_bin12 +=1
                if 3800 <= timestamp < 4100:
                    L_bin13 +=1
                if 4100 <= timestamp < 4400:
                    L_bin14 +=1
                if 4400 <= timestamp < 4700:
                    L_bin15 +=1
                if 4700 <= timestamp < 5000:
                    L_bin16 +=1
                if 5000 <= timestamp < 5300:
                    L_bin17 +=1
                if 5300 <= timestamp < 5600:
                    L_bin18 +=1

            L_bin_list = [L_bin1,L_bin2,L_bin3,L_bin4,L_bin5,L_bin6,L_bin7,L_bin8,L_bin9,
            L_bin10,L_bin11,L_bin12,L_bin13,L_bin14,L_bin15,L_bin16,L_bin17,L_bin18]

            for i, e in enumerate(L_bin_list,start=9):
                sheet1.write(2,i,e)

        L_binthere(L_Lever_Press_List)

    else:
        medtxtfile = open('output.txt')
        #content = medtxtfile.readlines()
        specific_linesB = list(range(97,119))
        R_Lever_Press_List = []
        print("R Lever Time Stamps (seconds)")
        for pos, B_num in enumerate(medtxtfile):
            B_num = B_num.rstrip()
            if pos in specific_linesB:
                #print(B_num)
                slicedp1 = B_num [12:18]
                slicedp2 = B_num [25:31]
                slicedp3 = B_num [38:44]
                slicedp4 = B_num [51:57]
                slicedp5 = B_num [64:70]
                
                if (float(slicedp1) > 0):
                    R_Lever_Press_List.append(float(slicedp1))
                    print(slicedp1)
                if (float(slicedp2) > 0):
                    R_Lever_Press_List.append(float(slicedp2))
                    print(slicedp2)
                if (float(slicedp3) > 0):
                    R_Lever_Press_List.append(float(slicedp3))
                    print(slicedp3)
                if (float(slicedp4) > 0):
                    R_Lever_Press_List.append(float(slicedp4))
                    print(slicedp4)
                if (float(slicedp5) > 0):
                    R_Lever_Press_List.append(float(slicedp5))
                    print(slicedp5)
                
        print(R_Lever_Press_List) 

        #for i, e in enumerate(R_Lever_List, start=2):
            #sheet1.write(i,9,e)
        def R_binthere(R_Lever_Press_List):
            R_bin1  = 0
            R_bin2  = 0
            R_bin3  = 0
            R_bin4  = 0
            R_bin5  = 0
            R_bin6  = 0
            R_bin7  = 0
            R_bin8  = 0
            R_bin9  = 0
            R_bin10 = 0
            R_bin11 = 0
            R_bin12 = 0
            R_bin13 = 0
            R_bin14 = 0
            R_bin15 = 0
            R_bin16 = 0
            R_bin17 = 0
            R_bin18 = 0

            for timestamp in R_Lever_Press_List:
                if timestamp < 300:
                    R_bin1 +=1
                if 300 <= timestamp < 600:
                    R_bin2 +=1
                if 600 <= timestamp < 900:
                    R_bin3 +=1
                if 900 <= timestamp < 1200:
                    R_bin4 +=1
                if 1200 <= timestamp < 1500:
                    R_bin5 +=1
                if 1500 <= timestamp < 1800:
                    R_bin6 +=1
                if 1800 <= timestamp < 2100:
                    R_bin7 +=1
                if 2100 <= timestamp < 2400:
                    R_bin8 +=1
                if 2400 <= timestamp < 2700:
                    R_bin9 +=1
                if 2700 <= timestamp < 3200:
                    R_bin10 +=1
                if 3200 <= timestamp < 3500:
                    R_bin11 +=1
                if 3500 <= timestamp < 3800:
                    R_bin12 +=1
                if 3800 <= timestamp < 4100:
                    R_bin13 +=1
                if 4100 <= timestamp < 4400:
                    R_bin14 +=1
                if 4400 <= timestamp < 4700:
                    R_bin15 +=1
                if 4700 <= timestamp < 5000:
                    R_bin16 +=1
                if 5000 <= timestamp < 5300:
                    R_bin17 +=1
                if 5300 <= timestamp < 5600:
                    R_bin18 +=1

            R_bin_list = [R_bin1,R_bin2,R_bin3,R_bin4,R_bin5,R_bin6,R_bin7,R_bin8,R_bin9,
            R_bin10,R_bin11,R_bin12,R_bin13,R_bin14,R_bin15,R_bin16,R_bin17,R_bin18]

            for i, e in enumerate(R_bin_list,start=9):
                sheet1.write(2,i,e)

        R_binthere(R_Lever_Press_List)

    try:
        name = ExcelFile+'.xls'
        book.save(name)
        book.save(TemporaryFile())

    except FileNotFoundError:
        print("!UNABLE TO CREATE EXCEL FILE!\n"
        "Ensure the chosen Excel file name for the exported data does not contain any unaccepted\n"
        "characters as part of the file name (e.g. '/' or '#') and is a valid Excel file name.")

        k = input("\n"
                "Enter 'e' to try re-entering inputs and run the MedTxt2Exl\n"
                "program again or enter any other key to exit the program. ")
        if k == 'e':
            LOCOLvr2ExlScript()
        else: quit()

    r = input("\n"
            "All systems nominal.\n"
            "\n"
            "The new "+ name+ " file is in the same location as the MedTxt2Exl program & original Med Assoc. file.\n"
            "An additional file was created, titled 'output' containing a version of the Med Assoc. file without spaces.\n"
            "Enter 'e' to convert another Med Assoc. file or enter any other key to exit the MedTxt2Exl program. ")
    if r == 'e':
            LOCOLvr2ExlScript()
    else: quit()

LOCOLvr2ExlScript()