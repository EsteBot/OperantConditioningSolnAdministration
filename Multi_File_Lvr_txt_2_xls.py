# pysimpleGUI for conversion of multiple Med Associate text files into one Excel file

# Import required libraries
import PySimpleGUI as sg
import xlwt
import os
from tempfile import TemporaryFile
from pathlib import Path
book = xlwt.Workbook()

# validate that the file paths are entered correctly
def is_valid_path(filepath):
    if filepath and Path(filepath).exists():
        return True
    sg.popup_error("A selected file path is incorrect or has been left empty.")
    return False

# window appears when the program successfully completes
def nom_window():
    layout = [[sg.Text("\n"
    " All Systems Nominal  \n"
    "\n"
    "")]]
    window = sg.Window((""), layout, modal=True)
    choice = None
    while True:
        event, values = window.read()
        if event == "Exit" or event == sg.WIN_CLOSED:
            break
    window.close()
    
# Define the location of the directory
def Lever_Presses_txt_to_Excel(input_folder, output_folder):
    name = Path(output_folder)

    # Change the directory
    os.chdir(input_folder)

    book = xlwt.Workbook()
    sheet1 = book.add_sheet('sheet1')

    # Variable to set the initial Bin start column
    strtbincol = 11

    # Variable to set the initial Bin start row
    strtdatarow = 1

    # Function to read files in defined path
    def read_files(file_path):
        with open(file_path, 'r') as file:
            (file.read())
            
    # creation of a maximum value for the progress bar function
    input_folder  = values["-IN-"]
    prog_bar_max_val = 1
    os.chdir(input_folder)
    for i in os.listdir():
        prog_bar_max_val += 1
    max = prog_bar_max_val

    # iterate over all the files in the directory 
    prog_bar_update_val = 0
    for file in os.listdir():
        prog_bar_update_val += 1
        #print("Files Complied: "+str(prog_bar_update_val))

        # records progress by updating prog bar with each file compiled
        window["-Progress_BAR-"].update(max = max, current_count=int(prog_bar_update_val))
        # check whether files are text files
        if file.endswith('.txt'):
            # Create the file path of the particular file
            file_path =f"{input_folder}/{file}"
            # Create new text files without spaces
            output=""
            with open(file_path) as f:
                for line in f:
                    if not line.isspace():
                        output+=line
                    f = open(file_path,"w")
                    f.write(output)
            read_files(file_path)
            
            # Variable to set a new start of Bin columns for each animal spacing them 21 columns apart
            #strtbincol += 21

            # Variable to set a new start of a data row for each animal spacing them 1 row apart
            #strtdatarow += 1

    sheet1.write(0,0,   input_folder)
    sheet1.write(1,0,  'Run Date')
    sheet1.write(1,1,  'Med Prgm')
    sheet1.write(1,2,  'Rat ID')
    sheet1.write(1,3,  'Sex')
    sheet1.write(1,4,  'Tx')
    sheet1.write(1,5,  'Box ID')
    sheet1.write(1,6,  'Grp ID')
    sheet1.write(1,7,  'R Lvr')
    sheet1.write(1,8,  'L Lvr')
    sheet1.write(1,9,  'Tot Rwd')
    sheet1.write(1,10, 'Tot Time')

    # Bin creation for 90min of lever pressing behavior. 18 5min Bins = 90min.
    for col in range(11,29,1):
        binnum = str(col - 10)
        sheet1.write(1,col,'Bin'+binnum)

    for file in os.listdir():
        strtdatarow += 1
        file_path =f"{input_folder}/{file}"
        medtxtfile = open(file_path)
        content = medtxtfile.readlines()
        
        RunDate_line = content[1]
        RDS = slice(12,20)
        sheet1.write(strtdatarow,0, RunDate_line[RDS])

        Exp_ID = content[9]
        EDS = slice(5,9)
        sheet1.write(strtdatarow,1, Exp_ID[EDS])

        RatID_line = content[3]
        IDS = slice(9,12)
        sheet1.write(strtdatarow,2, RatID_line[IDS]) 

        Box_line = content[6]
        BS = slice(5,8)
        sheet1.write(strtdatarow,5, Box_line[BS])

        Grp_line = content[5]
        GPS = slice(7,9)
        sheet1.write(strtdatarow,6, Grp_line[GPS])
        
        TotRpress_line = content[25]
        TPS = slice(5,13)
        sheet1.write(strtdatarow,7,float(TotRpress_line[TPS]))

        TotLpress_line = content[19]
        TPS = slice(5,13)
        sheet1.write(strtdatarow,8,float(TotLpress_line[TPS]))

        TotRwd_line = content[23]
        TPS = slice(5,13)
        sheet1.write(strtdatarow,9,float(TotRwd_line[TPS]))

        TotrunTime_line = content[27]
        TTS = slice(5,13)
        sheet1.write(strtdatarow,10,float(TotrunTime_line[TTS]))

        if (Exp_ID[EDS]) == "L_FR":  
            medtxtfile = open(file_path)
            specific_linesA = list(range(35,54))
            L_Lever_Press_List = []

            for pos, A_num in enumerate(medtxtfile):
                A_num = A_num.rstrip()
                if pos in specific_linesA:
                    
                    slicedp1 = A_num [12:18]
                    slicedp2 = A_num [25:31]
                    slicedp3 = A_num [38:44]
                    slicedp4 = A_num [51:57]
                    slicedp5 = A_num [64:70]
                    
                    if (float(slicedp1) > 0):
                        L_Lever_Press_List.append(float(slicedp1))
                        
                    if (float(slicedp2) > 0):
                        L_Lever_Press_List.append(float(slicedp2))
                        
                    if (float(slicedp3) > 0):
                        L_Lever_Press_List.append(float(slicedp3))
                    
                    if (float(slicedp4) > 0):
                        L_Lever_Press_List.append(float(slicedp4))
                        
                    if (float(slicedp5) > 0):
                        L_Lever_Press_List.append(float(slicedp5))

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
                
                for i, e in enumerate(L_bin_list,start=strtbincol):
                    sheet1.write(strtdatarow,i,e)
                    
            L_binthere(L_Lever_Press_List)

        else:
            medtxtfile = open(file_path)
            specific_linesB = list(range(57,76))
            R_Lever_Press_List = []

            for pos, B_num in enumerate(medtxtfile):
                B_num = B_num.rstrip()
                if pos in specific_linesB:
                    
                    slicedp1 = B_num [12:18]
                    slicedp2 = B_num [25:31]
                    slicedp3 = B_num [38:44]
                    slicedp4 = B_num [51:57]
                    slicedp5 = B_num [64:70]
                    
                    if (float(slicedp1) > 0):
                        R_Lever_Press_List.append(float(slicedp1))
                        
                    if (float(slicedp2) > 0):
                        R_Lever_Press_List.append(float(slicedp2))
                        
                    if (float(slicedp3) > 0):
                        R_Lever_Press_List.append(float(slicedp3))
                        
                    if (float(slicedp4) > 0):
                        R_Lever_Press_List.append(float(slicedp4))
                        
                    if (float(slicedp5) > 0):
                        R_Lever_Press_List.append(float(slicedp5))

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
                
                for i, e in enumerate(R_bin_list,start=strtbincol):
                    sheet1.write(strtdatarow,i,e)
                    
            R_binthere(R_Lever_Press_List)
        
    name = input_folder+'.xls'
    book.save(name)
    book.save(TemporaryFile())   

    # last prog bar addition indicating the end of the program run
    window["-Progress_BAR-"].update(current_count=int(prog_bar_update_val +1))

    # window telling the user the program functioned correctly
    nom_window()   

# creation of a maximum value for the progress bar function
def bar_max(input_folder):
    prog_bar_max_val = 0
    os.chdir(input_folder)
    for i in os.listdir():
        prog_bar_max_val += 1
    print(prog_bar_max_val)

# main GUI creation and GUI elements
sg.theme('DarkBlue7')

layout = [
    [sg.Text("Select the folder containing the\n"
             "Med Associates text(.txt) files                   \n" 
             "to be converted to an Excel file."),
    sg.Input(key="-IN-"),
    sg.FolderBrowse()],

    [sg.Text("Select a file to store the new Excel file.\n"
                "Data will be copied & transferred to this file."),
    sg.Input(key="-OUT-"),
    sg.FolderBrowse()],

    [sg.Exit(), sg.Button("Press to convert the .txt file into an .xsl file"), 
    sg.Text("eBot's progress..."),
    sg.ProgressBar(20, orientation='horizontal', size=(15,10), 
                border_width=4, bar_color=("Blue", "Grey"),
                key="-Progress_BAR-")]
    
]

# create the window
window = sg.Window("Welcom to eBot's Lever Press .txt to .xsl converter!", layout)

# create an event loop
while True:
    event, values = window.read()
    # end program if user closes window
    if event == "Exit" or event == sg.WIN_CLOSED:
        break
    if event == "Press to convert the .txt file into an .xsl file":
        # check file selections are valid
        if (is_valid_path(values["-IN-"])) and (is_valid_path(values["-OUT-"])):

            Lever_Presses_txt_to_Excel(
            input_folder  = values["-IN-"],
            output_folder = values["-OUT-"])   

window.close