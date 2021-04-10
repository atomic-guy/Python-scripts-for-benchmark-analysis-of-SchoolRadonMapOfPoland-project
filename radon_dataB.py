import pandas as pd
import os
from radon_statistics import getListofExcelFiles
import matplotlib.pyplot as plt
import numpy as np

def collectDataPartB(ListofExcelFiles):
    '''
    Collects all results of <benchmark part 2>. Data from every Excel file is stored in a dictionary.
            Parameters:
                    ListofExcelFiles (list of str): List of .xls files with full path
            Returns:
                    partaBData (pandas DataFrame):  DataFrame created from List of dictionaries (listPartB)
                                                    Every dictionary in listPartB collects data from one Excel file
    '''
    listPartB = []
    for file in ListofExcelFiles:
        dict_with_data_B = {}
        
        file_name = os.path.basename(file)
        id = int(file_name[0:2])
        dict_with_data_B['id'] = id
        
        #------------------------------------------------------
        #general info
        general_info_sheet = pd.read_excel(file, sheet_name="Informacje ogólne")
        
        school = general_info_sheet.at[1,"Unnamed: 2"]
        city = general_info_sheet.at[2,"Unnamed: 2"]
        school_full = school + " - " + city
        dict_with_data_B['school'] = school_full
        
        teacher = general_info_sheet.at[4,"Unnamed: 2"]
        dict_with_data_B['teacher'] = teacher
        
        students = general_info_sheet.loc[6:,['Unnamed: 2']].values.tolist()
        students_flat_list = [item for sublist in students for item in sublist]
        dict_with_data_B['students'] = students_flat_list
        
        #------------------------------------------------------
        #data - part B
        partB_sheet = pd.read_excel(file, sheet_name="Część 2")
        pd.set_option("display.max_rows", None, "display.max_columns", None)
        #print(partB_sheet)
        all_done = partB_sheet.at[0,"Unnamed: 5"]
        dict_with_data_B['all done'] = all_done

        #Calibration - Photo 1
        cal_photo1_n = float(partB_sheet.at[9,"Unnamed: 5"])
        dict_with_data_B['cal_photo1_n'] = cal_photo1_n
        cal_photo1_u_n = float(partB_sheet.at[10,"Unnamed: 5"])
        dict_with_data_B['cal_photo1_u_n'] = cal_photo1_u_n
        cal_photo1_S = float(partB_sheet.at[12,"Unnamed: 5"])
        dict_with_data_B['cal_photo1_S'] = cal_photo1_S
        cal_photo1_u_S = float(partB_sheet.at[13,"Unnamed: 5"])
        dict_with_data_B['cal_photo1_u_S'] = cal_photo1_u_S
        cal_photo1_ro = float(partB_sheet.at[15,"Unnamed: 5"])
        dict_with_data_B['cal_photo1_ro'] = cal_photo1_ro
        cal_photo1_u_ro = float(partB_sheet.at[16,"Unnamed: 5"])
        dict_with_data_B['cal_photo1_u_ro'] = cal_photo1_u_ro

        #Calibration - Photo 2
        cal_photo2_n = float(partB_sheet.at[9,"Unnamed: 7"])
        dict_with_data_B['cal_photo2_n'] = cal_photo2_n
        cal_photo2_u_n = float(partB_sheet.at[10,"Unnamed: 7"])
        dict_with_data_B['cal_photo2_u_n'] = cal_photo2_u_n
        cal_photo2_S = float(partB_sheet.at[12,"Unnamed: 7"])
        dict_with_data_B['cal_photo2_S'] = cal_photo2_S
        cal_photo2_u_S = float(partB_sheet.at[13,"Unnamed: 7"])
        dict_with_data_B['cal_photo2_u_S'] = cal_photo2_u_S
        cal_photo2_ro = float(partB_sheet.at[15,"Unnamed: 7"])
        dict_with_data_B['cal_photo2_ro'] = cal_photo2_ro
        cal_photo2_u_ro = float(partB_sheet.at[16,"Unnamed: 7"])
        dict_with_data_B['cal_photo2_u_ro'] = cal_photo2_u_ro

        #Calibration - Photo 3
        cal_photo3_n = float(partB_sheet.at[9,"Unnamed: 9"])
        dict_with_data_B['cal_photo3_n'] = cal_photo3_n
        cal_photo3_u_n = float(partB_sheet.at[10,"Unnamed: 9"])
        dict_with_data_B['cal_photo3_u_n'] = cal_photo3_u_n
        cal_photo3_S = float(partB_sheet.at[12,"Unnamed: 9"])
        dict_with_data_B['cal_photo3_S'] = cal_photo3_S
        cal_photo3_u_S = float(partB_sheet.at[13,"Unnamed: 9"])
        dict_with_data_B['cal_photo3_u_S'] = cal_photo3_u_S
        cal_photo3_ro = float(partB_sheet.at[15,"Unnamed: 9"])
        dict_with_data_B['cal_photo3_ro'] = cal_photo3_ro
        cal_photo3_u_ro = float(partB_sheet.at[16,"Unnamed: 9"])
        dict_with_data_B['cal_photo3_u_ro'] = cal_photo3_u_ro

        #Calibration - parameters
        cal_param_a = float(partB_sheet.at[21,"Unnamed: 7"])
        dict_with_data_B['cal_param_a'] = cal_param_a
        cal_param_u_a = float(partB_sheet.at[22,"Unnamed: 7"])
        dict_with_data_B['cal_param_u_a'] = cal_param_u_a
        cal_param_b = float(partB_sheet.at[24,"Unnamed: 7"])
        dict_with_data_B['cal_param_b'] = cal_param_b
        cal_param_u_b = float(partB_sheet.at[25,"Unnamed: 7"])
        dict_with_data_B['cal_param_u_b'] = cal_param_u_b

        #Photo 3
        photo3_n = float(partB_sheet.at[7,"Unnamed: 16"])
        dict_with_data_B['photo3_n'] = photo3_n
        photo3_u_n = float(partB_sheet.at[8,"Unnamed: 16"])
        dict_with_data_B['photo3_u_n'] = photo3_u_n
        photo3_S = float(partB_sheet.at[10,"Unnamed: 16"])
        dict_with_data_B['photo3_S'] = photo3_S
        photo3_u_S = float(partB_sheet.at[11,"Unnamed: 16"])
        dict_with_data_B['photo3_u_S'] = photo3_u_S
        photo3_ro = float(partB_sheet.at[13,"Unnamed: 16"])
        dict_with_data_B['photo3_ro'] = photo3_ro
        photo3_u_ro = float(partB_sheet.at[14,"Unnamed: 16"])
        dict_with_data_B['photo3_u_ro'] = photo3_u_ro
        photo3_Bq = float(partB_sheet.at[16,"Unnamed: 16"])
        dict_with_data_B['photo3_Bq'] = photo3_Bq
        photo3_u_Bq = float(partB_sheet.at[17,"Unnamed: 16"])
        dict_with_data_B['photo3_u_Bq'] = photo3_u_Bq
        
        #Photo 4
        photo4_n = float(partB_sheet.at[7,"Unnamed: 18"])
        dict_with_data_B['photo4_n'] = photo4_n
        photo4_u_n = float(partB_sheet.at[8,"Unnamed: 18"])
        dict_with_data_B['photo4_u_n'] = photo4_u_n
        photo4_S = float(partB_sheet.at[10,"Unnamed: 18"])
        dict_with_data_B['photo4_S'] = photo4_S
        photo4_u_S = float(partB_sheet.at[11,"Unnamed: 18"])
        dict_with_data_B['photo4_u_S'] = photo4_u_S
        photo4_ro = float(partB_sheet.at[13,"Unnamed: 18"])
        dict_with_data_B['photo4_ro'] = photo4_ro
        photo4_u_ro = float(partB_sheet.at[14,"Unnamed: 18"])
        dict_with_data_B['photo4_u_ro'] = photo4_u_ro
        photo4_Bq = float(partB_sheet.at[16,"Unnamed: 18"])
        dict_with_data_B['photo4_Bq'] = photo4_Bq
        photo4_u_Bq = float(partB_sheet.at[17,"Unnamed: 18"])
        dict_with_data_B['photo4_u_Bq'] = photo4_u_Bq
        
        #print(dict_with_data_B)
        #print('\n')
        listPartB.append(dict_with_data_B)
    partBData = pd.DataFrame(listPartB)
    return partBData


def partB_Cal_Photo1_NumberOfTracks(DirectoryPath, partBData):
    '''
    Creates a graph of number of tracs on 'calibration photo1' vs. the number of the student group.
            Parameters:
                    DirectoryPath (str): Path to folder where graph will be stored
                    partaBData (pandas DataFrame):  DataFrame with results of <bemchmark part2> collected from Excel files.
    '''
    x = partBData["id"]
    y = partBData["cal_photo1_n"]
    partBData['cal_photo1_u_n'] = partBData['cal_photo1_u_n'].fillna(0)
    yerror = partBData["cal_photo1_u_n"]
    filename = DirectoryPath + "B_kal_zdjecie1_n.png"
    #create line
    y_line = [y[0]] * len(y)

    plt.figure(figsize=(16,4))
    plt.errorbar(x[1:], y[1:], yerr=yerror[1:], capsize=3, fmt='o')
    plt.errorbar(x[0], y[0], yerr=yerror[0], capsize=3, fmt='ro')
    plt.plot(x, y_line, color='r', linestyle='-', linewidth=0.8)
    plt.xlabel("Numer grupy", fontsize='large')
    plt.ylabel("Liczba śladów", fontsize='large')
    plt.title("Część 2, Zdjęcie kalibracyjne 1, Liczba śladów", fontsize='large')
    plt.grid(axis='x', linestyle='--', color='0.95')
    plt.grid(axis='y', linestyle='--', color='0.95')
    plt.xticks(x)
    plt.savefig(filename, dpi=600)
    plt.close() 
    return

def partB_Cal_Photo2_NumberOfTracks(DirectoryPath, partBData):
    '''
    Creates a graph of number of tracs on 'calibration photo2' vs. the number of the student group.
            Parameters:
                    DirectoryPath (str): Path to folder where graph will be stored
                    partaBData (pandas DataFrame):  DataFrame with results of <bemchmark part2> collected from Excel files.
    '''
    x = partBData["id"]
    y = partBData["cal_photo2_n"]
    partBData['cal_photo2_u_n'] = partBData['cal_photo2_u_n'].fillna(0)
    yerror = partBData["cal_photo2_u_n"]
    filename = DirectoryPath + "B_kal_zdjecie2_n.png"
    #create line
    y_line = [y[0]] * len(y)

    plt.figure(figsize=(16,4))
    plt.errorbar(x[1:], y[1:], yerr=yerror[1:], capsize=3, fmt='o')
    plt.errorbar(x[0], y[0], yerr=yerror[0], capsize=3, fmt='ro')
    plt.plot(x, y_line, color='r', linestyle='-', linewidth=0.8)
    plt.xlabel("Numer grupy", fontsize='large')
    plt.ylabel("Liczba śladów", fontsize='large')
    plt.title("Część 2, Zdjęcie kalibracyjne 2, Liczba śladów", fontsize='large')
    plt.grid(axis='x', linestyle='--', color='0.95')
    plt.grid(axis='y', linestyle='--', color='0.95')
    plt.xticks(x)
    plt.savefig(filename, dpi=600)
    plt.close() 
    return

def partB_Cal_Photo3_NumberOfTracks(DirectoryPath, partBData):
    '''
    Creates a graph of number of tracs on 'calibration photo3' vs. the number of the student group.
            Parameters:
                    DirectoryPath (str): Path to folder where graph will be stored
                    partaBData (pandas DataFrame):  DataFrame with results of <bemchmark part2> collected from Excel files.
    '''
    x = partBData["id"]
    y = partBData["cal_photo3_n"]
    partBData['cal_photo3_u_n'] = partBData['cal_photo3_u_n'].fillna(0)
    yerror = partBData["cal_photo3_u_n"]
    filename = DirectoryPath + "B_kal_zdjecie3_n.png"
    #create line
    y_line = [y[0]] * len(y)

    plt.figure(figsize=(16,4))
    plt.errorbar(x[1:], y[1:], yerr=yerror[1:], capsize=3, fmt='o')
    plt.errorbar(x[0], y[0], yerr=yerror[0], capsize=3, fmt='ro')
    plt.plot(x, y_line, color='r', linestyle='-', linewidth=0.8)
    plt.xlabel("Numer grupy", fontsize='large')
    plt.ylabel("Liczba śladów", fontsize='large')
    plt.title("Część 2, Zdjęcie kalibracyjne 3, Liczba śladów", fontsize='large')
    plt.grid(axis='x', linestyle='--', color='0.95')
    plt.grid(axis='y', linestyle='--', color='0.95')
    plt.xticks(x)
    plt.savefig(filename, dpi=600)
    plt.close() 
    return

def partB_Cal_Photo1_Surface(DirectoryPath, partBData):
    '''
    Creates a graph of surface area of 'calibration photo1' vs. the number of the student group.
            Parameters:
                    DirectoryPath (str): Path to folder where graph will be stored
                    partaBData (pandas DataFrame):  DataFrame with results of <bemchmark part2> collected from Excel files.
    '''
    x = partBData["id"]
    y = partBData["cal_photo1_S"]
    partBData['cal_photo1_u_S'] = partBData['cal_photo1_u_S'].fillna(0)
    yerror = partBData["cal_photo1_u_S"]
    filename = DirectoryPath + "B_kal_zdjecie1_S.png"
    #create line
    y_line = [y[0]] * len(y)

    plt.figure(figsize=(16,4))
    plt.yscale('log')
    plt.errorbar(x[1:], y[1:], yerr=yerror[1:], capsize=3, fmt='o')
    plt.errorbar(x[0], y[0], yerr=yerror[0], capsize=3, fmt='ro')
    plt.plot(x, y_line, color='r', linestyle='-', linewidth=0.8)
    plt.xlabel("Numer grupy", fontsize='large')
    plt.ylabel("Pole powierzchni [$cm^2$]", fontsize='large')
    plt.title("Część 2, Zdjęcie kalibracyjne 1, Pole powierzchni", fontsize='large')
    plt.grid(axis='x', linestyle='--', color='0.95')
    plt.grid(axis='y', linestyle='--', color='0.95')
    plt.xticks(x)
    plt.savefig(filename, dpi=600)
    plt.close() 
    return

def partB_Cal_Photo2_Surface(DirectoryPath, partBData):
    '''
    Creates a graph of surface area of 'calibration photo2' vs. the number of the student group.
            Parameters:
                    DirectoryPath (str): Path to folder where graph will be stored
                    partaBData (pandas DataFrame):  DataFrame with results of <bemchmark part2> collected from Excel files.
    '''
    x = partBData["id"]
    y = partBData["cal_photo2_S"]
    partBData['cal_photo2_u_S'] = partBData['cal_photo2_u_S'].fillna(0)
    yerror = partBData["cal_photo2_u_S"]
    filename = DirectoryPath + "B_kal_zdjecie2_S.png"
    #create line
    y_line = [y[0]] * len(y)

    plt.figure(figsize=(16,4))
    plt.yscale('log')
    plt.errorbar(x[1:], y[1:], yerr=yerror[1:], capsize=3, fmt='o')
    plt.errorbar(x[0], y[0], yerr=yerror[0], capsize=3, fmt='ro')
    plt.plot(x, y_line, color='r', linestyle='-', linewidth=0.8)
    plt.xlabel("Numer grupy", fontsize='large')
    plt.ylabel("Pole powierzchni [$cm^2$]", fontsize='large')
    plt.title("Część 2, Zdjęcie kalibracyjne 2, Pole powwierzchni", fontsize='large')
    plt.grid(axis='x', linestyle='--', color='0.95')
    plt.grid(axis='y', linestyle='--', color='0.95')
    plt.xticks(x)
    plt.savefig(filename, dpi=600)
    plt.close() 
    return

def partB_Cal_Photo3_Surface(DirectoryPath, partBData):
    '''
    Creates a graph of surface area of 'calibration photo3' vs. the number of the student group.
            Parameters:
                    DirectoryPath (str): Path to folder where graph will be stored
                    partaBData (pandas DataFrame):  DataFrame with results of <bemchmark part2> collected from Excel files.
    '''
    x = partBData["id"]
    y = partBData["cal_photo3_S"]
    partBData['cal_photo3_u_S'] = partBData['cal_photo3_u_S'].fillna(0)
    yerror = partBData["cal_photo3_u_S"]
    filename = DirectoryPath + "B_kal_zdjecie3_S.png"
    #create line
    y_line = [y[0]] * len(y)

    plt.figure(figsize=(16,4))
    plt.yscale('log')
    plt.errorbar(x[1:], y[1:], yerr=yerror[1:], capsize=3, fmt='o')
    plt.errorbar(x[0], y[0], yerr=yerror[0], capsize=3, fmt='ro')
    plt.plot(x, y_line, color='r', linestyle='-', linewidth=0.8)
    plt.xlabel("Numer grupy", fontsize='large')
    plt.ylabel("Pole powierzchni [$cm^2$]", fontsize='large')
    plt.title("Część 2, Zdjęcie kalibracyjne 3, Pole powierzchni", fontsize='large')
    plt.grid(axis='x', linestyle='--', color='0.95')
    plt.grid(axis='y', linestyle='--', color='0.95')
    plt.xticks(x)
    plt.savefig(filename, dpi=600)
    plt.close() 
    return


def partB_Cal_Photo1_Rho(DirectoryPath, partBData):
    '''
    Creates a graph of dencity of traces of 'calibration photo1' vs. the number of the student group.
            Parameters:
                    DirectoryPath (str): Path to folder where graph will be stored
                    partaBData (pandas DataFrame):  DataFrame with results of <bemchmark part2> collected from Excel files.
    '''
    x = partBData["id"]
    y = partBData["cal_photo1_ro"]
    partBData['cal_photo1_u_ro'] = partBData['cal_photo1_u_ro'].fillna(0)
    yerror = partBData["cal_photo1_u_ro"]
    filename = DirectoryPath + "B_kal_zdjecie1_Rho.png"
    #create line
    y_line = [y[0]] * len(y)

    plt.figure(figsize=(16,4))
    plt.yscale('log')
    plt.errorbar(x[1:], y[1:], yerr=yerror[1:], capsize=3, fmt='o')
    plt.errorbar(x[0], y[0], yerr=yerror[0], capsize=3, fmt='ro')
    plt.plot(x, y_line, color='r', linestyle='-', linewidth=0.8)
    plt.xlabel("Numer grupy", fontsize='large')
    plt.ylabel("Gęstość śladów [$n/cm^2$]", fontsize='large')
    plt.title("Część 2, Zdjęcie kalibracyjne 1, Gęstość śladów", fontsize='large')
    plt.grid(axis='x', linestyle='--', color='0.95')
    plt.grid(axis='y', linestyle='--', color='0.95')
    plt.xticks(x)
    plt.savefig(filename, dpi=600)
    plt.close() 
    return

def partB_Cal_Photo2_Rho(DirectoryPath, partBData):
    '''
    Creates a graph of dencity of traces of 'calibration photo2' vs. the number of the student group.
            Parameters:
                    DirectoryPath (str): Path to folder where graph will be stored
                    partaBData (pandas DataFrame):  DataFrame with results of <bemchmark part2> collected from Excel files.
    '''
    x = partBData["id"]
    y = partBData["cal_photo2_ro"]
    partBData['cal_photo2_u_ro'] = partBData['cal_photo2_u_ro'].fillna(0)
    yerror = partBData["cal_photo2_u_ro"]
    filename = DirectoryPath + "B_kal_zdjecie2_Rho.png"
    #create line
    y_line = [y[0]] * len(y)

    plt.figure(figsize=(16,4))
    plt.yscale('log')
    plt.errorbar(x[1:], y[1:], yerr=yerror[1:], capsize=3, fmt='o')
    plt.errorbar(x[0], y[0], yerr=yerror[0], capsize=3, fmt='ro')
    plt.plot(x, y_line, color='r', linestyle='-', linewidth=0.8)
    plt.xlabel("Numer grupy", fontsize='large')
    plt.ylabel("Gęstość śladów [$n/cm^2$]", fontsize='large')
    plt.title("Część 2, Zdjęcie kalibracyjne 2, Gęstość śladów", fontsize='large')
    plt.grid(axis='x', linestyle='--', color='0.95')
    plt.grid(axis='y', linestyle='--', color='0.95')
    plt.xticks(x)
    plt.savefig(filename, dpi=600)
    plt.close() 
    return

def partB_Cal_Photo3_Rho(DirectoryPath, partBData):
    '''
    Creates a graph of dencity of traces of 'calibration photo3' vs. the number of the student group.
            Parameters:
                    DirectoryPath (str): Path to folder where graph will be stored
                    partaBData (pandas DataFrame):  DataFrame with results of <bemchmark part2> collected from Excel files.
    '''
    x = partBData["id"]
    y = partBData["cal_photo3_ro"]
    partBData['cal_photo3_u_ro'] = partBData['cal_photo3_u_ro'].fillna(0)
    yerror = partBData["cal_photo3_u_ro"]
    filename = DirectoryPath + "B_kal_zdjecie3_Rho.png"
    #create line
    y_line = [y[0]] * len(y)

    plt.figure(figsize=(16,4))
    plt.yscale('log')
    plt.errorbar(x[1:], y[1:], yerr=yerror[1:], capsize=3, fmt='o')
    plt.errorbar(x[0], y[0], yerr=yerror[0], capsize=3, fmt='ro')
    plt.plot(x, y_line, color='r', linestyle='-', linewidth=0.8)
    plt.xlabel("Numer grupy", fontsize='large')
    plt.ylabel("Gęstość śladów [$n/cm^2$]", fontsize='large')
    plt.title("Część 2, Zdjęcie kalibracyjne 3, Gęstość śladów", fontsize='large')
    plt.grid(axis='x', linestyle='--', color='0.95')
    plt.grid(axis='y', linestyle='--', color='0.95')
    plt.xticks(x)
    plt.savefig(filename, dpi=600)
    plt.close() 
    return

def partB_Cal_ParametrA(DirectoryPath, partBData):
    '''
    Creates a graph of parameter 'a' value of linear function (calibration) vs. the number of the student group.
            Parameters:
                    DirectoryPath (str): Path to folder where graph will be stored
                    partaBData (pandas DataFrame):  DataFrame with results of <bemchmark part2> collected from Excel files.
    '''
    x = partBData["id"]
    y = partBData["cal_param_a"]
    partBData['cal_param_u_a'] = partBData['cal_param_u_a'].fillna(0)
    yerror = partBData["cal_param_u_a"]
    filename = DirectoryPath + "B_kal_parametr_a.png"
    #create line
    y_line = [y[0]] * len(y)

    plt.figure(figsize=(16,6))
    plt.yscale('log')
    plt.errorbar(x[1:], y[1:], yerr=yerror[1:], capsize=3, fmt='o')
    plt.errorbar(x[0], y[0], yerr=yerror[0], capsize=3, fmt='ro')
    plt.plot(x, y_line, color='r', linestyle='-', linewidth=0.8)
    plt.xlabel("Numer grupy", fontsize='large')
    plt.ylabel("Wartość parametru $a$", fontsize='large')
    plt.title("Część 2, Krzywa kalibracyjna, parametr $a$", fontsize='large')
    plt.grid(axis='x', linestyle='--', color='0.95')
    plt.grid(axis='y', linestyle='--', color='0.95')
    plt.xticks(x)
    plt.savefig(filename, dpi=600)
    plt.close() 
    return

def partB_Cal_ParametrB(DirectoryPath, partBData):
    '''
    Creates a graph of parameter 'b' value of linear function (calibration) vs. the number of the student group.
            Parameters:
                    DirectoryPath (str): Path to folder where graph will be stored
                    partaBData (pandas DataFrame):  DataFrame with results of <bemchmark part2> collected from Excel files.
    '''
    x = partBData["id"]
    y = partBData["cal_param_b"]
    partBData['cal_param_u_b'] = partBData['cal_param_u_b'].fillna(0)
    yerror = partBData["cal_param_u_b"]
    filename = DirectoryPath + "B_kal_parametr_b.png"
    #create line
    y_line = [y[0]] * len(y)

    plt.figure(figsize=(16,6))
    plt.yscale('log')
    plt.errorbar(x[1:], y[1:], yerr=yerror[1:], capsize=3, fmt='o')
    plt.errorbar(x[0], y[0], yerr=yerror[0], capsize=3, fmt='ro')
    plt.plot(x, y_line, color='r', linestyle='-', linewidth=0.8)
    plt.xlabel("Numer grupy", fontsize='large')
    plt.ylabel("Wartość parametru $b$", fontsize='large')
    plt.title("Część 2, Krzywa kalibracyjna, parametr $b$", fontsize='large')
    plt.grid(axis='x', linestyle='--', color='0.95')
    plt.grid(axis='y', linestyle='--', color='0.95')
    plt.xticks(x)
    plt.savefig(filename, dpi=600)
    plt.close() 
    return


def partB_Photo3_NumberOfTracks(DirectoryPath, partBData):
    '''
    Creates a graph of number of tracks of photo3 vs. the number of the student group.
            Parameters:
                    DirectoryPath (str): Path to folder where graph will be stored
                    partaBData (pandas DataFrame):  DataFrame with results of <bemchmark part2> collected from Excel files.
    '''
    x = partBData["id"]
    y = partBData["photo3_n"]
    partBData['photo3_u_n'] = partBData['photo3_u_n'].fillna(0)
    yerror = partBData["photo3_u_n"]
    filename = DirectoryPath + "B_zdjecie3_n.png"
    #create line
    y_line = [y[0]] * len(y)

    plt.figure(figsize=(16,8))
    plt.errorbar(x[1:], y[1:], yerr=yerror[1:], capsize=3, fmt='o')
    plt.errorbar(x[0], y[0], yerr=yerror[0], capsize=3, fmt='ro')
    plt.plot(x, y_line, color='r', linestyle='-', linewidth=0.8)
    plt.xlabel("Numer grupy", fontsize='large')
    plt.ylabel("Liczba śladów", fontsize='large')
    plt.title("Część 2, Zdjęcie 3, Liczba śladów", fontsize='large')
    plt.grid(axis='x', linestyle='--', color='0.95')
    plt.grid(axis='y', linestyle='--', color='0.95')
    plt.xticks(x)
    plt.savefig(filename, dpi=600)
    plt.close() 
    return

def partB_Photo4_NumberOfTracks(DirectoryPath, partBData):
    '''
    Creates a graph of number of tracks of photo4 vs. the number of the student group.
            Parameters:
                    DirectoryPath (str): Path to folder where graph will be stored
                    partaBData (pandas DataFrame):  DataFrame with results of <bemchmark part2> collected from Excel files.
    '''
    x = partBData["id"]
    y = partBData["photo4_n"]
    partBData['photo4_u_n'] = partBData['photo4_u_n'].fillna(0)
    yerror = partBData["photo4_u_n"]
    filename = DirectoryPath + "B_zdjecie4_n.png"
    #create line
    y_line = [y[0]] * len(y)

    plt.figure(figsize=(16,8))
    plt.errorbar(x[1:], y[1:], yerr=yerror[1:], capsize=3, fmt='o')
    plt.errorbar(x[0], y[0], yerr=yerror[0], capsize=3, fmt='ro')
    plt.plot(x, y_line, color='r', linestyle='-', linewidth=0.8)
    plt.xlabel("Numer grupy", fontsize='large')
    plt.ylabel("Liczba śladów", fontsize='large')
    plt.title("Część 2, Zdjęcie 4, Liczba śladów", fontsize='large')
    plt.grid(axis='x', linestyle='--', color='0.95')
    plt.grid(axis='y', linestyle='--', color='0.95')
    plt.xticks(x)
    plt.savefig(filename, dpi=600)
    plt.close() 
    return

def partB_Photo3_Surface(DirectoryPath, partBData):
    '''
    Creates a graph of surface area of photo3 vs. the number of the student group.
            Parameters:
                    DirectoryPath (str): Path to folder where graph will be stored
                    partaBData (pandas DataFrame):  DataFrame with results of <bemchmark part2> collected from Excel files.
    '''
    x = partBData["id"]
    y = partBData["photo3_S"]
    partBData['photo3_S'] = partBData['photo3_S'].fillna(0)
    partBData['photo3_u_S'] = partBData['photo3_u_S'].fillna(0)
    yerror = partBData["photo3_u_S"]
    filename = DirectoryPath + "B_zdjecie3_S.png"
    #create line
    y_line = [y[0]] * len(y)

    plt.figure(figsize=(16,8))
    plt.yscale('log')
    plt.errorbar(x[1:], y[1:], yerr=yerror[1:], capsize=3, fmt='o')
    plt.errorbar(x[0], y[0], yerr=yerror[0], capsize=3, fmt='ro')
    plt.plot(x, y_line, color='r', linestyle='-', linewidth=0.8)
    plt.xlabel("Numer grupy", fontsize='large')
    plt.ylabel("Pole powierzchni [$cm^2$]", fontsize='large')
    plt.title("Część 2, Zdjęcie 3, Pole powierzchni zdjęcia", fontsize='large')
    plt.grid(axis='x', linestyle='--', color='0.95')
    plt.grid(axis='y', linestyle='--', color='0.95')
    plt.xticks(x)
    plt.savefig(filename, dpi=600)
    plt.close() 
    return

def partB_Photo4_Surface(DirectoryPath, partBData):
    '''
    Creates a graph of surface area of photo4 vs. the number of the student group.
            Parameters:
                    DirectoryPath (str): Path to folder where graph will be stored
                    partaBData (pandas DataFrame):  DataFrame with results of <bemchmark part2> collected from Excel files.
    '''
    x = partBData["id"]
    y = partBData["photo4_S"]
    partBData['photo4_S'] = partBData['photo4_S'].fillna(0)
    partBData['photo4_u_S'] = partBData['photo4_u_S'].fillna(0)
    yerror = partBData["photo4_u_S"]
    filename = DirectoryPath + "B_zdjecie4_S.png"
    #create line
    y_line = [y[0]] * len(y)

    plt.figure(figsize=(16,8))
    plt.yscale('log')
    plt.errorbar(x[1:], y[1:], yerr=yerror[1:], capsize=3, fmt='o')
    plt.errorbar(x[0], y[0], yerr=yerror[0], capsize=3, fmt='ro')
    plt.plot(x, y_line, color='r', linestyle='-', linewidth=0.8)
    plt.xlabel("Numer grupy", fontsize='large')
    plt.ylabel("Pole powierzchni [$cm^2$]", fontsize='large')
    plt.title("Część 2, Zdjęcie 4, Pole powierzchni zdjęcia", fontsize='large')
    plt.grid(axis='x', linestyle='--', color='0.95')
    plt.grid(axis='y', linestyle='--', color='0.95')
    plt.xticks(x)
    plt.savefig(filename, dpi=600)
    plt.close() 
    return

def partB_Photo3_Rho(DirectoryPath, partBData):
    '''
    Creates a graph of dencity of tracks of photo3 vs. the number of the student group.
            Parameters:
                    DirectoryPath (str): Path to folder where graph will be stored
                    partaBData (pandas DataFrame):  DataFrame with results of <bemchmark part2> collected from Excel files.
    '''
    x = partBData["id"]
    y = partBData["photo3_ro"]
    partBData['photo3_ro'] = partBData['photo3_ro'].fillna(0)
    partBData['photo3_u_ro'] = partBData['photo3_u_ro'].fillna(0)
    yerror = partBData["photo3_u_ro"]
    filename = DirectoryPath + "B_zdjecie3_rho.png"
    #create line
    y_line = [y[0]] * len(y)

    plt.figure(figsize=(16,8))
    plt.yscale('log')
    plt.errorbar(x[1:], y[1:], yerr=yerror[1:], capsize=3, fmt='o')
    plt.errorbar(x[0], y[0], yerr=yerror[0], capsize=3, fmt='ro')
    plt.plot(x, y_line, color='r', linestyle='-', linewidth=0.8)
    plt.xlabel("Numer grupy", fontsize='large')
    plt.ylabel("Gęstość śladów [$n/cm^2$]", fontsize='large')
    plt.title("Część 2, Zdjęcie 3, Gęstość śladów", fontsize='large')
    plt.grid(axis='x', linestyle='--', color='0.95')
    plt.grid(axis='y', linestyle='--', color='0.95')
    plt.xticks(x)
    plt.savefig(filename, dpi=600)
    plt.close() 
    return

def partB_Photo4_Rho(DirectoryPath, partBData):
    '''
    Creates a graph of dencity of tracks of photo4 vs. the number of the student group.
            Parameters:
                    DirectoryPath (str): Path to folder where graph will be stored
                    partaBData (pandas DataFrame):  DataFrame with results of <bemchmark part2> collected from Excel files.
    '''
    x = partBData["id"]
    y = partBData["photo4_ro"]
    partBData['photo4_ro'] = partBData['photo4_ro'].fillna(0)
    partBData['photo4_u_ro'] = partBData['photo4_u_ro'].fillna(0)
    yerror = partBData["photo4_u_ro"]
    filename = DirectoryPath + "B_zdjecie4_rho.png"
    #create line
    y_line = [y[0]] * len(y)

    plt.figure(figsize=(16,8))
    plt.yscale('log')
    plt.errorbar(x[1:], y[1:], yerr=yerror[1:], capsize=3, fmt='o')
    plt.errorbar(x[0], y[0], yerr=yerror[0], capsize=3, fmt='ro')
    plt.plot(x, y_line, color='r', linestyle='-', linewidth=0.8)
    plt.xlabel("Numer grupy", fontsize='large')
    plt.ylabel("Gęstość śladów [$n/cm^2$]", fontsize='large')
    plt.title("Część 2, Zdjęcie 4, Gęstość śladów", fontsize='large')
    plt.grid(axis='x', linestyle='--', color='0.95')
    plt.grid(axis='y', linestyle='--', color='0.95')
    plt.xticks(x)
    plt.savefig(filename, dpi=600)
    plt.close() 
    return

def partB_Photo3_RadonConcentration(DirectoryPath, partBData):
    '''
    Creates a graph of radon concentration from photo3 vs. the number of the student group.
            Parameters:
                    DirectoryPath (str): Path to folder where graph will be stored
                    partaBData (pandas DataFrame):  DataFrame with results of <bemchmark part2> collected from Excel files.
    '''
    x = partBData["id"]
    y = partBData["photo3_Bq"]
    partBData['photo3_Bq'] = partBData['photo3_Bq'].fillna(0)
    partBData['photo3_u_Bq'] = partBData['photo3_u_Bq'].fillna(0)
    yerror = partBData["photo3_u_Bq"]
    filename = DirectoryPath + "B_zdjecie3_Bq.png"
    #create line
    y_line = [y[0]] * len(y)

    plt.figure(figsize=(16,8))
    plt.yscale('log')
    plt.errorbar(x[1:], y[1:], yerr=yerror[1:], capsize=3, fmt='o')
    plt.errorbar(x[0], y[0], yerr=yerror[0], capsize=3, fmt='ro')
    plt.plot(x, y_line, color='r', linestyle='-', linewidth=0.8)
    plt.xlabel("Numer grupy", fontsize='large')
    plt.ylabel("Stężenie radonu [$Bq/m^3$]", fontsize='large')
    plt.title("Część 2, Zdjęcie 3, Stężenie radonu", fontsize='large')
    plt.grid(axis='x', linestyle='--', color='0.95')
    plt.grid(axis='y', linestyle='--', color='0.95')
    plt.xticks(x)
    plt.savefig(filename, dpi=600)
    plt.close() 
    return

def partB_Photo4_RadonConcentration(DirectoryPath, partBData):
    '''
    Creates a graph of radon concentration from photo4 vs. the number of the student group.
            Parameters:
                    DirectoryPath (str): Path to folder where graph will be stored
                    partaBData (pandas DataFrame):  DataFrame with results of <bemchmark part2> collected from Excel files.
    '''
    x = partBData["id"]
    y = partBData["photo4_Bq"]
    partBData['photo4_Bq'] = partBData['photo4_Bq'].fillna(0)
    partBData['photo4_u_Bq'] = partBData['photo4_u_Bq'].fillna(0)
    yerror = partBData["photo4_u_Bq"]
    filename = DirectoryPath + "B_zdjecie4_Bq.png"
    #create line
    y_line = [y[0]] * len(y)

    plt.figure(figsize=(16,8))
    plt.yscale('log')
    plt.errorbar(x[1:], y[1:], yerr=yerror[1:], capsize=3, fmt='o')
    plt.errorbar(x[0], y[0], yerr=yerror[0], capsize=3, fmt='ro')
    plt.plot(x, y_line, color='r', linestyle='-', linewidth=0.8)
    plt.xlabel("Numer grupy", fontsize='large')
    plt.ylabel("Stężenie radonu [$Bq/m^3$]", fontsize='large')
    plt.title("Część 2, Zdjęcie 4, Stężenie radonu", fontsize='large')
    plt.grid(axis='x', linestyle='--', color='0.95')
    plt.grid(axis='y', linestyle='--', color='0.95')
    plt.xticks(x)
    plt.savefig(filename, dpi=600)
    plt.close() 
    return

def drawCalibration(DirectoryPath, partBData):
    '''
    Creates a graph of linear functions - calibrations - created by all student's groups.
            Parameters:
                    DirectoryPath (str): Path to folder where graph will be stored
                    partaBData (pandas DataFrame):  DataFrame with results of <bemchmark part2> collected from Excel files.
    '''
    filename = DirectoryPath + "all_calibrations.png"
    partBData = partBData[partBData['cal_param_a'].notna()]
    id = partBData["id"]
    a = partBData["cal_param_a"]
    b = partBData["cal_param_b"]
    
    x = np.linspace(0,100000,1000)
    n_index = partBData.index

    plt.figure(figsize=(12,8))
    plt.xlabel("Gęstość śladów [$n/cm^2$]", fontsize='large')
    plt.ylabel("Stężenie radonu [$Bq/m^3$]", fontsize='large')
    plt.title("Krzywe kalibracyjne", fontsize='large') 
    plt.ylim([0, 20000])   
    for i in n_index:
        y = float(a[i]) * x + float(b[i])
        plt.plot(x, y, '-', label=str(i))
    plt.savefig(filename, dpi=600)
    plt.close() 
    return

DirectoryPath = "E:\\A-- FUNDACJA\\PROJEKTY\\Szkolna Radonowa Mapa Polski\\Benchmark\\"
ListofExcelFiles = getListofExcelFiles(DirectoryPath, 'xls')
#print(ListofExcelFiles)

#pd.set_option('display.max_columns', None)
partBData = collectDataPartB(ListofExcelFiles)
#print(partaBData)

#data = partaBData.iloc[:, :6]
#print(data)

partB_Cal_Photo1_NumberOfTracks(DirectoryPath, partBData)
partB_Cal_Photo2_NumberOfTracks(DirectoryPath, partBData)
partB_Cal_Photo3_NumberOfTracks(DirectoryPath, partBData)
partB_Cal_Photo1_Surface(DirectoryPath, partBData)
partB_Cal_Photo2_Surface(DirectoryPath, partBData)
partB_Cal_Photo3_Surface(DirectoryPath, partBData)
partB_Cal_Photo1_Rho(DirectoryPath, partBData)
partB_Cal_Photo2_Rho(DirectoryPath, partBData)
partB_Cal_Photo3_Rho(DirectoryPath, partBData)
partB_Cal_ParametrA(DirectoryPath, partBData)
partB_Cal_ParametrB(DirectoryPath, partBData)

partB_Photo3_NumberOfTracks(DirectoryPath, partBData)
partB_Photo4_NumberOfTracks(DirectoryPath, partBData)

partB_Photo3_Surface(DirectoryPath, partBData)
partB_Photo4_Surface(DirectoryPath, partBData)

partB_Photo3_Rho(DirectoryPath, partBData)
partB_Photo4_Rho(DirectoryPath, partBData)

partB_Photo3_RadonConcentration(DirectoryPath, partBData)
partB_Photo4_RadonConcentration(DirectoryPath, partBData)

drawCalibration(DirectoryPath, partBData)
