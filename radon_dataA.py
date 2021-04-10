import pandas as pd
import os
from radon_statistics import getListofExcelFiles
import matplotlib.pyplot as plt

def collectDataPartA(ListofExcelFiles):
    '''
    Collects all results of <benchmark part 1>. Data from every Excel file is stored in a dictionary.
            Parameters:
                    ListofExcelFiles (list of str): List of .xls files with full path
            Returns:
                    partaAData (pandas DataFrame):  DataFrame created from List of dictionaries (listPartA)
                                                    Every dictionary in listPartA collects data from one Excel file
    '''
    listPartA = []
    for file in ListofExcelFiles:
        dict_with_data_A = {}
        
        file_name = os.path.basename(file)
        id = int(file_name[0:2])
        dict_with_data_A['id'] = id
        
        #------------------------------------------------------
        #general info
        general_info_sheet = pd.read_excel(file, sheet_name="Informacje ogólne")
        
        school = general_info_sheet.at[1,"Unnamed: 2"]
        city = general_info_sheet.at[2,"Unnamed: 2"]
        school_full = school + " - " + city
        dict_with_data_A['school'] = school_full
        
        teacher = general_info_sheet.at[4,"Unnamed: 2"]
        dict_with_data_A['teacher'] = teacher
        
        students = general_info_sheet.loc[6:,['Unnamed: 2']].values.tolist()
        students_flat_list = [item for sublist in students for item in sublist]
        dict_with_data_A['students'] = students_flat_list
        
        #------------------------------------------------------
        #data - part A
        partA_sheet = pd.read_excel(file, sheet_name="Część 1")
        all_done = partA_sheet.at[0,"Unnamed: 4"]
        dict_with_data_A['all done'] = all_done

        #Photo 1
        photo1_n = int(partA_sheet.at[7,"Unnamed: 4"])
        dict_with_data_A['photo1_n'] = photo1_n
        photo1_u_n = float(partA_sheet.at[8,"Unnamed: 4"])
        dict_with_data_A['photo1_u_n'] = photo1_u_n
        photo1_S = float(partA_sheet.at[10,"Unnamed: 4"])
        dict_with_data_A['photo1_S'] = photo1_S
        photo1_u_S = float(partA_sheet.at[11,"Unnamed: 4"])
        dict_with_data_A['photo1_u_S'] = photo1_u_S
        photo1_ro = float(partA_sheet.at[13,"Unnamed: 4"])
        dict_with_data_A['photo1_ro'] = photo1_ro
        photo1_u_ro = float(partA_sheet.at[14,"Unnamed: 4"])
        dict_with_data_A['photo1_u_ro'] = photo1_u_ro
        photo1_Bq = float(partA_sheet.at[16,"Unnamed: 4"])
        dict_with_data_A['photo1_Bq'] = photo1_Bq
        photo1_u_Bq = float(partA_sheet.at[17,"Unnamed: 4"])
        dict_with_data_A['photo1_u_Bq'] = photo1_u_Bq
        #Photo 2
        photo2_n = int(partA_sheet.at[7,"Unnamed: 6"])
        dict_with_data_A['photo2_n'] = photo2_n
        photo2_u_n = float(partA_sheet.at[8,"Unnamed: 6"])
        dict_with_data_A['photo2_u_n'] = photo2_u_n
        photo2_S = float(partA_sheet.at[10,"Unnamed: 6"])
        dict_with_data_A['photo2_S'] = photo2_S
        photo2_u_S = float(partA_sheet.at[11,"Unnamed: 6"])
        dict_with_data_A['photo2_u_S'] = photo2_u_S
        photo2_ro = float(partA_sheet.at[13,"Unnamed: 6"])
        dict_with_data_A['photo2_ro'] = photo2_ro
        photo2_u_ro = float(partA_sheet.at[14,"Unnamed: 6"])
        dict_with_data_A['photo2_u_ro'] = photo2_u_ro
        photo2_Bq = float(partA_sheet.at[16,"Unnamed: 6"])
        dict_with_data_A['photo2_Bq'] = photo2_Bq
        photo2_u_Bq = float(partA_sheet.at[17,"Unnamed: 6"])
        dict_with_data_A['photo2_u_Bq'] = photo2_u_Bq

        #print(dict_with_data_A)
        #print('\n')
        listPartA.append(dict_with_data_A)
    partaAData = pd.DataFrame(listPartA)
    return partaAData


def partA_Photo1_NumberOfTracks(DirectoryPath, partaAData):
    '''
    Creates a graph of number of tracs on photo1 vs. the number of the student group.
            Parameters:
                    DirectoryPath (str): Path to folder where graph will be stored
                    partaAData (pandas DataFrame):  DataFrame with results of <bemchmark part1> collected from Excel files.
    '''
    x = partaAData["id"]
    y = partaAData["photo1_n"]
    partaAData['photo1_u_n'] = partaAData['photo1_u_n'].fillna(0)
    yerror = partaAData["photo1_u_n"]
    filename = DirectoryPath + "A_zdjecie1_n.png"
    #create line
    y_line = [y[0]] * len(y)

    plt.figure(figsize=(16,8))
    plt.errorbar(x[1:], y[1:], yerr=yerror[1:], capsize=3, fmt='o')
    plt.errorbar(x[0], y[0], yerr=yerror[0], capsize=3, fmt='ro')
    plt.plot(x, y_line, color='r', linestyle='-', linewidth=0.8)
    plt.xlabel("Numer grupy", fontsize='large')
    plt.ylabel("Liczba śladów", fontsize='large')
    plt.title("Część 1, Zdjęcie 1, Liczba śladów", fontsize='large')
    plt.grid(axis='x', linestyle='--', color='0.95')
    plt.grid(axis='y', linestyle='--', color='0.95')
    plt.xticks(x)
    plt.savefig(filename, dpi=600)
    plt.close() 
    return

def partA_Photo2_NumberOfTracks(DirectoryPath, partaAData):
    '''
    Creates a graph of number of tracs on photo2 vs. the number of the student group.
            Parameters:
                    DirectoryPath (str): Path to folder where graph will be stored
                    partaAData (pandas DataFrame):  DataFrame with results of <bemchmark part1> collected from Excel files.
    '''
    x = partaAData["id"]
    y = partaAData["photo2_n"]
    partaAData['photo2_u_n'] = partaAData['photo2_u_n'].fillna(0)
    yerror = partaAData["photo2_u_n"]
    filename = DirectoryPath + "A_zdjecie2_n.png"
    #create line
    y_line = [y[0]] * len(y)

    plt.figure(figsize=(16,8))
    plt.errorbar(x[1:], y[1:], yerr=yerror[1:], capsize=3, fmt='o')
    plt.errorbar(x[0], y[0], yerr=yerror[0], capsize=3, fmt='ro')
    plt.plot(x, y_line, color='r', linestyle='-', linewidth=0.8)
    plt.xlabel("Numer grupy", fontsize='large')
    plt.ylabel("Liczba śladów", fontsize='large')
    plt.title("Część 1, Zdjęcie 2, Liczba śladów", fontsize='large')
    plt.grid(axis='x', linestyle='--', color='0.95')
    plt.grid(axis='y', linestyle='--', color='0.95')
    plt.xticks(x)
    plt.savefig(filename, dpi=600)
    plt.close() 
    return

def partA_Photo1_Surface(DirectoryPath, partaAData):
    '''
    Creates a graph of surface area of photo1 vs. the number of the student group.
            Parameters:
                    DirectoryPath (str): Path to folder where graph will be stored
                    partaAData (pandas DataFrame):  DataFrame with results of <bemchmark part1> collected from Excel files.
    '''
    x = partaAData["id"]
    y = partaAData["photo1_S"]
    partaAData['photo1_S'] = partaAData['photo1_S'].fillna(0)
    partaAData['photo1_u_S'] = partaAData['photo1_u_S'].fillna(0)
    yerror = partaAData["photo1_u_S"]
    filename = DirectoryPath + "A_zdjecie1_S.png"
    #create line
    y_line = [y[0]] * len(y)

    plt.figure(figsize=(16,8))
    plt.yscale('log')
    plt.errorbar(x[1:], y[1:], yerr=yerror[1:], capsize=3, fmt='o')
    plt.errorbar(x[0], y[0], yerr=yerror[0], capsize=3, fmt='ro')
    plt.plot(x, y_line, color='r', linestyle='-', linewidth=0.8)
    plt.xlabel("Numer grupy", fontsize='large')
    plt.ylabel("Pole powierzchni zdjęcia [$cm^2$]", fontsize='large')
    plt.title("Część 1, Zdjęcie 1, Pole powierzchni zdjęcia", fontsize='large')
    plt.grid(axis='x', linestyle='--', color='0.95')
    plt.grid(axis='y', linestyle='--', color='0.95')
    plt.xticks(x)
    plt.savefig(filename, dpi=600)
    plt.close() 
    return

def partA_Photo2_Surface(DirectoryPath, partaAData):
    '''
    Creates a graph of surface area of photo2 vs. the number of the student group.
            Parameters:
                    DirectoryPath (str): Path to folder where graph will be stored
                    partaAData (pandas DataFrame):  DataFrame with results of <bemchmark part1> collected from Excel files.
    '''
    x = partaAData["id"]
    y = partaAData["photo2_S"]
    partaAData['photo2_S'] = partaAData['photo2_S'].fillna(0)
    partaAData['photo2_u_S'] = partaAData['photo2_u_S'].fillna(0)
    yerror = partaAData["photo2_u_S"]
    filename = DirectoryPath + "A_zdjecie2_S.png"
    #create line
    y_line = [y[0]] * len(y)

    plt.figure(figsize=(16,8))
    plt.yscale('log')
    plt.errorbar(x[1:], y[1:], yerr=yerror[1:], capsize=3, fmt='o')
    plt.errorbar(x[0], y[0], yerr=yerror[0], capsize=3, fmt='ro')
    plt.plot(x, y_line, color='r', linestyle='-', linewidth=0.8)
    plt.xlabel("Numer grupy", fontsize='large')
    plt.ylabel("Pole powierzchni zdjęcia [$cm^2$]", fontsize='large')
    plt.title("Część 1, Zdjęcie 2, Pole powierzchni zdjęcia", fontsize='large')
    plt.grid(axis='x', linestyle='--', color='0.95')
    plt.grid(axis='y', linestyle='--', color='0.95')
    plt.xticks(x)
    plt.savefig(filename, dpi=600)
    plt.close() 
    return

def partA_Photo1_Rho(DirectoryPath, partaAData):
    '''
    Creates a graph of density of the traces of photo1 vs. the number of the student group.
            Parameters:
                    DirectoryPath (str): Path to folder where graph will be stored
                    partaAData (pandas DataFrame):  DataFrame with results of <bemchmark part1> collected from Excel files.
    '''
    x = partaAData["id"]
    y = partaAData["photo1_ro"]
    partaAData['photo1_ro'] = partaAData['photo1_ro'].fillna(0)
    partaAData['photo1_u_ro'] = partaAData['photo1_u_ro'].fillna(0)
    yerror = partaAData["photo1_u_ro"]
    filename = DirectoryPath + "A_zdjecie1_rho.png"
    #create line
    y_line = [y[0]] * len(y)

    plt.figure(figsize=(16,8))
    plt.yscale('log')
    plt.errorbar(x[1:], y[1:], yerr=yerror[1:], capsize=3, fmt='o')
    plt.errorbar(x[0], y[0], yerr=yerror[0], capsize=3, fmt='ro')
    plt.plot(x, y_line, color='r', linestyle='-', linewidth=0.8)
    plt.xlabel("Numer grupy", fontsize='large')
    plt.ylabel("Gęstość śladów [$n/cm^2$]", fontsize='large')
    plt.title("Część 1, Zdjęcie 1, Gęstość śladów", fontsize='large')
    plt.grid(axis='x', linestyle='--', color='0.95')
    plt.grid(axis='y', linestyle='--', color='0.95')
    plt.xticks(x)
    plt.savefig(filename, dpi=600)
    plt.close() 
    return

def partA_Photo2_Rho(DirectoryPath, partaAData):
    '''
    Creates a graph of density of the traces of photo2 vs. the number of the student group.
            Parameters:
                    DirectoryPath (str): Path to folder where graph will be stored
                    partaAData (pandas DataFrame):  DataFrame with results of <bemchmark part1> collected from Excel files.
    '''
    x = partaAData["id"]
    y = partaAData["photo2_ro"]
    partaAData['photo2_ro'] = partaAData['photo2_ro'].fillna(0)
    partaAData['photo2_u_ro'] = partaAData['photo2_u_ro'].fillna(0)
    yerror = partaAData["photo2_u_ro"]
    filename = DirectoryPath + "A_zdjecie2_rho.png"
    #create line
    y_line = [y[0]] * len(y)

    plt.figure(figsize=(16,8))
    plt.yscale('log')
    plt.errorbar(x[1:], y[1:], yerr=yerror[1:], capsize=3, fmt='o')
    plt.errorbar(x[0], y[0], yerr=yerror[0], capsize=3, fmt='ro')
    plt.plot(x, y_line, color='r', linestyle='-', linewidth=0.8)
    plt.xlabel("Numer grupy", fontsize='large')
    plt.ylabel("Gęstość śladów [$n/cm^2$]", fontsize='large')
    plt.title("Część 1, Zdjęcie 2, Gęstość śladów", fontsize='large')
    plt.grid(axis='x', linestyle='--', color='0.95')
    plt.grid(axis='y', linestyle='--', color='0.95')
    plt.xticks(x)
    plt.savefig(filename, dpi=600)
    plt.close() 
    return

def partA_Photo1_RadonConcentration(DirectoryPath, partaAData):
    '''
    Creates a graph of radon concentration from photo1 vs. the number of the student group.
            Parameters:
                    DirectoryPath (str): Path to folder where graph will be stored
                    partaAData (pandas DataFrame):  DataFrame with results of <bemchmark part1> collected from Excel files.
    '''
    x = partaAData["id"]
    y = partaAData["photo1_Bq"]
    partaAData['photo1_Bq'] = partaAData['photo1_Bq'].fillna(0)
    partaAData['photo1_u_Bq'] = partaAData['photo1_u_Bq'].fillna(0)
    yerror = partaAData["photo1_u_Bq"]
    filename = DirectoryPath + "A_zdjecie1_Bq.png"
    #create line
    y_line = [y[0]] * len(y)

    plt.figure(figsize=(16,8))
    plt.yscale('log')
    plt.errorbar(x[1:], y[1:], yerr=yerror[1:], capsize=3, fmt='o')
    plt.errorbar(x[0], y[0], yerr=yerror[0], capsize=3, fmt='ro')
    plt.plot(x, y_line, color='r', linestyle='-', linewidth=0.8)
    plt.xlabel("Numer grupy", fontsize='large')
    plt.ylabel("Stężenie radonu [$Bq/m^3$]", fontsize='large')
    plt.title("Część 1, Zdjęcie 1, Stężenie radonu", fontsize='large')
    plt.grid(axis='x', linestyle='--', color='0.95')
    plt.grid(axis='y', linestyle='--', color='0.95')
    plt.xticks(x)
    plt.savefig(filename, dpi=600)
    plt.close() 
    return

def partA_Photo2_RadonConcentration(DirectoryPath, partaAData):
    '''
    Creates a graph of radon concentration from photo2 vs. the number of the student group.
            Parameters:
                    DirectoryPath (str): Path to folder where graph will be stored
                    partaAData (pandas DataFrame):  DataFrame with results of <bemchmark part1> collected from Excel files.
    '''
    x = partaAData["id"]
    y = partaAData["photo2_Bq"]
    partaAData['photo2_Bq'] = partaAData['photo2_Bq'].fillna(0)
    partaAData['photo2_u_Bq'] = partaAData['photo2_u_Bq'].fillna(0)
    yerror = partaAData["photo2_u_Bq"]
    filename = DirectoryPath + "A_zdjecie2_Bq.png"
    #create line
    y_line = [y[0]] * len(y)

    plt.figure(figsize=(16,8))
    plt.yscale('log')
    plt.errorbar(x[1:], y[1:], yerr=yerror[1:], capsize=3, fmt='o')
    plt.errorbar(x[0], y[0], yerr=yerror[0], capsize=3, fmt='ro')
    plt.plot(x, y_line, color='r', linestyle='-', linewidth=0.8)
    plt.xlabel("Numer grupy", fontsize='large')
    plt.ylabel("Stężenie radonu [$Bq/m^3$]", fontsize='large')
    plt.title("Część 1, Zdjęcie 2, Stężenie radonu", fontsize='large')
    plt.grid(axis='x', linestyle='--', color='0.95')
    plt.grid(axis='y', linestyle='--', color='0.95')
    plt.xticks(x)
    plt.savefig(filename, dpi=600)
    plt.close() 
    return

def plotHistogramN(DirectoryPath, partaAData):
    '''
    Creates a graph - histogram of number of tracks on photo2.
            Parameters:
                    DirectoryPath (str): Path to folder where graph will be stored
                    partaAData (pandas DataFrame):  DataFrame with results of <bemchmark part1> collected from Excel files
    '''
    filename = DirectoryPath + "histogram_photo2_n.png"
    
    plt.figure(figsize=(16,8))
    partaAData['photo2_n'].plot.hist(grid=True, bins=25, rwidth=0.9, color='#607c8e')
    plt.savefig(filename, dpi=600)
    return

def plotHistogramS(DirectoryPath, partaAData):
    '''
    Creates a graph - histogram of surface area of photo2.
            Parameters:
                    DirectoryPath (str): Path to folder where graph will be stored
                    partaAData (pandas DataFrame):  DataFrame with results of <bemchmark part1> collected from Excel files.
    '''
    filename = DirectoryPath + "histogram_photo2_S.png"
    
    plt.figure(figsize=(16,8))
    #plt.xlim([0, 1000])   
    partaAData['photo2_S'].plot.hist(grid=True, bins=8, rwidth=0.9, color='#336600', range=[0, 100])
    plt.savefig(filename, dpi=600)
    return


DirectoryPath = "E:\\A-- FUNDACJA\\PROJEKTY\\Szkolna Radonowa Mapa Polski\\Benchmark\\"
ListofExcelFiles = getListofExcelFiles(DirectoryPath, 'xls')
#print(ListofExcelFiles)

pd.set_option('display.max_columns', None)
partaAData = collectDataPartA(ListofExcelFiles)
#print(partaAData)

partA_Photo1_NumberOfTracks(DirectoryPath, partaAData)
partA_Photo2_NumberOfTracks(DirectoryPath, partaAData)
partA_Photo1_Surface(DirectoryPath, partaAData)
partA_Photo2_Surface(DirectoryPath, partaAData)
partA_Photo1_Rho(DirectoryPath, partaAData)
partA_Photo2_Rho(DirectoryPath, partaAData)
partA_Photo1_RadonConcentration(DirectoryPath, partaAData)
partA_Photo2_RadonConcentration(DirectoryPath, partaAData)

plotHistogramN(DirectoryPath, partaAData)
plotHistogramS(DirectoryPath, partaAData)
