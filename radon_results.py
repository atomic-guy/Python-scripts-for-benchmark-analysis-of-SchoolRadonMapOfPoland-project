import pandas as pd
import os
from radon_statistics import getListofExcelFiles
import matplotlib.pyplot as plt
from radon_dataA import collectDataPartA
from radon_dataB import collectDataPartB


def calculateAverageValues(dict):
    '''
    Calculates weighted average of number of traces and surface area of photo2 (part1).  
            Parameters:
                    dict (dict): dictionary with data collected from all Excel files 
            Returns:
                    n_avg (float): value of weighted average of number of tracks on photo2
                    S_avg (float): value of weighted average of surface area of photo2
    '''
    #Number of traces
    n = dict['photo2_n']
    u_n = dict['photo2_u_n']
    n_avg = (n * u_n).sum() / u_n.sum()
    #Surface
    S = dict['photo2_S']
    u_S = dict['photo2_u_S']
    S_avg = (S * u_S).sum() / u_S.sum()
    return n_avg, S_avg

def selectBestResultsPartA(DirectoryPath, partAData):
    '''
    Selects groups of students which obtained the best results of benchmark, part 1.
    The selection is based on the difference between the results obtained by the students 
    and the reference value obtained by the Atomic Forum Foundation. 
    The differences of 5%, 20%, 50% and 100% of the reference value were determined.
    Writes results of selection to "partA_best.txt" file.
    Selection    
            Parameters:
                    DirectoryPath (str): Path to folder where partA_best.txt file will be stored
                    partaAData (pandas DataFrame):  DataFrame with results of <bemchmark part1> collected from Excel files
    '''
    photo1_BestValue = partAData.loc[0, 'photo1_Bq']
    photo2_BestValue = partAData.loc[0, 'photo2_Bq']
    print(photo1_BestValue)
    print(photo2_BestValue)
    
    ###################################
    # PHOTO 1
    file1 = open(DirectoryPath + "partA_best.txt","w+")
    file1.write("PHOTO 1 \n\n")
    
    # 100%
    partAData_100pc = partAData[partAData.photo1_Bq <= (2 * photo1_BestValue)]
    #print(partAData_100pc)
    p1index100 = partAData_100pc.index
    number_of_results_photo1_100pc = len(p1index100)
    print("Photo 1 - 100%\n")
    file1.write("Photo 1 - 100%\n")
    print("Total number: " + str(number_of_results_photo1_100pc))
    file1.write("Total number: " + str(number_of_results_photo1_100pc) + "\n")
    for row in partAData_100pc.index:
        print(row, end=" ")
        file1.write(str(row) + " ")
    file1.write("\n\n")
    
    # 50%
    partAData_50pc = partAData[partAData.photo1_Bq <= (photo1_BestValue + (0.5 * photo1_BestValue))]
    partAData_50pc = partAData_50pc[partAData_50pc.photo1_Bq >= (photo1_BestValue - (0.5 * photo1_BestValue))] 
    #print(partAData_100pc)
    p1index50 = partAData_50pc.index
    number_of_results_photo1_50pc = len(p1index50)
    print("\nPhoto 1 - 50%\n")
    file1.write("\nPhoto 1 - 50%\n")
    print("Total number: " + str(number_of_results_photo1_50pc))
    file1.write("Total number: " + str(number_of_results_photo1_50pc) + "\n")
    for row in partAData_50pc.index:
        print(row, end=" ")
        file1.write(str(row) + " ")
    file1.write("\n\n")
    
    # 20%
    partAData_20pc = partAData[partAData.photo1_Bq <= (photo1_BestValue + (0.2 * photo1_BestValue))]
    partAData_20pc = partAData_20pc[partAData_20pc.photo1_Bq >= (photo1_BestValue - (0.2 * photo1_BestValue))]
    #print(partAData_20pc)
    p1index20 = partAData_20pc.index
    number_of_results_photo1_20pc = len(p1index20)
    print("\nPhoto 1 - 20%")
    file1.write("\nPhoto 1 - 20%\n")
    print("Total number: " + str(number_of_results_photo1_20pc))
    file1.write("Total number: " + str(number_of_results_photo1_20pc) + "\n")
    for row in partAData_20pc.index:
        print(row, end=" ")
        file1.write(str(row) + " ")
    file1.write("\n\n")
       
    # 5%
    partAData_5pc = partAData[partAData.photo1_Bq <= (photo1_BestValue + (0.05 * photo1_BestValue))]
    partAData_5pc = partAData_5pc[partAData_5pc.photo1_Bq >= (photo1_BestValue - (0.05 * photo1_BestValue))]
    #print(partAData_20pc)
    p1index5 = partAData_5pc.index
    number_of_results_photo1_5pc = len(p1index5)
    print("\nPhoto 1 - 5%")
    file1.write("\nPhoto 1 - 5%\n")
    print("Total number: " + str(number_of_results_photo1_5pc))
    file1.write("Total number: " + str(number_of_results_photo1_5pc) + "\n")
    for row in partAData_5pc.index:
        print(row, end=" ")
        file1.write(str(row) + " ")
    file1.write("\n\n")
    
    # 1%
    partAData_1pc = partAData[partAData.photo1_Bq <= (photo1_BestValue + (0.01 * photo1_BestValue))]
    partAData_1pc = partAData_1pc[partAData_1pc.photo1_Bq >= (photo1_BestValue - (0.01 * photo1_BestValue))]
    #print(partAData_20pc)
    p1index1 = partAData_1pc.index
    number_of_results_photo1_1pc = len(p1index1)
    print("\nPhoto 1 - 1%")
    file1.write("\nPhoto 1 - 1%\n")
    print("Total number: " + str(number_of_results_photo1_1pc))
    file1.write("Total number: " + str(number_of_results_photo1_1pc) + "\n")
    for row in partAData_1pc.index:
        print(row, end=" ")
        file1.write(str(row) + " ")
    file1.write("\n\n")    


    ###################################
    # PHOTO 2
    file1.write("\nPHOTO 2 \n\n")

    # 100%
    partAData_100pc = partAData[partAData.photo2_Bq.abs() <= (2 * photo2_BestValue)]
    #print(partAData_100pc)
    p2index100 = partAData_100pc.index
    number_of_results_photo2_100pc = len(p2index100)
    print("Photo 2 - 100%\n")
    file1.write("Photo 2 - 100%\n")
    print("Total number: " + str(number_of_results_photo2_100pc))
    file1.write("Total number: " + str(number_of_results_photo2_100pc)  + "\n")
    for row in partAData_100pc.index:
        print(row, end=" ")
        file1.write(str(row) + " ")
    n_avg, S_avg = calculateAverageValues(partAData_100pc)
    print("\nŚrednia ważona liczby śladów: " + str(n_avg))
    print("\nŚrednia ważona pola powierzchni: " + str(S_avg))
    file1.write("\nŚrednia ważona liczby śladów: " + str(n_avg) + "\n")
    file1.write("\nŚrednia ważona pola powierzchni: " + str(S_avg) + "\n")
    file1.write("\n\n")
    file1.write("\n\n")
    
    # 50%
    partAData_50pc = partAData[partAData.photo2_Bq <= (photo2_BestValue + (0.5 * photo2_BestValue))]
    partAData_50pc = partAData_50pc[partAData_50pc.photo2_Bq >= (photo2_BestValue - (0.5 * photo2_BestValue))] 
    #print(partAData_100pc)
    p2index50 = partAData_50pc.index
    number_of_results_photo2_50pc = len(p2index50)
    print("\nPhoto 2 - 50%\n")
    file1.write("\nPhoto 2 - 50%\n")
    print("Total number: " + str(number_of_results_photo2_50pc))
    file1.write("Total number: " + str(number_of_results_photo2_50pc) + "\n")
    for row in partAData_50pc.index:
        print(row, end=" ")
        file1.write(str(row) + " ")
    n_avg, S_avg = calculateAverageValues(partAData_50pc)
    print("\nŚrednia ważona liczby śladów: " + str(n_avg))
    print("\nŚrednia ważona pola powierzchni: " + str(S_avg))
    file1.write("\nŚrednia ważona liczby śladów: " + str(n_avg) + "\n")
    file1.write("\nŚrednia ważona pola powierzchni: " + str(S_avg) + "\n")
    file1.write("\n\n")
    
    # 20%
    partAData_20pc = partAData[partAData.photo2_Bq <= (photo2_BestValue + (0.2 * photo2_BestValue))]
    partAData_20pc = partAData_20pc[partAData_20pc.photo2_Bq >= (photo2_BestValue - (0.2 * photo2_BestValue))]
    #print(partAData_20pc)
    p2index20 = partAData_20pc.index
    number_of_results_photo2_20pc = len(p2index20)
    print("\nPhoto 2 - 20%")
    file1.write("\nPhoto 2 - 20%\n")
    print("Total number: " + str(number_of_results_photo2_20pc))
    file1.write("Total number: " + str(number_of_results_photo2_20pc) + "\n")
    for row in partAData_20pc.index:
        print(row, end=" ")
        file1.write(str(row) + " ")
    n_avg, S_avg = calculateAverageValues(partAData_20pc)
    print("\nŚrednia ważona liczby śladów: " + str(n_avg))
    print("\nŚrednia ważona pola powierzchni: " + str(S_avg))
    file1.write("\nŚrednia ważona liczby śladów: " + str(n_avg) + "\n")
    file1.write("\nŚrednia ważona pola powierzchni: " + str(S_avg) + "\n")
    file1.write("\n\n")
       
    # 5%
    partAData_5pc = partAData[partAData.photo2_Bq <= (photo2_BestValue + (0.05 * photo2_BestValue))]
    partAData_5pc = partAData_5pc[partAData_5pc.photo2_Bq >= (photo2_BestValue - (0.05 * photo2_BestValue))]
    #print(partAData_20pc)
    p2index5 = partAData_5pc.index
    number_of_results_photo2_5pc = len(p2index5)
    print("\nPhoto 2 - 5%")
    file1.write("\nPhoto 2 - 5%\n")
    print("Total number: " + str(number_of_results_photo2_5pc))
    file1.write("Total number: " + str(number_of_results_photo2_5pc) + "\n")
    for row in partAData_5pc.index:
        print(row, end=" ")
        file1.write(str(row) + " ")
    file1.write("\n\n")
    
    # 1%
    partAData_1pc = partAData[partAData.photo2_Bq <= (photo2_BestValue + (0.01 * photo2_BestValue))]
    partAData_1pc = partAData_1pc[partAData_1pc.photo2_Bq >= (photo2_BestValue - (0.01 * photo2_BestValue))]
    #print(partAData_20pc)
    p2index1 = partAData_1pc.index
    number_of_results_photo2_1pc = len(p2index1)
    print("\nPhoto 1 - 1%")
    file1.write("\nPhoto 2 - 1%\n")
    print("Total number: " + str(number_of_results_photo2_1pc))
    file1.write("Total number: " + str(number_of_results_photo2_1pc) + "\n")
    for row in partAData_1pc.index:
        print(row, end=" ")
        file1.write(str(row) + " ")
    file1.write("\n\n")   
            
    file1.close()
    
    return

def selectBestResultsPartB(DirectoryPath, partBData):
    photo3_BestValue = partBData.loc[0, 'photo3_Bq']
    photo4_BestValue = partBData.loc[0, 'photo4_Bq']
    print(photo3_BestValue)
    print(photo4_BestValue)
    
    ###################################
    # PHOTO 3
    file2 = open(DirectoryPath + "partB_best.txt","w+")
    file2.write("PHOTO 3 \n\n")
    
    # 100%
    partBData_100pc = partBData[partBData.photo3_Bq <= (2 * photo3_BestValue)]
    #print(partAData_100pc)
    p3index100 = partBData_100pc.index
    number_of_results_photo3_100pc = len(p3index100)
    print("Photo 3 - 100%\n")
    file2.write("Photo 3 - 100%\n")
    print("Total number: " + str(number_of_results_photo3_100pc))
    file2.write("Total number: " + str(number_of_results_photo3_100pc) + "\n")
    for row in partBData_100pc.index:
        print(row, end=" ")
        file2.write(str(row) + " ")
    file2.write("\n\n")
    
    # 50%
    partBData_50pc = partBData[partBData.photo3_Bq <= (photo3_BestValue + (0.5 * photo3_BestValue))]
    partBData_50pc = partBData_50pc[partBData_50pc.photo3_Bq >= (photo3_BestValue - (0.5 * photo3_BestValue))] 
    #print(partAData_100pc)
    p3index50 = partBData_50pc.index
    number_of_results_photo3_50pc = len(p3index50)
    print("\nPhoto 3 - 50%\n")
    file2.write("\nPhoto 3 - 50%\n")
    print("Total number: " + str(number_of_results_photo3_50pc))
    file2.write("Total number: " + str(number_of_results_photo3_50pc) + "\n")
    for row in partBData_50pc.index:
        print(row, end=" ")
        file2.write(str(row) + " ")
    file2.write("\n\n")
    
    # 20%
    partBData_20pc = partBData[partBData.photo3_Bq <= (photo3_BestValue + (0.2 * photo3_BestValue))]
    partBData_20pc = partBData_20pc[partBData_20pc.photo3_Bq >= (photo3_BestValue - (0.2 * photo3_BestValue))]
    #print(partAData_20pc)
    p3index20 = partBData_20pc.index
    number_of_results_photo3_20pc = len(p3index20)
    print("\nPhoto 3 - 20%")
    file2.write("\nPhoto 3 - 20%\n")
    print("Total number: " + str(number_of_results_photo3_20pc))
    file2.write("Total number: " + str(number_of_results_photo3_20pc) + "\n")
    for row in partBData_20pc.index:
        print(row, end=" ")
        file2.write(str(row) + " ")
    file2.write("\n\n")
       
    # 5%
    partBData_5pc = partBData[partBData.photo3_Bq <= (photo3_BestValue + (0.05 * photo3_BestValue))]
    partBData_5pc = partBData_5pc[partBData_5pc.photo3_Bq >= (photo3_BestValue - (0.05 * photo3_BestValue))]
    #print(partAData_20pc)
    p3index5 = partBData_5pc.index
    number_of_results_photo3_5pc = len(p3index5)
    print("\nPhoto 3 - 5%")
    file2.write("\nPhoto 3 - 5%\n")
    print("Total number: " + str(number_of_results_photo3_5pc))
    file2.write("Total number: " + str(number_of_results_photo3_5pc) + "\n")
    for row in partBData_5pc.index:
        print(row, end=" ")
        file2.write(str(row) + " ")
    file2.write("\n\n")
    
    ###################################
    # PHOTO 4
    file2.write("\nPHOTO 4 \n\n")

    # 100%
    partBData_100pc = partBData[partBData.photo4_Bq <= (2 * photo4_BestValue)]
    #print(partAData_100pc)
    p4index100 = partBData_100pc.index
    number_of_results_photo4_100pc = len(p4index100)
    print("Photo 4 - 100%\n")
    file2.write("Photo 4 - 100%\n")
    print("Total number: " + str(number_of_results_photo4_100pc))
    file2.write("Total number: " + str(number_of_results_photo4_100pc) + "\n")
    for row in partBData_100pc.index:
        print(row, end=" ")
        file2.write(str(row) + " ")
    file2.write("\n\n")
    
    # 50%
    partBData_50pc = partBData[partBData.photo4_Bq <= (photo4_BestValue + (0.5 * photo4_BestValue))]
    partBData_50pc = partBData_50pc[partBData_50pc.photo4_Bq >= (photo4_BestValue - (0.5 * photo4_BestValue))] 
    #print(partAData_100pc)
    p4index50 = partBData_50pc.index
    number_of_results_photo4_50pc = len(p4index50)
    print("\nPhoto 4 - 50%\n")
    file2.write("\nPhoto 4 - 50%\n")
    print("Total number: " + str(number_of_results_photo4_50pc))
    file2.write("Total number: " + str(number_of_results_photo4_50pc) + "\n")
    for row in partBData_50pc.index:
        print(row, end=" ")
        file2.write(str(row) + " ")
    file2.write("\n\n")
    
    # 20%
    partBData_20pc = partBData[partBData.photo4_Bq <= (photo4_BestValue + (0.2 * photo4_BestValue))]
    partBData_20pc = partBData_20pc[partBData_20pc.photo4_Bq >= (photo4_BestValue - (0.2 * photo4_BestValue))]
    #print(partAData_20pc)
    p4index20 = partBData_20pc.index
    number_of_results_photo4_20pc = len(p4index20)
    print("\nPhoto 4 - 20%")
    file2.write("\nPhoto 4 - 20%\n")
    print("Total number: " + str(number_of_results_photo4_20pc))
    file2.write("Total number: " + str(number_of_results_photo4_20pc) + "\n")
    for row in partBData_20pc.index:
        print(row, end=" ")
        file2.write(str(row) + " ")
    file2.write("\n\n")
       
    # 5%
    partBData_5pc = partBData[partBData.photo4_Bq <= (photo4_BestValue + (0.05 * photo4_BestValue))]
    partBData_5pc = partBData_5pc[partBData_5pc.photo4_Bq >= (photo4_BestValue - (0.05 * photo4_BestValue))]
    #print(partAData_20pc)
    p4index5 = partBData_5pc.index
    number_of_results_photo4_5pc = len(p4index5)
    print("\nPhoto 4 - 5%")
    file2.write("\nPhoto 4 - 5%\n")
    print("Total number: " + str(number_of_results_photo4_5pc))
    file2.write("Total number: " + str(number_of_results_photo4_5pc) + "\n")
    for row in partBData_5pc.index:
        print(row, end=" ")
        file2.write(str(row) + " ")
    file2.write("\n\n")
    
    file2.close()
    
    return



DirectoryPath = "E:\\A-- FUNDACJA\\PROJEKTY\\Szkolna Radonowa Mapa Polski\\Benchmark\\"
ListofExcelFiles = getListofExcelFiles(DirectoryPath, 'xls')
#print(ListofExcelFiles)

#pd.set_option('display.max_columns', None)
partAData = collectDataPartA(ListofExcelFiles)
#print(partAData)

#pd.set_option('display.max_columns', None)
partBData = collectDataPartB(ListofExcelFiles)
#print(partBData)

selectBestResultsPartA(DirectoryPath, partAData)
selectBestResultsPartB(DirectoryPath, partBData)