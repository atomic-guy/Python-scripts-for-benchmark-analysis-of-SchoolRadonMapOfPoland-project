import pandas as pd
import glob, os

def getListofExcelFiles(DirectoryPath, ext):
    '''
    Returns a list of files with specific extension stored in the given folder.
            Parameters:
                    DirectoryPath (str): Path to folder
                    ext (str): extension of files, e.g. 'xls'
            Returns:
                    ListofExcelFiles (list of str): List of files with specific extension with full path
    '''
    ListofExcelFiles = []
    for file in glob.iglob(DirectoryPath+'/**/*', recursive=True):
        if file.endswith(ext):
            ListofExcelFiles.append(file)
	return ListofExcelFiles

def printList(DirectoryPath, title, list):
    '''
    Prints and writes to txt file total number of items in the list.
    Prints and writes to txt file all items in the list.
            Parameters:
                    DirectoryPath (str): Path to folder where file with results will be stored
                    title (str): Title of file with results
                    list (list of str): list of strings
    '''
    i = 1
    filename = DirectoryPath + title + ".txt"
    file = open(filename, "w") 
    file.write("TOTAL NUMBER: " + str(len(list)) + "\n\n")   
    print("TOTAL NUMBER: " + str(len(list)) + "\n\n")
    for item in list:
        file.write(str(i) + '. ' + item + '\n') 
        print(str(i) + '. ' + item)
        i += 1
    file.close()
    return
 
def schoolsStatistics(ListofExcelFiles, DirectoryPath):
    '''
    Makes statistics of benchmark participants from Excel files. In Excel files first sheet, called "Informacje Ogólne"
    containes detailed data about participants: School's Name, City, Teacher's Name, List of Students (names).
    Prints list of schools, teachers and students participated in the benchmark.
    Use .xls files.
    Function uses printList() function
            Parameters:
                    ListofExcelFiles (list of str): List of .xls files with full path
                    DirectoryPath (str): Path to folder where file with results will be stored
    '''
    ListofSchools = []
    ListofTeachers = []
    ListofStudents = []
    
    for file in ListofExcelFiles:
        general_info_sheet = pd.read_excel(file, sheet_name="Informacje ogólne")
        #print(general_info_sheet)
        school = general_info_sheet.at[1,"Unnamed: 2"]
        city = general_info_sheet.at[2,"Unnamed: 2"]
        school_full = school + " - " + city
        ListofSchools.append(str(school_full))
        
        teacher = general_info_sheet.at[4,"Unnamed: 2"]
        ListofTeachers.append(str(teacher))
        
        students = general_info_sheet.loc[6:,['Unnamed: 2']].values.tolist()
        students_flat_list = [item for sublist in students for item in sublist]
        ListofStudents.extend(students_flat_list)
    
    ListofStudents = [item for item in ListofStudents if "Prosimy" not in item]
    ListofSchools = list(set(ListofSchools))
    ListofTeachers = list(set(ListofTeachers))
    
    print("\n")
    printList(DirectoryPath, title="lista_szkol", list=ListofSchools)
    print("\n")
    printList(DirectoryPath, title="lista_nauczycieli", list=ListofTeachers)
    print("\n")
    printList(DirectoryPath, title="lista_uczniow", list=ListofStudents)
    print("\n")
    
    #print(ListofSchools)
    #print("\n")
    #print(ListofTeachers)
    #print("\n")
    #print(ListofStudents)     
    
    #STATISTICS:
    print("STATISTICS:")
    print("Number of Schools: " + str(len(ListofSchools)))
    print("Number of Teachers: " + str(len(ListofTeachers)))
    print("Number of Students: " + str(len(ListofStudents)))
    print("Number of Excel files: " + str(len(ListofExcelFiles)))
    
    return

def generateGroups(DirectoryPath, ListofExcelFiles):
    '''
    Generates information about every group of students paricipated in the banchmark and writes it to 'groups.txt' file.
    Information about gruop consists of: number o gruop, school's name, city, teacher and list of students.
    Use .xls files.
            Parameters:
                    ListofExcelFiles (list of str): List of .xls files with full path
                    DirectoryPath (str): Path to folder where 'groups.txt' file will be stored
    '''
    listOfGroups = []
    for file in ListofExcelFiles:
        dict_with_data = {}
        
        file_name = os.path.basename(file)
        id = int(file_name[0:2])
        dict_with_data['id'] = id
        
        #------------------------------------------------------
        #general info
        general_info_sheet = pd.read_excel(file, sheet_name="Informacje ogólne")
        
        school = general_info_sheet.at[1,"Unnamed: 2"]
        city = general_info_sheet.at[2,"Unnamed: 2"]
        school_full = school + " - " + city
        dict_with_data['school'] = school_full
        
        teacher = general_info_sheet.at[4,"Unnamed: 2"]
        dict_with_data['teacher'] = teacher
        
        students = general_info_sheet.loc[6:,['Unnamed: 2']].values.tolist()
        students_flat_list = [item for sublist in students for item in sublist]
        dict_with_data['students'] = students_flat_list
        
        listOfGroups.append(dict_with_data)
    
    data = pd.DataFrame(listOfGroups)
    
    file = open(DirectoryPath + "groups.txt","w+")
    file.write("BENCHMARK, LISTA GRUP \n\n")
    
    n_index = data.index
    for i in n_index:
        file.write(str(i) + " -- " + str(data.loc[i]['school']) + "\n")
        file.write("Opiekun grupy: " + str(data.loc[i]['teacher']) + "\n")
        file.write("Uczniowie: ")
        for i in data.loc[i]['students']:
            file.write(str(i) + ", ")
        file.write("\n\n")
    file.close()
    return



DirectoryPath = "E:\\A-- FUNDACJA\\PROJEKTY\\Szkolna Radonowa Mapa Polski\\Benchmark\\"
ListofExcelFiles = getListofExcelFiles(DirectoryPath, 'xls')
#print(ListofExcelFiles)
#print("\n")
schoolsStatistics(ListofExcelFiles, DirectoryPath)
generateGroups(DirectoryPath, ListofExcelFiles)
