import pandas as pd
import docx
import pprint
from datetime import datetime
from os import listdir
from docx.shared import Pt
from docx.shared import Inches
from docx.enum.section import WD_ORIENT
import warnings
warnings.filterwarnings('ignore')

def clean_spaces(question):
    return " ".join(question.replace('\xa0', '').split())

def clean_spaces2(question):
    while question[-1].isnumeric():
        question = question[:-1]
    return " ".join(question.replace('\xa0', '').split()).replace("’", "'")

'''def similar(a, b):
    return SequenceMatcher(None, a, b).ratio()'''

normed_files = []

def standardNormalisation():

    file_list = listdir("/Users/stewartrobertson/Documents/Redfield and Wilton/10. Projects/2. Multi-reg Cover Sheet/2. Region files")
    file_list = [fle for fle in file_list if fle != '.DS_Store']
    pprint.pprint(file_list)
    
    for fle in file_list:
        if fle[:4].isnumeric():
            dt = fle.split(" - ")[0]
            date = datetime.strptime(dt, '%Y-%m-%d').date()
            reg = fle.split(" - ")[1]
        else:
            dt = str(input('Enter the date (DD/MM/YYYY): '))
            date = datetime.strptime(dt, '%d/%m/%Y').date()
            reg = fle.split(" ")[0]
        xls = pd.ExcelFile("/Users/stewartrobertson/Documents/Redfield and Wilton/10. Projects/2. Multi-reg Cover Sheet/2. Region files/" + str(fle))
        sheets = xls.sheet_names

        Doc = []
        Cat = []
        q_all = []
        c = 1

        # Create list of all questions, and rearrange sheets into separate tables
        for x in range(len(sheets)):
            doc = pd.read_excel(xls, sheet_name = sheets[x], header=None)
            q_sheet = list(doc.iloc[:,0].dropna())
            q_all.extend(q_sheet)

            t = list(doc.iloc[:, 3]).count('Total')
            if t > 1:
                Total_ind = [i for i, j in enumerate(doc.iloc[:,3]) if j == "Total"]

                Total_ind.append(len(doc)+1)
                for y in range(t):
                    doc.iloc[Total_ind[y]+1, 3] = "Total"
                    table = doc.iloc[Total_ind[y]:Total_ind[y+1]-1]
                    Table = table.dropna(how='all', axis=1)
                    Doc.append(Table)
                    Cat.append(sheets[x])
            if t == 1:
                doc.iloc[2, 3] = "Total"
                Doc.append(doc.iloc[1:])
                Cat.append(sheets[x])

                
        q_all = [clean_spaces(ques) for ques in q_all]
        # Loop through each table, and normalise the data
        All_data = []
        Duplicates = []
        Excluded = []
        for a in range(len(Doc)):

            doc = Doc[a]
            # Identify question cells and prepare the crosstab headings a.k.a "Attribute"
            q_ind = [i for i, j in enumerate(doc.iloc[:,0]) if pd.isnull(j) == False]
            init = doc.iloc[:2, 3:]
            for x in range(1, len(init.iloc[0,:])):
                if pd.isnull(init.iloc[0,x]) == True:
                    init.iloc[0,x] = init.iloc[0,x-1]

            for x in range(1, len(init.iloc[0,:])):
                if pd.isnull(init.iloc[1,x]) == False:
                    init.iloc[0,x] = clean_spaces(init.iloc[0,x])
                    init.iloc[1,x] = str(init.iloc[0,x]) + ': ' + str(init.iloc[1,x])

            init = init.iloc[1, :].values.tolist()

            Data = []
            for x in range(len(q_ind)-1):
                Data.append(doc.iloc[q_ind[x]:q_ind[x+1]])
            Data.append(doc.iloc[q_ind[-1]:])

            Sheet = []
            for x in range(len(Data)):
                table = Data[x]
                q = clean_spaces(table.iloc[0,0])
                answers = table[1].unique()[2:]
                # commenting out the renaming of duplicates
                if q_all.count(q) > 1:
                    if q not in Duplicates and q not in Excluded:
                        print(q)
                        renameDuplicates = str(input("Would you like to rename these duplicates? (y/n) "))
                        if renameDuplicates.lower() == "y":
                            Duplicates.append(q)
                        elif renameDuplicates.lower() == "n":
                            Excluded.append(q)
                    if q in Duplicates:
                        print(q)
                        print(answers)
                        rename = str(input("Enter the replacement: "))
                        q = q + " " + rename
                        q = clean_spaces(q)
                    else:
                        q = q + str(c)
                        c += 1
                data = table.iloc[:, 3:]

                dr = []
                for y in range(data.shape[0]):
                    if data.iloc[y].isnull().all() == True:
                        dr.append(y)
                data = data.reset_index(drop=True)
                data = data.drop(dr)

                r1 = list(data.iloc[0,:])
                r2 = list(data.iloc[1,:])
                R3 = list(data.iloc[2,:])
                R4 = list(data.iloc[3,:])
                R1 = r1.copy()
                R2 = r2.copy()
                Attr = init.copy()

                for y in range(int((len(data) - 4)/2)):
                    Attr.extend(init)
                    R1.extend(r1)
                    R2.extend(r2)
                    R3.extend(list(data.iloc[4+2*y,:]))
                    R4.extend(list(data.iloc[5+2*y,:]))

                Ans = []
                Date = []
                Ques = []
                QC = []
                
                for y in range(len(answers)):
                    for z in range(len(init)):
                        Ans.append(answers[y])
                        Ques.append(str(q))
                        Date.append(date)
                        QC.append(str(Cat[a]))
                df = pd.DataFrame(list(zip(QC, Ques, Date, Ans, Attr, R3, R4, R2, R1)), columns = ['Sheet Name', 'Question', 'Date', 'Answer', 'Attribute', '% Votes', 'Votes', 'Weighted Total', 'Unweighted Total'])
                Sheet.append(df)
            Dataframe = pd.concat(Sheet)
            Dataframe = Dataframe.replace({"’": "'"}, regex = True)
            All_data.append(Dataframe)

        # Concatenate all data into a single dataframe, and export as an excel file
        D = pd.concat(All_data)
        name = date.strftime('%y.%m.%d') + " " + reg
        D.to_excel(name +'.xlsx', sheet_name='All Data', header=True, index=False)
        normed_files.append(name + '.xlsx')
        print("Data normalised")
    return sorted(normed_files)

standardNormalisation()