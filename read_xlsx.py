import pandas as pd
import openpyxl
from datetime import date
import os

def unitMap(param1):
    unitMap = {'1.0':4,
                '2.0':12,
                'CDM 2.0 A':18,
                'CDM 2.0 B':24,
                'CDM 2.0 C':30,
                'CDM 2.0 D':36,
                '3.0':44,
                'CDM 3.0 A':50,
                '4.0':58,
                'CDM 4.0 A':64,
                '5.0':72,
                '6.0':80,
                'CDM 6.0 A':86,
                '7.0':94,
                'CDM 7.0 A':100,
                '8.0':108,
                '9.0':116,
                '10.0':124,
                '11.0':132,
                '12.0':140
                }
    try:
    #returns the value of a dictionary {key:value} data structure
      return unitMap[param1]
    except KeyError:
        print("Code not found in unitMap dict. Consider updating dict if "
              "value exists in dataframe")


if __name__ == "__main__":
    ''' Read in the all.xlsx sheet and map to the Gilead invoicer spreadsheet. '''
    template_cdm = pd.read_excel("/Users/david/projects/cca/clients/ccr/chan/Gilead - SBC - DM 337-1431 - template.xlsx", sheet_name="CDM", header=1)
    template_st = pd.read_excel("/Users/david/projects/cca/clients/ccr/chan/Gilead - SBC - DM 337-1431 - template.xlsx", sheet_name="Study Totals")
    all =pd.read_excel("/Users/david/projects/cca/clients/ccr/chan/all.xlsx", sheet_name="Sheet1", header=1)

    #parse file name to get key for value in all.xlsx project
    
    head, tail = os.path.split("/Users/david/projects/cca/clients/ccr/chan/Gilead - SBC - DM 337-1431 - template.xlsx")
    print(tail)

    #print out as a list for visualization purposes
    for item in tail.split(" - "):
        print(item)

    client = tail.split(" - ")[0]
    print(client)
    groupID = tail.split(" - ")[1]
    print(client)
    fxnID = tail.split(" - ")[2].split(" ")[0]
    print(fxnID)
    projectID = tail.split(" - ")[2].split(" ")[1] 
    print(projectID)


    #write to csv
    template_cdm.to_csv("data/intermed/cdm.csv")
    template_st.to_csv("data/intermed/study_totals.csv")
    all.to_csv("data/intermed/all.csv")

    #load workbook template file
    wb = openpyxl.load_workbook("testing.xltx")
    # wb.template = False
    ws = wb['CDM']
    print(type(ws))

    #make all.xlsx long, subset out one study
    studies = pd.melt(all,id_vars="Study", var_name='taskID')
    study = studies[studies.Study == "337-1431"]
    # study.reset_index(drop=True, inplace=True)
    study['row'] = study['taskID'].apply(unitMap)
    study['col'] = 11
    print(study)

    task_sections = template_st.iloc[4:16,0:2]
    task_sections.rename(columns={'VENDOR':'taskID', 'Catalyst Clinical Research':'task_label'},inplace=True)
    task_sections.set_index('taskID',inplace=True)
    
    task_dict = task_sections.to_csv('data/intermed/task_sections.csv')

    invoicer = template_cdm
    invoicer.iat[1,10] = study["value"] 
    print(invoicer.index)

    for row in study.itertuples():
        try:
            cell = ws.cell(row=row.row, column=row.col, value=row.value)
        except:
            if type(cell).__name__ == 'MergedCell':
                print(str(cell) +"is a merged cell." )
            else:
                print(str(cell) +"is not a merged cell.")
        
        wb.template = False

        
        today = date.today()
        fileName = os.path.join(head, tail.replace('template', str(today.strftime("%b %Y"))))
        print(fileName)

        wb.save(



    
    