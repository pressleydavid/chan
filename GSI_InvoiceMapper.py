import pandas as pd
import openpyxl
from datetime import date
import os
import timeit
from timer import Timer

def unitMap(param1):
    ''' Map taskIDs to Gilead columns. Analagous to proc format'''
    #called below as: study['row'] = study['taskID'].apply(unitMap)
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
    #returns the _value_ of a dictionary {key:_value_} data structure (e.g. key is '8.0', value is 108)
    #applied to all items in a dataframe 
      return unitMap[param1]
    #handle KeyError exceptions (see: https://docs.python.org/3/library/exceptions.html#KeyError)
    except KeyError:
        print("Code not found in unitMap dict. Consider updating dict if "
              "value exists in dataframe")


if __name__ == "__main__":
    ''' Read in the all.xlsx sheet and map to the Gilead invoicer spreadsheet. '''

    all =pd.read_excel("all.xlsx", sheet_name="Sheet1", header=1)

    #make all.xlsx long, subset out one study
    studies = pd.melt(all,id_vars="Study", var_name='taskID')
    print(studies)
    studies['row'] = studies['taskID'].apply(unitMap)
    studies['col'] = 11
    print(studies)

    #FYI: This returns a list of the unique values of the 'Study' column of the studies dataframe
    #Probably doesn't matter which iterator you use (all_study vs projectID, study in studygrp), 
    # you still need to traverse the list from L to R to get all values 
    all_study = studies.Study.unique()
    print(all_study)

    #Note this is not a list. It is a pandas dataframe object
    studygrp = studies.groupby("Study")
    print(type(studygrp))
    
    for projectID, study in studygrp:
        print("Study:" + str(study))
        print("ProjectID: " + str(projectID))
        study.reset_index(drop=True, inplace=True)

        #parse file name to get relevant tokens for saving file. 
        # TODO: Note: this is hack. should obtain vendor, group, function, protocol from known
        #tasks, not a template file
        #filepath is bound to head as a string
        #filename is bound to tail as a string
        head, tail = os.path.split("Gilead - SBC - DM XXX-XXXX - template.xlsx")
    
        #printing to console is poor man's debugger
        print("Head: " + str(head))
        print("Tail: " + str(tail))

        #FYI
        #print out as a list
        print(tail.split(" - "))
        #print out each item in list 
        for item in tail.split(" - "):
            print(item)


        #split, defined with token (e.g. " - ") creates a 0-indexed list, referenced with array notation
        client = tail.split(" - ")[0]
        print(client)
        groupID = tail.split(" - ")[1]
        print(client)
        #note how "DM XXX-XXXX" split is a space,
        fxnID = tail.split(" - ")[2].split(" ")[0]
        print(fxnID)
        projID = tail.split(" - ")[2].split(" ")[1]
        print(projID)


        #load workbook template file
        wb = openpyxl.load_workbook("Gilead - SBC - DM XXX-XXXX - template.xlsx", keep_links=True)
        wb.template = False
        #FYI: have a look at the object type
        type(wb)
        cdm = wb["CDM"]
        #I actually don't know if the exception code will ever execute. I kinda just stashed it there
        #in case I needed to unfuck the merged cell issue openpyxl was not dealing with elegantly
        for row in study.itertuples():
            try:
                cell = cdm.cell(row=row.row, column=row.col, value=row.value)
            except:
                if type(cell).__name__ == 'MergedCell':
                    print(str(cell) +"is a merged cell." )
                else:
                    print(str(cell) +"is not a merged cell.")
        
        
        today = date.today()
        fileName = os.path.join(head, tail.replace('XXX-XXXX', str(projectID)).replace('template', str(today.strftime("%b %Y"))))
        print(fileName)
        wb.template = False
        wb.save(fileName)
        wb.close()

       

   
    
    # print(timeit.timeit("main()", setup="from __main__ import main"))