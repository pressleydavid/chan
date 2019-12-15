import pandas as pd

if __name__ == "__main__":
    cdm = pd.read_excel("/Users/david/projects/cca/clients/ccr/chan/Gilead - SBC - DM 337-1431 - template.xlsx", sheet_name="CDM", header=1)
    st = pd.read_excel("/Users/david/projects/cca/clients/ccr/chan/Gilead - SBC - DM 337-1431 - template.xlsx", sheet_name="Study Totals")
    
    #write to csv
    cdm.to_csv("./cdm.csv")
    st.to_csv("./study_totals.csv")

    task_sections = st.iloc[4:16,0:2]
    task_sections.rename(columns={'VENDOR':'taskID', 'Catalyst Clinical Research':'task_label'},inplace=True)
    task_sections.set_index('taskID',inplace=True)
    print(task_sections)
    task_dict = task_sections.to_csv('data/intermed/task_sections.csv')
    print(task_dict)

    
    