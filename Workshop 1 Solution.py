# -*- coding: utf-8 -*-

import pandas


# Bringing in the Excel files
excel_file_path = r'C:\Users\E100174\OneDrive - RSM\Software\Python\Workshop\Inputs\Data Tables.xlsx'
datatables = pandas.read_excel(excel_file_path)

excel_file_path_PGR = r'C:\Users\E100174\OneDrive - RSM\Software\Python\Workshop\Inputs\Product Group Reference.xlsx'
PGR = pandas.read_excel(excel_file_path_PGR)

excel_file_path_RF= r'C:\Users\E100174\OneDrive - RSM\Software\Python\Workshop\Inputs\Region Reference.xlsx'
RF = pandas.read_excel(excel_file_path_RF)

# Joining - inner bc similar to a join tool where you want what's coming out of the J rather than L or R. 
DT_PGR = pandas.merge(datatables,PGR, on='Product Group', how='inner')

DTPGR_RF = pandas.merge(DT_PGR, RF, on='Region', how='inner')

#Create output name column - from Jason's script. Creating output column names and specifying string type for each column using astype
DTPGR_RF['Output Name'] = "Region" + DTPGR_RF['Region Key'].astype(int).astype(str)+"_Product" + DTPGR_RF['Product Key'].astype(int).astype(str)

# Aggregate - you can agg by as many as you want, you'd just use a comma to list more in the squiggly brackets
Summarization = DTPGR_RF.groupby(['Region']).agg({'Client':'nunique'})

#Output
DataOutput = r'C:\Users\E100174\OneDrive - RSM\Software\Python\Workshop\Products_Output.xlsx'
TotalsOutput = r'C:\Users\E100174\OneDrive - RSM\Software\Python\Workshop\Clients_By_Region.xlsx'

with pandas.ExcelWriter(DataOutput, engine = 'xlsxwriter') as Writer:
        for Output in DTPGR_RF['Output Name'].unique():
            OutputDTPGR_RF = DTPGR_RF[DTPGR_RF['Output Name']== Output]
            OutputDTPGR_RF.to_excel(Writer, sheet_name = Output, index = False)
            
with pandas.ExcelWriter(TotalsOutput, engine = 'xlsxwriter') as Writer:
    Summarization.to_excel(Writer, sheet_name= 'Totals', index = False)