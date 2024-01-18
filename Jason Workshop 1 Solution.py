# Libraries #
import pandas
import os

# Input Location #
InputFolder = r'\\AZDEVALTYXENT01\AlteryxDevShare\Python_Upskilling\Workshop 1\Inputs'

# File Names #
DataTables = 'Data Tables.xlsx'
ProductGroupReference = 'Product Group Reference.xlsx'
RegionReference = 'Region Reference.xlsx'

# Import Data - When there's only one sheet, don't need to specify sheet name #
DataTablesDF = pandas.read_excel(io = os.path.join(InputFolder, DataTables), sheet_name = 'Data Tables')
ProductGroupReferenceDF = pandas.read_excel(io = os.path.join(InputFolder, ProductGroupReference))
RegionReferenceDF = pandas.read_excel(io = os.path.join(InputFolder, RegionReference))

# Join Data #
DF = pandas.merge(left = DataTablesDF, right = RegionReferenceDF, how = 'inner', on = ['Region'])
DF = pandas.merge(left = DF, right = ProductGroupReferenceDF, how = 'inner', on = ['Product Group'])

# Create output name column #
DF['Output Name'] = "Region" + DF['Region Key'].astype(int).astype(str) + "_Product" + DF['Product Key'].astype(int).astype(str)

# Aggregation #
TotalsDF = DF.groupby(['Region']).agg({'Client': 'nunique'}).rename(columns = {'Client': 'Client Count'}).reset_index()

# Output Location #
ProductsOutput = r'\\AZDEVALTYXENT01\AlteryxDevShare\Python_Upskilling\Jason\Outputs\JF_Products_Output.xlsx'
TotalsOutput = r'\\AZDEVALTYXENT01\AlteryxDevShare\Python_Upskilling\Jason\Outputs\JF_Totals_Output.xlsx'

# Output for products, using loop to create each tab #
with pandas.ExcelWriter(ProductsOutput, engine = 'xlsxwriter') as Writer:

    # Write each DataFrame to a different worksheet.
    for Output in DF['Output Name'].unique():
        OutputDF = DF[DF['Output Name'] == Output]
        OutputDF.to_excel(Writer, sheet_name = Output, index = False)

# Output for totals #
with pandas.ExcelWriter(TotalsOutput, engine = 'xlsxwriter') as Writer:

    TotalsDF.to_excel(Writer, sheet_name = 'Totals', index = False)
