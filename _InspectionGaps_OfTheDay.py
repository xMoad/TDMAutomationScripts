from openpyxl import *
import pandas as pd
from datetime import date


# Charger le fichier Excel source avec pandas
source_data = pd.read_excel('C:/Users/Mouad Zaid/Desktop/Tech/Python/RTR_RS_gaps_RN_SWEET400_AXS_EPBi.xlsx',sheet_name='No TC_TR')
source_data = source_data.fillna('0')

# Définir le motif de chaîne à rechercher et la colonne cible
Tday = date.today().strftime('%Y-%m-%d')

column_name_1 = 'Column7'
column_name_2 = 'Column2'
column_name_3 = 'Column5'
column_name_4 = 'Nom de la RS'


pattern_1 = 'Labcar/HIL'
pattern_2 = 'SIL Simulation'
pattern_3 = 'Vehicle'
pattern_4 = 'accepted'
Pattern_5 = 'Inspection'


#Generation HiL Gaps Excel file 
InspectionGaps = source_data[(source_data[column_name_1] == Pattern_5)&(source_data[column_name_3]==pattern_4)]
InspectionGaps = InspectionGaps[~HiLGaps[column_name_2].str.contains('Gen', na=False)]

InspectionGapsFile = f'InspectionGaps_OfTheDay_{Tday}.xlsx'
InspectionGaps.to_excel(InspectionGapsFile)

valUniqueInspecetion = InspectionGaps[column_name_2].unique()

for RSVal in valUniqueInspecetion:
    filterLine = InspectionGaps[InspectionGaps['Column2'] == RSVal]
    filName = f'_Inspection_{RSVal}.xlsx'
    filterLine.to_excel(filName, index=False)
    
#Create Metrics / KPI ==> For each function list Gaps and Responsible for 