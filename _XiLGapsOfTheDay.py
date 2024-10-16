from openpyxl import *
import pandas as pd
from datetime import date


# Charger le fichier Excel source avec pandas
source_data = pd.read_excel('C:/Users/Desktop/Tech/Python/RTRRSFile.xlsx',sheet_name='No TC_TR')
source_data = source_data.fillna('0')

# Définir le motif de chaîne à rechercher et la colonne cible
Tday = date.today().strftime('%Y-%m-%d')

column_name_1 = 'Column7'
column_name_2 = 'Column2'
column_name_3 = 'Column5'

pattern_1 = 'Labcar/HIL'
pattern_2 = 'SIL Simulation'
pattern_3 = 'Vehicle'
pattern_4 = 'accepted'


#Generation HiL Gaps Excel file 
HiLGaps = source_data[(source_data[column_name_1] == pattern_1)&(source_data[column_name_3]==pattern_4)]
HiLGaps = HiLGaps[~HiLGaps[column_name_2].str.contains('Gen', na=False)]

HiLGapsFile = f'HiLGaps_OfTheDay_{Tday}.xlsx'
HiLGaps.to_excel(HiLGapsFile)

valUniqueHiL = HiLGaps[column_name_2].unique()

for RSVal in valUniqueHiL:
    filterLine = HiLGaps[HiLGaps['Column2'] == RSVal]
    filName = f'_HiL_{RSVal}.xlsx'
    filterLine.to_excel(filName, index=False)

#Generation SiL Gaps Excel file 
SiLGaps = source_data[(source_data[column_name_1] == pattern_2)&(source_data[column_name_3]==pattern_4)]
SiLGaps = SiLGaps[~SiLGaps[column_name_2].str.contains('Gen', na=False)]

SiLGapsFile = f'SiLGaps_OfTheDay_{Tday}.xlsx'
SiLGaps.to_excel(SiLGapsFile)

valUniqueSiL = SiLGaps[column_name_2].unique()

for RSVal in valUniqueSiL:
    filterLine = SiLGaps[SiLGaps['Column2'] == RSVal]
    filName = f'_SiL_{RSVal}.xlsx'
    filterLine.to_excel(filName, index=False)
    
    
#Generation VTC Gaps Excel file 
VTCGaps = source_data[(source_data[column_name_1] == pattern_3)&(source_data[column_name_3]==pattern_4)]
VTCGaps = VTCGaps[~VTCGaps[column_name_2].str.contains('Gen', na=False)]

VTCGapsFile = f'VTCGaps_OfTheDay_{Tday}.xlsx'
VTCGaps.to_excel(VTCGapsFile)

valUniqueVTC = VTCGaps[column_name_2].unique()

for RSVal in valUniqueVTC:
    filterLine = VTCGaps[VTCGaps['Column2'] == RSVal]
    filName = f'_VTC_{RSVal}.xlsx'
    filterLine.to_excel(filName, index=False)
