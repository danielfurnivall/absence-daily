import pandas as pd

df = pd.read_excel('C:/storyboards/newpivot - 2020-05-20.xlsx')

org_structure = pd.read_excel('W:/Master_Org_Structure/Org_Structure.xlsx')

print(org_structure.columns)
org_structure = org_structure.rename(columns={'department_reference':'Cost Centre'}, inplace=True)


df = df.merge(org_structure, on='Cost Centre', how='left')
df.to_excel('W:/MFT/ggh_abs.xlsx')