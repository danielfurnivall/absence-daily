import pandas as pd

df = pd.read_excel('W:/MFT/heather-piv.xls', sheet_name='Raw Data')
print(df.columns)


job_fam_piv = pd.pivot_table(df, index='JOB FAMILY', values='Candidate ID Number', aggfunc ='count', margins=True)


print(job_fam_piv)

cand_stat_piv = pd.pivot_table(df, index='Candidate Status', values='Candidate ID Number', aggfunc='count', margins=True)
print(cand_stat_piv)

job_stat_piv = pd.pivot_table(df, index='Job Status', values='Job Reference', aggfunc='count', margins=True)
print(job_stat_piv)