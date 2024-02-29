import pandas as pd
# import numpy as np

TA_df = pd.read_excel('./TAs.xlsx')
TA_df['學號'] = TA_df['學號'].str.lower()
print("TA list:")
# print(TA_df)

drawed_df = pd.read_excel('./Drawed_50.xlsx')
drawed_df['學號'] = drawed_df['學號'].str.lower()
print(drawed_df)
print()

# Load the registration data
df = pd.read_excel('./registration.xlsx')

# Ensure '學號' is a string and lowercase
df['學號'] = df['學號'].astype(str).str.lower().str.strip()

# filter out TAs and 50 students that are already registered
df = df[~df['學號'].isin(TA_df['學號'])]
# df = df[~df['學號'].isin(drawed_df['學號'])]
df = df[['Email Address', '姓名', '學號', '聯絡信箱', '系所 (含雙主修、輔系、學程)', '是否有加簽或旁聽意願', '如果沒有加簽到的話，需要加入旁聽嗎']]

# =============== for Auditting ===============
# audit_df = df[df['是否有加簽或旁聽意願'] == '我沒有要加簽，只要旁聽']
audit_df = df[(df['是否有加簽或旁聽意願'] == '我要加簽') & (df['如果沒有加簽到的話，需要加入旁聽嗎'] == '要') & (df['系所 (含雙主修、輔系、學程)'] == '其他')]
audit_df.to_excel('Auditting.xlsx', index=False)
print(audit_df.head(10))
print("Auditting DONE!")


# =============== for Deduplicating ===============
# df = df[df['是否有加簽或旁聽意願'] != '我沒有要加簽，只要旁聽']
df = df[(df['是否有加簽或旁聽意願'] != '我沒有要加簽，只要旁聽') & (df['是否有加簽或旁聽意願'] != '沒有要加簽，也沒有要旁聽') & (df['學號'] != '無') & (df['學號'] != '无') & (df['聯絡信箱'].str.contains('ntu.edu.tw'))]
df = df.drop_duplicates(subset='學號', keep='first')

# =============== for EECS and Art ===============
remaining_df = df[df['系所 (含雙主修、輔系、學程)'] != '其他']
remaining_df.to_excel('EECS_and_Art.xlsx', index=False)
print("EECS and ART DONE!")

# =============== for Others who wanna registry ===============
others_df = df[df['系所 (含雙主修、輔系、學程)'] == '其他']
selected_df = others_df.sample(n=min(50, len(others_df)), random_state=74774)
selected_df.to_excel('Drawed_50.xlsx', index=False)
print("Draw DONE!")