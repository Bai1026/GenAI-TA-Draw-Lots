import pandas as pd
import numpy as np

TA_df = pd.read_excel('./TAs.xlsx')
TA_df['學號'] = TA_df['學號'].str.lower()
print("TA list:")
print(TA_df)

# Load the registration data
df = pd.read_excel('./registration.xlsx')

# Ensure '學號' is a string and lowercase
df['學號'] = df['學號'].astype(str).str.lower().str.strip()
df = df[~df['學號'].isin(TA_df['學號'])]
df = df[['Email Address', '姓名', '學號', '聯絡信箱', '系所 (含雙主修、輔系、學程)', '是否有加簽或旁聽意願']]

# =============== for Auditting ===============
audit_df = df[df['是否有加簽或旁聽意願'] == '我沒有要加簽，只要旁聽']
audit_df.to_excel('Auditting.xlsx', index=False)
print("Auditting DONE!")

# =============== for Deduplicating ===============
df = df[df['是否有加簽或旁聽意願'] != '我沒有要加簽，只要旁聽']
df = df.drop_duplicates(subset='學號', keep='first')

# =============== for EECS and Art ===============
remaining_df = df[df['系所 (含雙主修、輔系、學程)'] != '其他']
remaining_df.to_excel('EECS_and_Art.xlsx', index=False)
print("EECS and ART DONE!")

# =============== for Others who wanna registry ===============
others_df = df[df['系所 (含雙主修、輔系、學程)'] == '其他']
selected_df = others_df.sample(n=min(10, len(others_df)), random_state=74774)
selected_df.to_excel('Drawed_50.xlsx', index=False)
print("Draw DONE!")