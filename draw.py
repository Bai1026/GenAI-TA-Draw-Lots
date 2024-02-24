import pandas as pd
import numpy as np

# Load the data from the Excel file
df = pd.read_excel('./registration.xlsx')

# ===== for auditting =====
audit_df = df[df['是否有加簽或旁聽意願'] == '我沒有要加簽，只要旁聽']
audit_df = audit_df[['Email Address', '姓名', '學號', '聯絡信箱', '系所 (含雙主修、輔系、學程)', '是否有加簽或旁聽意願']]
audit_df.to_excel('Audit.xlsx', index=False)
print("Audit DONE!")

# ===== for EECS and Art =====
remaining_df = df[(df['系所 (含雙主修、輔系、學程)'] != '其他') & (df['是否有加簽或旁聽意願'] != '我沒有要加簽，只要旁聽')]
remaining_df = remaining_df[['Email Address', '姓名', '學號', '聯絡信箱', '系所 (含雙主修、輔系、學程)', '是否有加簽或旁聽意願']]
remaining_df.to_excel('EECS_and_Art.xlsx', index=False)
print("EECS and ART DONE!")

# ===== for Others who wanna registry =====
# selected_others = df[(df['系所 (含雙主修、輔系、學程)'] == '其他') & (df['是否有加簽或旁聽意願'] != '我沒有要加簽，只要旁聽')]
others_df = df[(df['系所 (含雙主修、輔系、學程)'] == '其他') & (df['是否有加簽或旁聽意願'] != '我沒有要加簽，只要旁聽')]
selected_others = others_df.sample(n=min(10, len(others_df)), random_state=1)
selected_df = selected_others[['Email Address', '姓名', '學號', '聯絡信箱', '系所 (含雙主修、輔系、學程)', '是否有加簽或旁聽意願']]
selected_df.to_excel('Drawed_50.xlsx', index=False)
print("Draw DONE!")

# print(selected_df)
# Save the final dataframe to a new Excel file

# print(final_df)
