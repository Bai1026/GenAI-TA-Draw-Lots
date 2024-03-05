import pandas as pd
# import numpy as np

# # Load the registration data
df = pd.read_excel('./registration.xlsx')

# Ensure '學號' is a string and lowercase
df['學號'] = df['學號'].astype(str).str.lower().str.strip()

# already on the ntu cool
TA_df = pd.read_excel('./TAs.xlsx')
TA_df['學號'] = TA_df['學號'].str.lower()

drawed_df = pd.read_excel('./Drawed_50.xlsx')
drawed_df['學號'] = drawed_df['學號'].str.lower()

EECS_df = pd.read_excel('./EECS_and_Art.xlsx')
EECS_df['學號'] = EECS_df['學號'].str.lower()

# filter out TAs and 50 students that are already registered
df = df[~df['學號'].isin(TA_df['學號'])]
df = df[~df['學號'].isin(EECS_df['學號'])]
df = df[~df['學號'].isin(drawed_df['學號'])]
df = df[['Email Address', '姓名', '學號', '聯絡信箱', '是否有加簽或旁聽意願', '如果沒有加簽到的話，需要加入旁聽嗎']]

# =============== for wanting auditting ===============
audit_df = df[((df['是否有加簽或旁聽意願'] == '我沒有要加簽，只要旁聽') | (df['如果沒有加簽到的話，需要加入旁聽嗎'] == '要'))]
print(audit_df.head(10))
print("Auditting DONE!")
print()


student_df = pd.read_excel('./student_list.xlsx')
student_df['學號'] = student_df['學號'].str.lower()
student_df = student_df[['身份', '姓名', '信箱']]
student_df = student_df[student_df['身份'] == '旁聽生']
print("student list:")
print(student_df.head(10))
print()


# Filter out the student already in the ntu cool
audit_df['信箱'] = audit_df['Email Address'].str.lower()
student_df['信箱'] = student_df['信箱'].str.lower()

merged_df = pd.merge(audit_df, student_df, on='信箱', how='left', indicator=True)

filtered_df = merged_df[merged_df['_merge'] == 'left_only']

print(filtered_df.head(10))

filtered_emails = filtered_df['Email Address']

with open('filtered_emails.txt', 'w') as file:
    for email in filtered_emails:
        file.write(email + '\n')

print("Filtered emails saved to filtered_emails.txt")