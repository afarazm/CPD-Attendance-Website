import pandas as pd

excel_ehc_directory = '/Users/arhamfaraz/Downloads/CPD/DirectoryList.xlsx'
new_excel = pd.read_excel(excel_ehc_directory)

form_submission_sharepoint_excel = '/Users/arhamfaraz/Downloads/CPD/CPDDataList.xlsx'
new_form = pd.read_excel(form_submission_sharepoint_excel)

# if 'Email Address' not in new_form.columns:
#     raise KeyError("'Email Address' column not found in new_form")
# if 'Work Email' not in new_excel.columns:
#     raise KeyError("'Work Email' column not found in new_excel")

merged_excel = pd.merge(new_form, new_excel[['Work Email', 'Name', 'Title', 'Specialty', 'Mobile No.', 'QID', 'Medical License']], left_on='Email Address', right_on='Work Email', how='left')

merged_excel.drop(columns=['Work Email', 'Item Type', 'Path', 'Submitted By'], inplace=True)

columns = ['Name'] + [col for col in merged_excel if col != 'Name']
merged_excel = merged_excel[columns]

final_file = '/Users/arhamfaraz/Downloads/CPD/CPD Attendance Data for MOPH.xlsx'
merged_excel.to_excel(final_file, index=False)

print(f'Merged CPD data saved to {final_file}')





