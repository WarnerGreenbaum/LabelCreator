import pyautogui
import time
import csv
import os
import openpyxl
import shutil
from docxtpl import DocxTemplate


old_ports = []
new_ports = []

script_dir = os.path.dirname(os.path.abspath(__file__))


input_folder = os.path.join(script_dir, 'inputs')


files_in_folder = os.listdir(input_folder)


# Path to the Avery label template in the "template" folder
template_folder = os.path.join(script_dir, 'template_ignore_me')
template_path = os.path.join(template_folder, 'AveryTemplate.docx')  # Replace with your template file name

# Load the template
doc = DocxTemplate(template_path)

excel_file = next(file for file in files_in_folder if file.endswith('.xls') or file.endswith('.xlsx'))

excel_path = os.path.join(input_folder, excel_file)


workbook = openpyxl.load_workbook(excel_path)
sheet = workbook.active

data = []

for row in sheet.iter_rows(min_row=2, values_only=True):
    data.append(row)

workbook.close()

columns = list(zip(*data))

old_ports = columns[0]
new_ports = columns[1]


old_ports = [port for port in old_ports if port is not None]
new_ports = [port for port in new_ports if port is not None]


print("Old Ports:", old_ports)

print("New Ports:", new_ports)

switch_name = os.path.splitext(excel_file)[0]


# Populate the template with old_ports
context = {'old_ports': old_ports}
doc.render(context)

# Define the output folder path
output_folder = os.path.join(script_dir, 'output')
os.makedirs(output_folder, exist_ok=True)

# TESTING
# TESTING

# TESTING
# TESTING

# Define the output path for the populated document
output_name = switch_name + '.docx'
output_path = os.path.join(output_folder, output_name)




# Save the populated document
doc.save(output_path)

 

print(f"Populated labels saved at: {output_path}")



# time.sleep(2)

# c = 1

# pyautogui.write(switch_name)
# for i in new_ports:
#     if c % 2 == 0:
#         pyautogui.press('tab')
#     else: 
#         pyautogui.press('tab')
#         pyautogui.press('tab')  

#     c += 1
#     pyautogui.write(i)


#     TO DO: Have it write switch name automatically in top left sticker, have it take in csv, automatically filter it and fill in the doc and output it into a folder.