![alt text](https://github.com/WarnerGreenbaum/LabelCreator/blob/Label_Creator/template_ignore_me/networklogo.jpg?raw=true)

**What Does It Do?**
This tool is designed to automate the process of populating a label sheet based on a network switch cutsheet.

**Instructions:**

1. **Download and Setup:**
   - Download the repository and place it on your desktop.

2. **Install Dependencies:**
   - Open your command line or terminal.
   - Install the required dependencies using these commands:
     ```
     pip install openpyxl
     pip install docx
     pip install docxtpl
     ```

3. **Prepare Input:**
   - Locate the 'inputs' folder within the downloaded repository.
   - Place your Excel cutsheet into the 'inputs' folder.
   - There is a sample cutsheet ('mock-idf-vfsw.xlsx') inside the 'inputs' folder to demonstrate the required format.

4. **Run the Program:**
   - Open your command line or terminal.
   - Navigate to the directory where the Python file is located using the `cd` command. For example:
     ```
     cd Desktop/Label_Creator
     ```
   - Run the program using the command:
     ```
     py main.py
     ```
   - If successful, a prompt will display the number of sheets required for the labels.
   - The generated DOCX file will be located in the 'outputs' folder.
   - If an error occurs, carefully read the error message for troubleshooting.

5. **Print Onto Paper:**
   - Load your printer with either of the following sheets: 5366 or 45366. The sheets below are supposedly also compatible, but I haven't tested with them.
     (48266, 48366, 5029, 5566, 6505, 75366, 8066, 8366, 8478, 8590, 8593, Presta 94210)

**Important Notes:**
- This tool was developed quickly and may not adhere to the highest standards of code cleanliness.
- While functional, it may lack the polish of a fully refined product.
- The most common error may be related to formatting. Ensure that your old ports and new ports match up as shown in the example cutsheet ('mock-idf-vfsw.xlsx').
- Go beaves :)
