What does it do?
This is created to automatically populate a label sheet given a network switch cutsheet.

1.
  Download the repo and put it on your desktop

2.
Install the dependencies:
pip install openpyxl
pip install docx
pip install docxtpl

3.
  Drop your excel cutsheet into the inputs folder. (There is a mock cutsheet inside the input folder to show how it should be formatted).

4. Run the program from command line using py main.py (if confused by this step: Open up cmd, navigate to where the python file is located, you can navigate directories by using dir to display the directories and change directories using cd <directory to move to>)
   
   ex:
   cd Desktop/Label_Creator

   py main.py
   
7. If successful a prompt will let you know how many sheets you need for the labels and the docx will be located in the outputs folder. If there is an error please read it.

NOTES:
This was a tool made in a very short time in a not so clean way. It works well for what it is but is by no means a polished product.
-The one error that will most likely occur is with formatting, make sure your old ports and new ports match up like the example cutsheet mock-idf-vfsw
