This is my first project for Rubix Data Sciences where I have created an application to traverse through large Excel files and simplifies work by allowing the user to easily select multiple sheets and columns, then save them into a new file.

This program is a sheet and column selection tool for Excel files, built using Tkinter and CustomTkinter for the graphical interface. Here's a quick breakdown of what it does:

File Selection: It prompts the user to select an Excel file (with .xlsx or .xls extension).

Sheet Listing: After the file is selected, it loads the sheet names and displays them in a Listbox. The sheet names are grouped alphabetically by their starting letter (A, B, C, etc.) to make it easier for the user to browse.

Multiple Sheet Selection: The user can select multiple sheets from the list.

Column Selection: Once sheets are selected, the program opens a new window where the user can choose specific columns from the selected sheets. It shows the columns of each sheet for easy selection.

Download Selected Data: After selecting the desired columns, the user can click a "Download" button to save the selected columns into a new Excel file, with each sheet containing only the selected columns.





