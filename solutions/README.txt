Author: Ismael Ibrahim
Date: 17.08.2020
Content: xlsx parser for excel table to calcule the sum of a skill for a given date

---PYTHON-Solution---
For the python script the xlrd module is used.
It can be downloaded via the pip tool with the command
"pip install xlrd"

The Python file is in the folder "python_solution" the xlsx file must be located in the same folder.
Then you can run the program with Python
"excel_read.py"


---C#-Solution---
The C# solution runs with the NPOI package which can be installed with the NuGet package manager.
It is already in the project folder.
The xlsx file must be located 2 folders higher.
In the folder "csharp_solution/sumskill_excel/sumskill_excel/bin/Release" you find the runnable file
"sumskill_excel.exe"


---User-Descprition---
The program requests the parameters "Skill", "Year" and "Month" one after another.
You have to enter each of them.
If you enter a string instead of a number for year or month, it request the parameter again.
If he couldn't find a column with the entered date, then the program notifies the user. 
If everything is fine you receive the skill result.
Finally you can quit the programm by entering "q" 
or you can make a new run if you enter any other string.

