# excel_py_runner.py

Executes sequence of python functions, with optional arguments, read from a specially formatted spreadsheet.

Values returned from each function are written back to the spreadsheet.

At end of run a new "results" copy of the spreadsheet is saved to the results sub-directory. Filename is automatic and contains time & date.

Relies on the following:

- **excel_py_runner.py** - execute this. Reads action steps from the specially formatted spreadsheet.

- **py_runner.xlsx** - spreadsheet read by above. Action column contains names of Python functions to be executed.
Args column contains optional comma-separated arguments for the function.
Note one limitation is these are always read as text. Returned values from each function are saved to the Result column.
Scope of run can be adjusted using start row and end row values in cells C3 and C4.
Also individual rows can be skiped by placing "y" in the Skip column.

- **actions.py** - holds definitions of the functions called from the Actions column of the spreadsheet.
