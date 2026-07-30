# Anviz_report_processor
Software for converting the original Anviz AIM CrossChex system report into the required format.
The executable files are located in the dist folder.
The program versions differ only in the interface design.
Workflow:
1) Obtain the “event report by date” form from AIM CrossChex.
2) Convert the report to the modern XLSX format (save as ... -> XLSX).
3) Run the program.
3) Select the source data file obtained in step 2.
4) Select a file to save or enter its name (if you select an existing file,
all data in it will be lost).
5) Start processing and wait for it to finish.

Note: 
1) The position field is left blank because none of the system reports 
contain this information.
2) In case of an employee being late or leaving early, the corresponding 
cells are highlighted in red as in the example.
3) Before processing, the file MUST be converted to XLSX format, otherwise an error will occur.
