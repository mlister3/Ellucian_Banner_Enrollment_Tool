# Ellucian_Banner_Enrollment_Tool
~~ Osceola_Prosper_Enrollment_Tool ~~
- - -

> Application Name: JN_OP_Tool.exe
> Current Version: 0.2 Alpha

### Update Notes:
1. Select from any query file in folder.
2. Loading percent while dependencies load. Ideal for seeing the exe load from off a USB.
3. Added improved figures for visualizations.
4. Added course loads figures to report.
5. Exports excel file for withdrawals and non-enrollees.
6. Rename to Banner Enrollment Tool due to curfuffle. See FERPA Compliance section.

### Summary of use:
- This app takes queries and outputs enrollment and admittance figures in the form of 3 items:
    - Visualizations - a number of pie charts can be saved. These charts check enrollment against enrollment of a prior semester or admission terms of a prior semester. All visualizations are made using Matplotlib.
    - Report - a txt file that shows numbers for enrollment for all terms available in the query, all admissions numbers for all terms in the query, answers to consistent specified questions regarding the query, and course load numbers for the selected term.
    - Excel of students that withdrew or were admitted and have not enrolled. These outputs are able to be handed to different departments to place academic holds, pursue outreach programs, or continue with further analysis.

1. Place Excel query into folder `QUERY_FILE_GOES_HERE`
2. Run JN_OP_Tool.exe and select the query for analysis when prompted.
3. When requested, review the list of terms available to set as the reference term.*
4. Pie charts will be provided
5. App will close automatically. report.txt will be created automatically and Student Withdrawls & Non-Enrollees.xlsx will also be created automatically.

> * Reference term is the term you are looking to get data for. Admitted in Summer 2023 and looking to get enrollment numbers for terms after, the reference term is summer 2023 or 202330.

- The JN_OP_Tool is a uncompiled executable made specifically for the computation, analysis, and visualization of the queries from Valencia College's Banner database. Users and/or editors of this application or any iteration after do not represent Valencia College. 

- Last successful compile was with CX_Freeze verion 6.14.2 - setup.py file is in the repository. To compile, ensure you have all necessary dependencies installed in your Python Environment. Install CX_Freeze using `pip` or `conda` and run the following command in your terminal `python setup.py build` to start the compile process. The user will need to create a folder called `QUERY_FILE_GOES_HERE` to place the Excel queries into.

- JN_OP_Tool was designed to run on locked pc's and all of the necessary Python libraries will be added to a library folder. In total, as of version 0.2 Alpha, the total size of the app is just under 2gb and can be ran off of a USB. 

FERPA Compliance:

- Student records are processed directly on the local machine and nothing is stored by the application. By deleting the initial query and exported files, no data remains. Data records transfered from one system to another has nothing to do with OP_Tool and users/editors are fully responsibile for the data they hold stewardship over. This app does not store any data and is not even functional until compiled or ran in a Python environment with all necessary dependencies. 

Past Versions:
0.1 Alpha
Notes:
1. Initial app

- - -

MIT License

Copyright (c) 2023 Matthew Lister
