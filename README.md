# Luminex-LABScreen-Tracker
Macro-enabled workbook to tabulate and format Luminex LABScreen Single Antigen data for a single patient over multiple tests.  Allows visual tracking of inter-test MFI variation.

Create file for raw data: 'rptantibodychart.xls'.  Format as shown in the attached sample file.

Save to same folder as Antibody Tracking Master.  Open Master and click 'Create Chart' to parse, create pivot tables and format by MFI strenth.

Assemble raw data by exporting from your LIMS. The export file can contain any number of tests but must all be from the same patient.  Amend field names to match as needed.

TestTypeCd contains the name of each test, in this case 'LABScreen Single Antigen Class I/II'. Amend as necessary in export file.

The attached sample file 'rptantibodychart.xls' contains data from 3 tests to demonstrate functionality.  Download both files to the same folder, Open Master and click 'Create Chart'.
