# Luminex-LABScreen-Tracker
Macro-enabled workbook to tabulate and format Luminex LABScreen Single Antigen data for a single patient over multiple tests.  Allows visual tracking of inter-test MFI variation.

Create file for raw data: 'rptantibodychart.xls'.  Format as below:

Forename	SURNAME	Hospital_no					
*leave row B blank*							
Date	samplenbr	SingleAgBead	SingleAgRaw	SingleAgNormalized	SingleAgSpecificity	SingleAgSpecAbbr  TestTypeCd

Save to same folder as Antibody Tracking Master.  Open Master and click 'Create Chart' to parse, create pivot tables and format by MFI strenth.

Assemble raw data by exporting from your LIMS. The export file can contain any number of tests but must all be from the same patient.  Amend field names to match as needed.

TestTypeCd contains the name of each test, in this case 'LABScreen Single Antigen Class I/II'. Amend as necessary in export file.
