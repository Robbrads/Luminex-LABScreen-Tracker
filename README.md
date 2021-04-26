# Luminex-LABScreen-Tracker
VBA modules to tabulate and format Luminex LABScreen data for a single patient over multiple tests.  Allows visual tracking of inter-test MFI variation.
This has been developed for use in concert with HistoTrac LIMS. Luminex .csv files are imported to HistoTrac SQL and analysed. A patient's antibody history can then be exported to .xls to allow analysis on Luminex-LABScreen-Tracker.   

Prepare raw data in the following format:

Forename	SURNAME	Hospital_no					
*leave row B blank*							
Date	samplenbr	SingleAgBead	SingleAgRaw	SingleAgNormalized	SingleAgSpecificity	SingleAgSpecAbbr  TestTypeCd

Save as 'rptantibodychart.xls' to same folder as Antibody Tracking Master.  Run macro from control button on master sheet to parse, create pivot tables and format by MFI strenth.
