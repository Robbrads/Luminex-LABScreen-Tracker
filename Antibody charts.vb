Sub GetRawData()
Application.DisplayAlerts = False
'Open HistoTrac export
Dim FilePath As String
FilePath = Application.ActiveWorkbook.Path
Workbooks.Open Filename:=FilePath & "\rptantibodychart.xls"
Workbooks("rptantibodychart").Activate
'Concatenate sample number and sample date, separate with carriage return
Dim LastRow As Long
Dim I As Long
Dim myText As String
LastRow = Range("A" & Rows.Count).End(xlUp).Row
For I = 3 To LastRow
	Range("B" & I) = (Range("B" & I) & vbNewLine & Range("A" & I))
Next I
'Copy HistoTrac export into master sheet
Worksheets("sheet1").Range("A1:H10000").Copy
Workbooks("Antibody tracking master").Activate
ActiveSheet.Paste Destination:=Worksheets("Raw").Range("A1:H10000")
'Save as patient name and S number
Dim PatientForename As String
Dim PatientSurname As String
Dim PatientName As String
Dim SNumber As String
PatientForename = Sheets("Raw").Range("A1")
PatientSurname = Sheets("Raw").Range("B1")
SNumber = Sheets("Raw").Range("C1")
PatientName = PatientForename & " " & PatientSurname & " " & SNumber
Application.ActiveWorkbook.SaveAs Filename:=FilePath & "\" & PatientName
End Sub


Sub AssignControls()
Dim LastRow As Long
Dim I As Long
'Label neg and pos controls based on LABScreen bead number
LastRow = Range("A" & Rows.Count).End(xlUp).Row
For I = 2 To LastRow
    If (Range("C" & I).Value = "001") Then
        Range("F" & I).Value = "NEG control"
        Range("E" & I).Value = Range("D" & I).Value
    End If
Next I

I = 2
For I = 2 To LastRow
    If (Range("C" & I).Value = "002") Then
        Range("F" & I).Value = "POS control"
        Range("E" & I).Value = Range("D" & I).Value
    End If
Next I

End Sub


Sub DeleteBlankRows()
Columns("E").SpecialCells(xlCellTypeBlanks).EntireRow.Delete
End Sub


Sub RemoveButton()
Sheets("Raw").Buttons("Button 1").Delete
End Sub


Sub MergeDPDQ()
Worksheets("Class II").Activate
Application.DisplayAlerts = False
Dim LastRow As Integer
Dim StartRow As Integer
With ActiveSheet.Sort
     .SortFields.Add Key:=Range("B1"), Order:=xlAscending
     .SortFields.Add Key:=Range("C1"), Order:=xlAscending
     .SetRange Range("A1:Z10000")
     .Header = xlYes
     .Apply
End With
StartRow = 2
LastRow = Range("C" & Rows.Count).End(xlUp).Row
'Merge DQB1 row names with DQA1 
Dim StartMerge As Integer
StartMerge = StartRow
For I = StartRow + 1 To LastRow
        If Cells(I, 3) <> "" Then
            If Cells(I, 3) = Cells(I - 1, 3) Then
                Cells(I, 6) = Cells(I, 6) & " " & Cells(I - 1, 6)
                StartMerge = I
            End If
        End If
    Next I
'Delete DQA1 and DPA1 only rows 
Dim r As Integer
For r = ActiveSheet.UsedRange.Rows.Count To 1 Step -1
    If InStr(Cells(r, "G").Value, "DPA1") > 0 Then
        ActiveSheet.Rows(r).EntireRow.Delete
   ElseIf InStr(Cells(r, "G").Value, "DQA1") > 0 Then
        ActiveSheet.Rows(r).EntireRow.Delete
End If
Next
End Sub



Sub SplitClassI()
'Copy class I tests to new sheet based on test name
Dim LastRow As Long
Dim I As Long
LastRow = Range("A" & Rows.Count).End(xlUp).Row
For I = 2 To LastRow
    If (Range("H" & I).Value = "LABScreen Single Antigen Class I") Or (Range("H" & I).Value = "LABScreen Single Antigen Class I Ads") Then
    Range("H" & I).EntireRow.Copy Destination:=Worksheets("Class I").Range("A" & I)
    End If
Next I

Worksheets("Class I").Activate
Columns("E").SpecialCells(xlCellTypeBlanks).EntireRow.Delete

End Sub


Sub SplitClassII()
'Copy class II tests to new sheet based on test name
Worksheets("Raw").Activate
Dim LastRow As Long
Dim I As Long
LastRow = Range("A" & Rows.Count).End(xlUp).Row
For I = 2 To LastRow
    If (Range("H" & I).Value = "LABScreen Single Antigen Class II") Or (Range("H" & I).Value = "LABScreen Single Antigen Class II Ads") Then
    Range("H" & I).EntireRow.Copy Destination:=Worksheets("Class II").Range("A" & I)
    End If
Next I
Worksheets("Class II").Activate
Columns("E").SpecialCells(xlCellTypeBlanks).EntireRow.Delete
End Sub


Sub CreatePivotTableI()
Worksheets("Class I").Activate
Dim sht As Worksheet
Dim pvtCache As PivotCache
Dim pvt As PivotTable
Dim StartPvt As String
Dim SrcData As String
Dim DataRange As Range
Dim LastRow As Long
LastRow = Sheets("Class I").Cells.Find("*", searchorder:=xlByRows, searchdirection:=xlPrevious).Row
Set DataRange = Range("A1:H" & LastRow)
'Determine the data range you want to pivot
 SrcData = ActiveSheet.Name & "!" & DataRange.Address(ReferenceStyle:=xlR1C1)
'Create a new worksheet
 Set sht = Sheets.Add
'Where do you want Pivot Table to start?
 StartPvt = sht.Name & "!" & sht.Range("A3").Address(ReferenceStyle:=xlR1C1)
'Create Pivot Cache from Source Data
 Set pvtCache = ActiveWorkbook.PivotCaches.Create( _
    SourceType:=xlDatabase, _
    SourceData:=SrcData)
'Create Pivot table from Pivot Cache
 Set pvt = pvtCache.CreatePivotTable( _
    TableDestination:=StartPvt, _
    TableName:="PivotTable1")
pvt.HasAutoFormat = False
End Sub




Sub CreatePivotTableII()
Worksheets("Class II").Activate
Dim sht As Worksheet
Dim pvtCache As PivotCache
Dim pvt As PivotTable
Dim StartPvt As String
Dim SrcData As String
Dim DataRange As Range
Dim LastRow As Long
LastRow = Sheets("Class II").Cells.Find("*", searchorder:=xlByRows, searchdirection:=xlPrevious).Row
Set DataRange = Range("A1:H" & LastRow)
'Determine the data range you want to pivot
 SrcData = ActiveSheet.Name & "!" & DataRange.Address(ReferenceStyle:=xlR1C1)
'Create a new worksheet
 Set sht = Sheets.Add
'Where do you want Pivot Table to start?
 StartPvt = sht.Name & "!" & sht.Range("A3").Address(ReferenceStyle:=xlR1C1)
'Create Pivot Cache from Source Data
 Set pvtCache = ActiveWorkbook.PivotCaches.Create( _
    SourceType:=xlDatabase, _
    SourceData:=SrcData)
'Create Pivot table from Pivot Cache
 Set pvt = pvtCache.CreatePivotTable( _
    TableDestination:=StartPvt, _
    TableName:="PivotTable1")
pvt.HasAutoFormat = False
End Sub



Sub AddPivotFieldsI()
Worksheets("Sheet1").Activate
Dim pvt As PivotTable
Set pvt = ActiveSheet.PivotTables("PivotTable1")
'Add item to the Column Labels
pvt.PivotFields("samplenbr").Orientation = xlColumnField
'Add item to the Row Labels
pvt.PivotFields("singleagspecificity").Orientation = xlRowField
pvt.AddDataField pvt.PivotFields("singleAgnormalized")
pvt.ColumnGrand = False
pvt.RowGrand = False
End Sub



Sub AddPivotFieldsII()
Worksheets("Sheet2").Activate
Dim pvt As PivotTable
Set pvt = ActiveSheet.PivotTables("PivotTable1")
'Add item to the Column Labels
pvt.PivotFields("samplenbr").Orientation = xlColumnField
'Add item to the Row Labels
pvt.PivotFields("singleagspecificity").Orientation = xlRowField
pvt.AddDataField pvt.PivotFields("singleAgnormalized")
pvt.ColumnGrand = False
pvt.RowGrand = False
End Sub



Sub FormatTables()
Worksheets("Sheet1").Activate
'Hide pivot table text
Range("A3:B3").Font.Color = RGB(220, 230, 241)
Range("A4").Font.Color = RGB(220, 230, 241)
'Conditional formatting
Dim FormatRange As Range
Set FormatRange = Sheets("Sheet1").Range("B5:Z10000")
FormatRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlBetween, _
	Formula1:="=750", Formula2:="=1999"
FormatRange.FormatConditions(1).Interior.Color = vbYellow
FormatRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlBetween, _
	Formula1:="=2000", Formula2:="=4999"
FormatRange.FormatConditions(2).Interior.Color = vbBlue
FormatRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
	Formula1:="=4999"
FormatRange.FormatConditions(3).Interior.Color = vbRed
'Freeze top rows
Rows("5:5").Select
ActiveWindow.FreezePanes = True
'Set column width
Columns("A").ColumnWidth = 22
Range("B:Z").ColumnWidth = 10
Columns("A:W").HorizontalAlignment = xlCenter
Rows("4").WrapText = True
Worksheets("Sheet2").Activate
'Conditional formatting
Set FormatRange = Sheets("Sheet2").Range("B5:Z10000")
FormatRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlBetween, _
	Formula1:="=750", Formula2:="=1999"
FormatRange.FormatConditions(1).Interior.Color = vbYellow
FormatRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlBetween, _
	Formula1:="=2000", Formula2:="=4999"
FormatRange.FormatConditions(2).Interior.Color = vbBlue
FormatRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
	Formula1:="=4999"
FormatRange.FormatConditions(3).Interior.Color = vbRed
'Freeze top rows
Rows("5:5").Select
ActiveWindow.FreezePanes = True
'Set column width
Columns("A").ColumnWidth = 22
Range("B:Z").ColumnWidth = 10
Columns("A:W").HorizontalAlignment = xlCenter
Rows("4").WrapText = True
'Hide pivot table text
Range("A3:B3").Font.Color = RGB(220, 230, 241)
Range("A4").Font.Color = RGB(220, 230, 241)
End Sub





Sub PatientName()
Dim PatientForename As String
Dim PatientSurname As String
Dim PatientName As String
Dim SNumber As String
'Obtain patient name and S number from HistoTrac export file
Workbooks("rptantibodychart").Activate
PatientForename = Sheets("Sheet1").Range("A1")
PatientSurname = Sheets("Sheet1").Range("B1")
SNumber = Sheets("Sheet1").Range("C1")
PatientName = PatientForename & " " & PatientSurname & " " & SNumber
Workbooks("rptantibodychart").Close SaveChanges:=False
'Copy patient name to first cell of both table sheets. Format and align
ActiveWorkbook.Worksheets("Class I Table").Range("A1") = PatientName
Range("A1").Font.Bold = True
Range("A1").Font.Size = 14
Range("A1").HorizontalAlignment = xlLeft
Worksheets("Class II Table").Activate
Range("A1") = PatientName
Range("A1").Font.Bold = True
Range("A1").Font.Size = 14
Range("A1").HorizontalAlignment = xlLeft
Worksheets("Class I Table").Activate
End Sub



Sub RunAll()
Application.ScreenUpdating = False
Call GetRawData
Call RemoveButton
Call AssignControls
Call DeleteBlankRows
Call SplitClassI
Call CreatePivotTableI
Call AddPivotFieldsI
Call SplitClassII
Call MergeDPDQ
Call CreatePivotTableII
Call AddPivotFieldsII
Call FormatTables
'Rename table sheets and reorder
Sheets("Sheet1").Name = "Class I Table"
Sheets("Sheet2").Name = "Class II Table"
Sheets("Class II Table").Move Before:=Sheets("Raw")
Sheets("Class I Table").Move Before:=Sheets("Class II Table")
Call PatientName
Application.ScreenUpdating = True
ActiveWorkbook.Save
End Sub