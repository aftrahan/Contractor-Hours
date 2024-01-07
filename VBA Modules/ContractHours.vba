
Option Explicit

Public startDate As Date, endDate As Date, master As Workbook
Public cons As Worksheet, emps As Worksheet, dets As Worksheet

Sub AllInFolder():

'***************************************************************************
'Purpose: Import data from all WO files in a given folder (excluding
'         subfolders)
'Inputs:  None
'Outputs: None
'***************************************************************************

'Sections of this sub dealing with sequential file access were adapted
'from: www.TheSpreadsheetGuru.com

Dim myPath As String, myFile As String, myExtension As String
Dim FldrPicker As FileDialog, ws As Worksheet, wb As Workbook
Dim empStart As Long, conStart As Long, rng As Range
Dim woname As String

Set cons = ThisWorkbook.Sheets("Contractors")
Set emps = ThisWorkbook.Sheets("Employees")
Set dets = ThisWorkbook.Sheets("Details")

startDate = ThisWorkbook.Sheets(1).Range("PP_Start").Value
Call checkDate
endDate = startDate + 14 'this picks up the date after, to account for
                         '"overnight" hours on the final date
                         
'Find the starting row on each sheet
If cons.Range("A2").Value = "" Then
    conStart = 2
Else
    conStart = cons.Range("A1").End(xlDown).Row + 1
End If

If emps.Range("A2").Value = "" Then
    empStart = 2
Else
    empStart = emps.Range("A1").End(xlDown).Row + 1
End If

Application.ScreenUpdating = False
Application.EnableEvents = False
Application.Calculation = xlCalculationManual

Set master = Workbooks.Open(Application.ThisWorkbook.Path & "\MASTER LISTS.xlsm")

Set FldrPicker = Application.FileDialog(msoFileDialogFolderPicker)

    With FldrPicker
        .Title = "Select Target Folder"
        .AllowMultiSelect = False
            If .Show <> -1 Then GoTo NextCode
            myPath = .SelectedItems(1) & "\"
    End With
    
NextCode:
    
myPath = myPath
If myPath = "" Then GoTo ResetSettings
myExtension = "*.xls*"

myFile = Dir(myPath & myExtension)

Do While myFile <> ""
    woname = Left(myFile, InStr(myFile, ".") - 1)
    Set wb = Workbooks.Open(Filename:=myPath & myFile)
        DoEvents
    Set ws = wb.Worksheets(1)
    
    'Check if WO needs to be added to list
    Set rng = dets.ListObjects("WO").ListColumns(1).Range.Find(woname)
    If rng Is Nothing Then
        Call AddWo(woname)
    End If
    
    Call ReadSheet(ws)
    wb.Close False
        DoEvents
    myFile = Dir
Loop

Call checkWorkers(conStart, empStart)
Call checkWOs
Call sortAll

master.Close True
    
ResetSettings:

Application.ScreenUpdating = True
Application.EnableEvents = True
Application.Calculation = xlCalculationAutomatic

End Sub

Sub ReadSheet(src):

'***************************************************************************
'Purpose: Transfer data from source file to times spreadsheet
'Inputs:  src - source file
'Outputs: None
'***************************************************************************

Dim srcRow As Long, lastDate As Date, runDate As Date
Dim tbl As ListObject, foundRow As Long

'Check if WO has already been imported; if so there will be a 'last date'
'value to ensure only "new" data is imported
Set tbl = dets.ListObjects("WO")
foundRow = tbl.ListColumns(1).Range.Find(src.Cells(2, 5).Value).Row

If IsEmpty(tbl.Range(foundRow, 4).Value) Then
    lastDate = 0
Else
    lastDate = tbl.Range(foundRow, 4).Value
End If

srcRow = 2
runDate = lastDate

'Script will ignore rows that have already been imported (determined by
'lastDate value) and ignore rows that were added to DB by user ADMINASST
'as these will be edits made by spreadsheet user that will already be
'accounted for.
Do Until src.Cells(srcRow, 3).Value = ""
    If src.Cells(srcRow, 13) > lastDate And src.Cells(srcRow, 14) <> "ADMINASST" Then
        Call WriteLine(src, srcRow)
        If src.Cells(srcRow, 13) > runDate Then
            runDate = src.Cells(srcRow, 13).Value
        End If
    End If
    srcRow = srcRow + 1
Loop

With tbl
    .Range(foundRow, 4).Value = runDate
End With

End Sub

Sub WriteLine(src, srcRow):
'***************************************************************************
'Purpose: Transfer specific row to times spreadsheet
'Inputs:  src - source file
'         srcRow - row to be read
'Outputs: None
'***************************************************************************

Dim sheetName As String, myRow As Long, hrType As Worksheet

'Check if row date is inside the pay period
If src.Cells(srcRow, 4) < startDate Or src.Cells(srcRow, 4) > endDate Then
    Exit Sub
End If

'Determine which sheet the data belongs in
If IsNumeric(src.Cells(srcRow, 3).Value) Then
    Set hrType = emps
Else
    Set hrType = cons
End If

'Find first empty row of appropriate sheet
If hrType.Range("A2").Value = "" Then
    myRow = 2
Else
    myRow = hrType.Range("A1").End(xlDown).Row + 1
End If

hrType.Range(hrType.Cells(myRow, 1), hrType.Cells(myRow, 4)).Value = _
    src.Range(src.Cells(srcRow, 1), src.Cells(srcRow, 4)).Value
hrType.Range(hrType.Cells(myRow, 6), hrType.Cells(myRow, 7)).Value = _
    src.Range(src.Cells(srcRow, 5), src.Cells(srcRow, 6)).Value
hrType.Range(hrType.Cells(myRow, 10), hrType.Cells(myRow, 11)).Value = _
    src.Range(src.Cells(srcRow, 7), src.Cells(srcRow, 8)).Value
hrType.Range(hrType.Cells(myRow, 12), hrType.Cells(myRow, 13)).Value = _
    src.Range(src.Cells(srcRow, 7), src.Cells(srcRow, 8)).Value
hrType.Range(hrType.Cells(myRow, 14), hrType.Cells(myRow, 17)).Value = _
    src.Range(src.Cells(srcRow, 9), src.Cells(srcRow, 12)).Value
    
'Calculate the "true" date and re-check if it's in range.

hrType.Range(hrType.Cells(myRow, 5).Address).Calculate

If hrType.Cells(myRow, 5) < startDate Or _
            hrType.Cells(myRow, 5) > endDate - 1 Then
    Call ClearValues(myRow, hrType)
End If
    
End Sub

Sub ClearValues(r, sheet):
'***************************************************************************
'Purpose: Clear constants in a given row without affecting formulas
'Inputs:  r - target row
'         sheet - sheet containing target row
'Outputs: None
'***************************************************************************

Dim rVals As Range

Set rVals = sheet.Range(sheet.Cells(r, 1), sheet.Cells(r, 18)).SpecialCells(xlCellTypeConstants)
rVals.ClearContents

End Sub

Sub checkDate()

'***************************************************************************
'Purpose: Ensure date given is valid and starts on a Saturday. Ends all
'         running macros if not.
'Inputs:  None
'Outputs: None
'***************************************************************************

If startDate = 0 Or Not IsDate(startDate) Then
    MsgBox "Please enter a valid date on the first sheet."
    End
ElseIf Weekday(startDate) <> 7 Then
    MsgBox "Pay period must start on a Saturday."
    End
End If

End Sub

Sub checkWorkers(conStart, empStart)

'***************************************************************************
'Purpose: After adding WOs, check through all workers to make sure they're
'         listed in the proper tables
'Inputs:  conStart - the first row with new data on the Contractor sheet
'         empStart - the first row with new data on the Employee sheet
'Outputs: None
'***************************************************************************

Dim rrow As Long, rng As Range

'Check Contractor sheet

For rrow = conStart To cons.ListObjects("Contractors").ListRows.Count + 1
    Set rng = dets.ListObjects("Con_Rates").ListColumns(3).Range.Find(cons.Cells(rrow, 3))
    If rng Is Nothing Then
        Call AddCon(cons.Cells(rrow, 3).Value, rrow)
    End If
Next rrow

For rrow = empStart To emps.ListObjects("Employees").ListRows.Count + 1
    Set rng = dets.ListObjects("Emp_Status").ListColumns(3).Range.Find(emps.Cells(rrow, 3))
    If rng Is Nothing Then
        Call AddEmp(emps.Cells(rrow, 3).Value, rrow)
    End If
Next rrow

End Sub

Sub checkWOs()

'***************************************************************************
'Purpose: After adding WOs, check to make sure there are no "orphan" WOs in
'         the list (i.e. WOs opened that had no data within range)
'Inputs:  None
'Outputs: None
'***************************************************************************

Dim i As Long, counter As Long, rng1 As Range, rng2 As Range
Dim woname As Long

Set cons = ThisWorkbook.Sheets("Contractors")
Set emps = ThisWorkbook.Sheets("Employees")
Set dets = ThisWorkbook.Sheets("Details")

i = 2
woname = dets.Cells(i, 1)

For counter = 1 To dets.ListObjects("WO").ListRows.Count + 1
    If dets.ListObjects("WO").ListRows.Count + 1 < i Then
        Exit For
    End If
    
    Set rng1 = cons.ListObjects("Contractors").ListColumns(6).Range.Find(dets.Cells(i, 1))
    Set rng2 = emps.ListObjects("Employees").ListColumns(6).Range.Find(dets.Cells(i, 1))
    If rng1 Is Nothing And rng2 Is Nothing Then
        dets.ListObjects("WO").ListRows(i - 1).Delete
    Else
        i = i + 1
    End If
Next

End Sub

Sub AddEmp(empID, rrow)

'***************************************************************************
'Purpose: Add an employee to the Emp_Status table on the Details sheet
'Inputs:  empID - the employee's ID number
'         rrow - the row in the Employee sheet where the employee's ID
'         was found
'Outputs: None
'***************************************************************************

Dim tbl As ListObject, masterTbl As ListObject
Dim newrowA As ListRow, newrowB As ListRow
Dim foundRange As Range, foundRow As Long, lastRow As Range, col As Integer

Set tbl = dets.ListObjects("Emp_Status")
Set masterTbl = master.Sheets("Employees").ListObjects("Employees")
Set foundRange = masterTbl.ListColumns(1).Range.Find(empID)

'If employee is not in the master list, add them
If foundRange Is Nothing Then
    Set newrowA = masterTbl.ListRows.Add
    With newrowA
        .Range(1).Value = empID
        .Range(2).Value = emps.Cells(rrow, 1).Value
        .Range(3).Value = emps.Cells(rrow, 2).Value
    End With
    foundRow = newrowA.Index
Else
    foundRow = foundRange.Row
End If

'Add employee to table based on data in master list

'Check if last row is empty and if not, add a row
If tbl.ListRows.Count > 0 Then
    Set lastRow = tbl.ListRows(tbl.ListRows.Count).Range
    For col = 1 To lastRow.Columns.Count
        If Trim(CStr(lastRow.Cells(1, col).Value)) <> "" Then
            tbl.ListRows.Add
            Exit For
        End If
    Next col
Else
    tbl.ListRows.Add
End If

'Populate last row with employee info
Set lastRow = tbl.ListRows(tbl.ListRows.Count).Range
With lastRow
    .Cells(1, 1).Value = masterTbl.Range(foundRow, 2).Value
    .Cells(1, 2).Value = masterTbl.Range(foundRow, 3).Value
    .Cells(1, 3).Value = empID
    .Cells(1, 4).Value = masterTbl.Range(foundRow, 4).Value
End With

End Sub

Sub AddCon(conID, rrow)

'***************************************************************************
'Purpose: Add a contractor to the Con_Rates table on the Details sheet
'Inputs:  conID - the contractor's ID
'         rrow - the row in the Contractor sheet where the contractor's ID
'                was found
'Outputs: None
'***************************************************************************

Dim tbl As ListObject, masterTbl As ListObject
Dim newrowA As ListRow, newrowB As ListRow
Dim foundRange As Range, foundRow As Long, lastRow As Range, col As Integer

Set tbl = dets.ListObjects("Con_Rates")
Set masterTbl = master.Sheets("Contractors").ListObjects("Contractors")
Set foundRange = masterTbl.ListColumns(1).Range.Find(conID)

'If contractor is not in the master list, add them
If foundRange Is Nothing Then
    Set newrowA = masterTbl.ListRows.Add
    With newrowA
        .Range(1).Value = conID
        .Range(2).Value = cons.Cells(rrow, 1).Value
        .Range(3).Value = cons.Cells(rrow, 2).Value
    End With
    foundRow = newrowA.Index
Else
    foundRow = foundRange.Row
End If

'Add contractor to table based on data in master list

'Check if last row is empty and if not, add a row
If tbl.ListRows.Count > 0 Then
    Set lastRow = tbl.ListRows(tbl.ListRows.Count).Range
    For col = 1 To lastRow.Columns.Count
        If Trim(CStr(lastRow.Cells(1, col).Value)) <> "" Then
            tbl.ListRows.Add
            Exit For
        End If
    Next col
Else
    tbl.ListRows.Add
End If

'Populate last row with contractor info
Set lastRow = tbl.ListRows(tbl.ListRows.Count).Range
With lastRow
    .Cells(1, 1).Value = masterTbl.Range(foundRow, 2).Value
    .Cells(1, 2).Value = masterTbl.Range(foundRow, 3).Value
    .Cells(1, 3).Value = conID
    .Cells(1, 4).Value = masterTbl.Range(foundRow, 4).Value
    .Cells(1, 5).Value = masterTbl.Range(foundRow, 5).Value
End With

End Sub

Sub AddWo(woID)

'***************************************************************************
'Purpose: Add a WO to the WO table on the Details sheet
'Inputs:  woID - the work order's numeric ID
'Outputs: None
'***************************************************************************

Dim tbl As ListObject, masterTbl As ListObject
Dim newrowA As ListRow, newrowB As ListRow
Dim foundRange As Range, foundRow As Long, lastRow As Range, col As Integer

Set tbl = dets.ListObjects("WO")
Set masterTbl = master.Sheets("WOs").ListObjects("WOs")
Set foundRange = masterTbl.ListColumns(1).Range.Find(woID)

'If WO is not in the master list, add it
If foundRange Is Nothing Then
    Set newrowA = masterTbl.ListRows.Add
    With newrowA
        .Range(1).Value = woID
        .Range(2).Value = InputBox("Please enter WO category (HMV/Line/Shop)", woID)
        .Range(3).Value = InputBox("Please enter WO aircraft if applicable (leave blank if not)", woID)
    End With
    foundRow = newrowA.Range.Row
Else
    foundRow = foundRange.Row
End If

'Add WO to table based on data in master list

'Check if last row is empty and if not, add a row
If tbl.ListRows.Count > 0 Then
    Set lastRow = tbl.ListRows(tbl.ListRows.Count).Range
    For col = 1 To lastRow.Columns.Count
        If Trim(CStr(lastRow.Cells(1, col).Value)) <> "" Then
            tbl.ListRows.Add
            Exit For
        End If
    Next col
Else
    tbl.ListRows.Add
End If

'Populate last row with WO info
Set lastRow = tbl.ListRows(tbl.ListRows.Count).Range
With lastRow
    .Cells(1, 1).Value = woID
    .Cells(1, 2).Value = masterTbl.Range(foundRow, 2).Value
    .Cells(1, 3).Value = masterTbl.Range(foundRow, 3).Value
End With

End Sub

Sub sortAll()

'Sort WO list
dets.ListObjects("WO").Sort.SortFields.Clear
dets.ListObjects("WO").Sort.SortFields.Add Key _
    :=dets.ListObjects("WO").ListColumns(1).Range, SortOn:=xlSortOnValues, Order:=xlAscending, _
    DataOption:=xlSortNormal
With dets.ListObjects("WO").Sort
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With

'Sort contractor list

dets.ListObjects("Con_Rates").Sort.SortFields.Clear
dets.ListObjects("Con_Rates").Sort.SortFields.Add Key _
    :=dets.ListObjects("Con_Rates").ListColumns(1).Range, SortOn:=xlSortOnValues, _
    Order:=xlAscending, DataOption:=xlSortNormal
dets.ListObjects("Con_Rates").Sort.SortFields.Add Key _
    :=dets.ListObjects("Con_Rates").ListColumns(2).Range, SortOn:=xlSortOnValues, _
    Order:=xlAscending, DataOption:=xlSortNormal
With dets.ListObjects("Con_Rates").Sort
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With

'Sort employee list
dets.ListObjects("Emp_Status").Sort.SortFields.Clear
dets.ListObjects("Emp_Status").Sort.SortFields.Add Key _
    :=dets.ListObjects("Emp_Status").ListColumns(1).Range, SortOn:=xlSortOnValues, _
    Order:=xlAscending, DataOption:=xlSortNormal
dets.ListObjects("Emp_Status").Sort.SortFields.Add Key _
    :=dets.ListObjects("Emp_Status").ListColumns(2).Range, SortOn:=xlSortOnValues, _
    Order:=xlAscending, DataOption:=xlSortNormal
With dets.ListObjects("Emp_Status").Sort
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With

'Sort master WO list
master.Sheets("WOs").ListObjects(1).Sort.SortFields.Clear
master.Sheets("WOs").ListObjects(1).Sort.SortFields.Add Key:= _
        master.Sheets("WOs").ListObjects(1).ListColumns(1).Range, SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
With master.Sheets("WOs").ListObjects(1).Sort
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With

'Sort master contractor list
master.Sheets("Contractors").ListObjects(1).Sort.SortFields.Clear
master.Sheets("Contractors").ListObjects(1).Sort.SortFields.Add Key _
    :=master.Sheets("Contractors").ListObjects(1).ListColumns(2).Range, SortOn:=xlSortOnValues, _
    Order:=xlAscending, DataOption:=xlSortNormal
master.Sheets("Contractors").ListObjects(1).Sort.SortFields.Add Key _
    :=master.Sheets("Contractors").ListObjects(1).ListColumns(3).Range, SortOn:=xlSortOnValues, _
    Order:=xlAscending, DataOption:=xlSortNormal
With master.Sheets("Contractors").ListObjects(1).Sort
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With

'Sort master employee list
master.Sheets("Employees").ListObjects(1).Sort.SortFields.Clear
master.Sheets("Employees").ListObjects(1).Sort.SortFields.Add Key _
    :=master.Sheets("Employees").ListObjects(1).ListColumns(2).Range, SortOn:=xlSortOnValues, _
    Order:=xlAscending, DataOption:=xlSortNormal
master.Sheets("Employees").ListObjects(1).Sort.SortFields.Add Key _
    :=master.Sheets("Employees").ListObjects(1).ListColumns(3).Range, SortOn:=xlSortOnValues, _
    Order:=xlAscending, DataOption:=xlSortNormal
With master.Sheets("Employees").ListObjects(1).Sort
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With

End Sub
