Option Explicit
Public empSrc As Workbook


Sub main():

Dim wb As Workbook, src As Worksheet, srcRow As Double, myRow As Double
Dim emp As String, wDate As Date, i As Double

Application.ScreenUpdating = False
Application.EnableEvents = False
Application.Calculation = xlCalculationManual

Set wb = Workbooks.Open(GetFile())
Set src = wb.Worksheets(1)
Set empSrc = Workbooks.Open(Application.ThisWorkbook.Path & "\MASTER LISTS.xlsm")
srcRow = 6
myRow = ThisWorkbook.Sheets("Hours").Range("A1").End(xlDown).Row + 1

Do While src.Cells(srcRow, 1).Value <> "Summary"
    emp = src.Cells(srcRow, 1).Value
    wDate = src.Cells(srcRow, 6).Value
    i = 1
        Do Until src.Cells(srcRow + i, 6).Value <> wDate
            i = i + 1
        Loop
    Call WriteDate(src, srcRow, i - 1, myRow)
    myRow = myRow + 1
    srcRow = srcRow + i
Loop
    
wb.Close False
empSrc.Close False

Application.ScreenUpdating = True
Application.EnableEvents = True
Application.Calculation = xlCalculationAutomatic


End Sub

Function GetFile()

'***************************************************************************
'Purpose: Prompt user to select file for data importation
'Inputs:  None
'Outputs: None
'***************************************************************************

Dim fpath As String

fpath = Application.GetOpenFilename(FileFilter:="Excel Files (*.xlsx), *.xlsx", _
        Title:="Select Source File")
        
GetFile = fpath

End Function

Sub WriteDate(src, i, j, X)

'***************************************************************************
'Purpose: Take data from source workbook and write it into the Hours sheet
'Inputs:  src - The source workbook
'         i - The first row for a given employee/date
'         j - The offset from i to the last row for a given employee/date
'         x - The current row in the Hours sheet
'Outputs: None
'***************************************************************************

Dim id As Long, idRef As Range, hrs As Worksheet

Set hrs = ThisWorkbook.Sheets("Hours")

hrs.Cells(X, 3).Value = CLng(Left(src.Cells(i, 1), 6))
hrs.Cells(X, 6).Value = src.Cells(i, 6)
hrs.Cells(X, 7).Value = Application.WorksheetFunction.Min(Range(src.Cells(i, 9), _
            src.Cells(i + j, 9)))
hrs.Cells(X, 8).Value = Application.WorksheetFunction.Max(Range(src.Cells(i, 10), _
            src.Cells(i + j, 10)))
hrs.Cells(X, 10).Value = Application.WorksheetFunction.Sum(Range(src.Cells(i, 8), _
            src.Cells(i + j, 8))) / 24
            
id = hrs.Cells(X, 3).Value
Set idRef = hrs.Range("C" & X)

Call GetEmp(id, idRef)

End Sub

Sub GetEmp(id, idRef)

'***************************************************************************
'Purpose: Given an employee number, retrieve their name in TRAX format and
'         write it into the appropriate cells.
'Inputs:  id - the employee ID
'         idref - the cell containing the employee ID
'Outputs: Employee's first and last names, as formatted in TRAX reports
'***************************************************************************

Dim idSrc As Range

With empSrc.Worksheets("Employees").Range("A:A")
    Set idSrc = .Find(What:=id)
    If Not idSrc Is Nothing Then
        idRef.Offset(0, -2).Value = idSrc.Offset(0, 1).Value
        idRef.Offset(0, -1).Value = idSrc.Offset(0, 2).Value
    End If
End With

End Sub
