Attribute VB_Name = "Export_tables_to_excel"
Option Explicit

Sub Export_to_excel()
Dim tableNum As Integer 'number of tables
Dim Ex As New Excel.Application
Dim WB As Excel.Workbook
Dim Sh As Excel.Worksheet
Dim i As Integer 'row counter
Dim j As Integer 'coloumn counter
Dim t As Integer 'tables counter
Dim resultrow As Integer
Ex.Visible = True
Dim NewFileName As String

With Application
.StatusBar = "Wait"
.ScreenUpdating = False
.DisplayAlerts = False
End With

With ActiveDocument 'if there are no tables in document just exit woth message box
    tableNum = .Tables.Count
    If tableNum = 0 Then
        MsgBox "There are no tables in document. Nothing to export"
        GoTo Handle
    End If

    NewFileName = ActiveDocument.Path & "\ " & Replace(ActiveDocument.Name, ".docx", "") & " tables" & ".xlsx" 'name for excel file
    resultrow = 1 'number of start row in excel
End With

Set WB = Ex.Workbooks.Add

On Error Resume Next

For t = 1 To tableNum 'for all tables in document
    With ActiveDocument.Tables(t)
        For i = 1 To .Rows.Count
            For j = 1 To .Columns.Count
                WB.ActiveSheet.Cells(resultrow, j) = WorksheetFunction.Clean(.Cell(i, j).Range.Text)
            Next j
            resultrow = resultrow + 1
        Next i
    End With
    resultrow = 1
    Set Sh = WB.Sheets.Add(After:=WB.Worksheets(WB.Worksheets.Count)) 'add sheet for next table
Next t
With WB
    .Sheets(WB.Worksheets.Count).Delete 'delete last sheet cause it is empty
    .SaveAs FileName:=NewFileName, _
    FileFormat:=51, _
    AccessMode:=xlExclusive, _
    ConflictResolution:=xlLocalSessionChanges
    .Close
End With
Ex.Quit

Handle:
With Application
    .StatusBar = "Ready"
    .ScreenUpdating = True
    .DisplayAlerts = True
End With
End Sub
