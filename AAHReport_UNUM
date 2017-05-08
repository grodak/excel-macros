Sub AAHReport_UNUM()
    
Dim wb As Workbook
Dim ws As Worksheet
Dim mDate As String

Set wb = ThisWorkbook
Set ws = ActiveSheet


'====================
'Freeze the first row
'====================
If Not ActiveWindow.FreezePanes Then
    ActiveWindow.FreezePanes = False
    ActiveWindow.SplitColumn = 0
    ActiveWindow.SplitRow = 1
    ActiveWindow.FreezePanes = True
End If

'===============================
'Delete unnecessary data columns
'===============================
ws.Range("I:I,M:M,Q:Q,U:U,Y:Y,AA:AA,AC:AC,AE:AE,AG:AG,AK:AK,AO:AO,AS:AS").Delete

'=================================================
'Create Covered_Monthly_Payroll column
'If the salary is over 200K, enter in 16667.67
'If less, divide by 12 and round to nearest dollar
'=================================================
ws.Columns(9).Insert xlToRight
ws.Cells(1, 9).Value = "Covered_Monthly_Payroll"
For i = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row
    If ws.Cells(i, 8).Value >= 200000 Then
        ws.Cells(i, 9).Value = 16667.67
        Else
            ws.Cells(i, 9).Value = Application.WorksheetFunction.Round(ws.Cells(i, 8).Value / 12, 0)
    End If
Next

'=====================================================================
'Sum the different benefit coverage amounts as well as premium amounts
'=====================================================================
lrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
ws.Cells(lrow + 1, 8).Value = Application.WorksheetFunction.SUM(ws.Range(ws.Cells(1, 8), ws.Cells(lrow, 8)))
ws.Range("H:I,K:L,N:O,Q:R,T:T,U:U,W:W,Y:Y,AA:AB,AD:AE,AG:AH").NumberFormat = "#,0.00"
ws.Cells(lrow + 1, 9).Value = Application.WorksheetFunction.SUM(ws.Range(ws.Cells(1, 9), ws.Cells(lrow, 9)))
'LIF
ws.Cells(lrow + 1, 11).Value = Application.WorksheetFunction.SUM(ws.Range(ws.Cells(1, 11), ws.Cells(lrow, 11)))
ws.Cells(lrow + 1, 12).Value = Application.WorksheetFunction.SUM(ws.Range(ws.Cells(1, 12), ws.Cells(lrow, 12)))
'SUP
ws.Cells(lrow + 1, 14).Value = Application.WorksheetFunction.SUM(ws.Range(ws.Cells(1, 14), ws.Cells(lrow, 14)))
ws.Cells(lrow + 1, 15).Value = Application.WorksheetFunction.SUM(ws.Range(ws.Cells(1, 15), ws.Cells(lrow, 15)))
'SSP
ws.Cells(lrow + 1, 17).Value = Application.WorksheetFunction.SUM(ws.Range(ws.Cells(1, 17), ws.Cells(lrow, 17)))
ws.Cells(lrow + 1, 18).Value = Application.WorksheetFunction.SUM(ws.Range(ws.Cells(1, 18), ws.Cells(lrow, 18)))
'SDP
ws.Cells(lrow + 1, 20).Value = Application.WorksheetFunction.SUM(ws.Range(ws.Cells(1, 20), ws.Cells(lrow, 20)))
ws.Cells(lrow + 1, 21).Value = Application.WorksheetFunction.SUM(ws.Range(ws.Cells(1, 21), ws.Cells(lrow, 21)))
'STD
ws.Cells(lrow + 1, 23).Value = Application.WorksheetFunction.SUM(ws.Range(ws.Cells(1, 23), ws.Cells(lrow, 23)))
'LTD
ws.Cells(lrow + 1, 25).Value = Application.WorksheetFunction.SUM(ws.Range(ws.Cells(1, 25), ws.Cells(lrow, 25)))
'SAD
ws.Cells(lrow + 1, 27).Value = Application.WorksheetFunction.SUM(ws.Range(ws.Cells(1, 27), ws.Cells(lrow, 27)))
ws.Cells(lrow + 1, 28).Value = Application.WorksheetFunction.SUM(ws.Range(ws.Cells(1, 28), ws.Cells(lrow, 28)))
'SM2
ws.Cells(lrow + 1, 30).Value = Application.WorksheetFunction.SUM(ws.Range(ws.Cells(1, 30), ws.Cells(lrow, 30)))
ws.Cells(lrow + 1, 31).Value = Application.WorksheetFunction.SUM(ws.Range(ws.Cells(1, 31), ws.Cells(lrow, 31)))
'DM2
ws.Cells(lrow + 1, 33).Value = Application.WorksheetFunction.SUM(ws.Range(ws.Cells(1, 33), ws.Cells(lrow, 33)))
ws.Cells(lrow + 1, 34).Value = Application.WorksheetFunction.SUM(ws.Range(ws.Cells(1, 34), ws.Cells(lrow, 34)))

'============================================================
'Count the number of records that have the different benefits
'============================================================
'LIF
ws.Cells(lrow + 1, 10).Value = "Count: " & ws.Range(ws.Cells(2, 10), ws.Cells(lrow, 10)).Cells.SpecialCells(xlCellTypeConstants).Count
'SUP
ws.Cells(lrow + 1, 13).Value = "Count: " & ws.Range(ws.Cells(2, 13), ws.Cells(lrow, 13)).Cells.SpecialCells(xlCellTypeConstants).Count
'SSP
ws.Cells(lrow + 1, 16).Value = "Count: " & ws.Range(ws.Cells(2, 16), ws.Cells(lrow, 16)).Cells.SpecialCells(xlCellTypeConstants).Count
'SDP
ws.Cells(lrow + 1, 19).Value = "Count: " & ws.Range(ws.Cells(2, 19), ws.Cells(lrow, 19)).Cells.SpecialCells(xlCellTypeConstants).Count
'STD
ws.Cells(lrow + 1, 22).Value = "Count: " & ws.Range(ws.Cells(2, 22), ws.Cells(lrow, 22)).Cells.SpecialCells(xlCellTypeConstants).Count
'LTD
ws.Cells(lrow + 1, 24).Value = "Count: " & ws.Range(ws.Cells(2, 24), ws.Cells(lrow, 24)).Cells.SpecialCells(xlCellTypeConstants).Count
'SAD
ws.Cells(lrow + 1, 26).Value = "Count: " & ws.Range(ws.Cells(2, 26), ws.Cells(lrow, 26)).Cells.SpecialCells(xlCellTypeConstants).Count
'SM2
ws.Cells(lrow + 1, 29).Value = "Count: " & ws.Range(ws.Cells(2, 29), ws.Cells(lrow, 29)).Cells.SpecialCells(xlCellTypeConstants).Count
'DM2
ws.Cells(lrow + 1, 32).Value = "Count: " & ws.Range(ws.Cells(2, 32), ws.Cells(lrow, 32)).Cells.SpecialCells(xlCellTypeConstants).Count

'=========================================
'Bold the totals row and autofit the cells
'=========================================
ws.Rows(lrow + 1).Font.Bold = True
ws.UsedRange.Rows.AutoFit
ws.UsedRange.Columns.AutoFit

End Sub
