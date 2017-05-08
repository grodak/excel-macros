Attribute VB_Name = "Module15"
Sub RHTPensionDeductionFile()

Dim ws As Worksheet
Dim lr As Long, lc As Long
Dim time1 As Double, time2 As Double
Dim pvtRng As Range
Dim pvtCache As PivotCache
Dim pvtTbl As PivotTable
Dim nextMonth As String

Application.ScreenUpdating = False

'===========
'start timer
'===========
time1 = Timer

'=======================
'adding sheets for later
'=======================
ActiveWorkbook.Sheets.Add.Name = "PivotTable"
ActiveWorkbook.Sheets.Add.Name = "Data"
ActiveWorkbook.Sheets.Add.Name = "RHTPayrDedFile"
Sheets("Sheet3").Name = "COPSTrust"

'==========================
'set worksheets to variable
'==========================
Set ws = Sheets("RHTPayrDedFile")
Set ws2 = Sheets("PivotTable")
Set ws3 = Sheets("Data")
Set wsCOPS = Sheets("Sheet2")
Set ws4 = Sheets("COPSTrust")

'================================
'grabs RHTPayrDedFile From Access
'================================
Sheets("Sheet1").Activate
With Sheets("Sheet1").ListObjects.Add(SourceType:=0, Source:=Array( _
    "OLEDB;Provider=Microsoft.ACE.OLEDB.12.0;Password="""";User ID=Admin;Data Source=Z:\XXX\XXX\XXX.mdb;Mode=ReadWrite;Extended Properties="""";Jet OLE" _
    , _
    "DB:System database="""";Jet OLEDB:Registry Path="""";Jet OLEDB:Database Password="""";Jet OLEDB:Engine Type=5;Jet OLEDB:Database Locki" _
    , _
    "ng Mode=1;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="""";Jet OLEDB:" _
    , _
    "Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Wit" _
    , _
    "hout Replica Repair=False;Jet OLEDB:SFP=False;Jet OLEDB:Support Complex Data=False;Jet OLEDB:Bypass UserInfo Validation=False" _
    ), Destination:=Range("$A$1")).QueryTable
    .CommandType = xlCmdTable
    .CommandText = Array("RHTPayrDedFile")
    .RowNumbers = False
    .FillAdjacentFormulas = False
    .PreserveFormatting = True
    .RefreshOnFileOpen = False
    .BackgroundQuery = True
    .RefreshStyle = xlInsertDeleteCells
    .SavePassword = False
    .SaveData = True
    .AdjustColumnWidth = True
    .RefreshPeriod = 0
    .PreserveColumnInfo = True
    .SourceDataFile = "Z:\XXX\XXX\XXX.mdb"
    .ListObject.DisplayName = "Table_RHTPenDeductions_"
    .Refresh BackgroundQuery:=False
End With

'============================
'Copys Access Table as Values
'============================
Range("Table_RHTPenDeductions_[[#Headers],[PEFROMDT]]").Select
Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
Selection.Copy
ws.Select
Selection.PasteSpecial Paste:=xlPasteValues

ws.Activate
ws.Columns("R:R").Insert Shift:=xlToRight

'===========================================
'finds how many coverages a Participant has.
'===========================================
lr = ws.Cells(Rows.Count, 1).End(xlUp).Row
For i = 2 To lr
    If Len(ws.Cells(i, 6).Value) > 1 Then
        If Len(ws.Cells(i, 7).Value) = 0 Then
            ws.Cells(i, 18).Value = "1"
            Else
            If Len(ws.Cells(i, 8).Value) = 0 Then
                ws.Cells(i, 18).Value = "2"
                Else
                If Len(ws.Cells(i, 9).Value) = 0 Then
                    ws.Cells(i, 18).Value = "3"
                    Else
                        ws.Cells(i, 18).Value = "4"
    End If
    End If
    End If
    End If
                
Next i

'==============================
'Moves coverages to one per row
'==============================
For i = lr To 2 Step -1
If ws.Cells(i, 18).Value > 1 Then
    If ws.Cells(i, 18).Value = 2 Then
        Rows(i).Offset(1).Insert
        ws.Range(Cells(i, 1), Cells(i, 5)).Copy ws.Range("A" & i + 1)
        ws.Cells(i, 7).Cut ws.Range("F" & i + 1)
        ws.Cells(i, 11).Cut ws.Range("J" & i + 1)
        ws.Cells(i, 15).Cut ws.Range("N" & i + 1)
    End If
    If ws.Cells(i, 18).Value = 3 Then
        Rows(i).Offset(1).Insert
        Rows(i).Offset(1).Insert
        ws.Range(Cells(i, 1), Cells(i, 5)).Copy ws.Range("A" & i + 1)
        ws.Range(Cells(i, 1), Cells(i, 5)).Copy ws.Range("A" & i + 2)
        ws.Cells(i, 7).Cut ws.Range("F" & i + 1)
        ws.Cells(i, 11).Cut ws.Range("J" & i + 1)
        ws.Cells(i, 15).Cut ws.Range("N" & i + 1)
        ws.Cells(i, 8).Cut ws.Range("F" & i + 2)
        ws.Cells(i, 12).Cut ws.Range("J" & i + 2)
        ws.Cells(i, 16).Cut ws.Range("N" & i + 2)
    End If
    If ws.Cells(i, 18).Value = 4 Then
        Rows(i).Offset(1).Insert
        Rows(i).Offset(1).Insert
        Rows(i).Offset(1).Insert
        ws.Range(Cells(i, 1), Cells(i, 5)).Copy ws.Range("A" & i + 1)
        ws.Range(Cells(i, 1), Cells(i, 5)).Copy ws.Range("A" & i + 2)
        ws.Range(Cells(i, 1), Cells(i, 5)).Copy ws.Range("A" & i + 3)
        ws.Cells(i, 7).Cut ws.Range("F" & i + 1)
        ws.Cells(i, 11).Cut ws.Range("J" & i + 1)
        ws.Cells(i, 15).Cut ws.Range("N" & i + 1)
        ws.Cells(i, 8).Cut ws.Range("F" & i + 2)
        ws.Cells(i, 12).Cut ws.Range("J" & i + 2)
        ws.Cells(i, 16).Cut ws.Range("N" & i + 2)
        ws.Cells(i, 9).Cut ws.Range("F" & i + 3)
        ws.Cells(i, 13).Cut ws.Range("J" & i + 3)
        ws.Cells(i, 17).Cut ws.Range("N" & i + 3)
    End If
End If
Next i

'=============================
'gets rid of un-needed columns
'=============================
ws.Range("G:I,K:M,O:R").Select
Selection.Delete Shift:=xlToLeft

'===================
'formatting for File
'===================
ws.Cells(1, 10).Value = "Deduction Amount"
ws.Cells(1, 11).Value = "Fringe Amount"

'=========================
'deletes HRA coverage rows
'=========================
lr = ws.Cells(Rows.Count, 1).End(xlUp).Row
For i = lr To 2 Step -1
If ws.Cells(i, 6).Value = "HRA" Then
    Rows(i).EntireRow.Delete
End If
Next i

'=============================
'deletes MED Opt Out coverages
'=============================
lr = ws.Cells(Rows.Count, 1).End(xlUp).Row
For i = lr To 2 Step -1
If ws.Cells(i, 7).Value = "RHT-MED-ADV-OPTOUT" Then
    Rows(i).EntireRow.Delete
End If
Next i

'==========================================================================
'deletes people with no coverage. SQL statement should make this not needed
'==========================================================================
lr = ws.Cells(Rows.Count, 1).End(xlUp).Row
For i = lr To 2 Step -1
If ws.Cells(i, 6).Value = "" Then
    Rows(i).EntireRow.Delete
End If
Next i

'=======================================
'deletes people without valid pension ID
'=======================================
lr = ws.Cells(Rows.Count, 1).End(xlUp).Row
For i = lr To 2 Step -1
If Len(ws.Cells(i, 2).Value) <> 10 Then
    Rows(i).EntireRow.Delete
End If
Next i

'==============================================
'adds deduction amount based on Plan & Cov Tier
'==============================================
lr = ws.Cells(Rows.Count, 1).End(xlUp).Row
For i = 2 To lr
Select Case ws.Cells(i, 7).Value
    Case "AETNA-MAPD-PPO"
        ws.Cells(i, 10).Value = "23.76"
        ws.Cells(i, 11).Value = "90.00"
    Case "HAP-MAPD-HMO"
        ws.Cells(i, 10).Value = "40.55"
        ws.Cells(i, 11).Value = "90.00"
    Case "COPS-DELTA-DEN-HIGH"
        Select Case ws.Cells(i, 8).Value
            Case "P00"
                ws.Cells(i, 10).Value = "35.77"
                ws.Cells(i, 11).Value = "0.00"
            Case "F00"
                ws.Cells(i, 10).Value = "70.82"
                ws.Cells(i, 11).Value = "0.00"
            Case "P01"
                ws.Cells(i, 10).Value = "70.82"
                ws.Cells(i, 11).Value = "0.00"
            Case "F99"
                ws.Cells(i, 10).Value = "119.07"
                ws.Cells(i, 11).Value = "0.00"
            Case Else
                ws.Cells(i, 9).Value = "NO COV MAT IN COPS-DELTA-DEN-HIGH"
        End Select
    Case "COPS-DELTA-DEN-LOW"
        Select Case ws.Cells(i, 8).Value
            Case "P00"
                ws.Cells(i, 10).Value = "29.76"
                ws.Cells(i, 11).Value = "0.00"
            Case "F00"
                ws.Cells(i, 10).Value = "56.04"
                ws.Cells(i, 11).Value = "0.00"
            Case "P01"
                ws.Cells(i, 10).Value = "56.04"
                ws.Cells(i, 11).Value = "0.00"
            Case "F99"
                ws.Cells(i, 10).Value = "96.76"
                ws.Cells(i, 11).Value = "0.00"
            Case Else
                ws.Cells(i, 9).Value = "NO COV MATCH IN COPS-DELTA-DEN-LOW"
        End Select
    Case "VIS-12"
        ws.Cells(i, 10).Value = "13.50"
        ws.Cells(i, 11).Value = "0.00"
    Case Else
        ws.Cells(i, 9).Value = "NO PLAN MATCH!"
    End Select
Next i



'==========================================================
'creates pivot table to combine amounts based on Pension ID
'==========================================================
lr = ws.Cells(Rows.Count, 1).End(xlUp).Row
lc = ws.Cells(1, Columns.Count).End(xlToLeft).Column
Set pvtRng = ws.Cells(1, 1).Resize(lr, lc)
Set pvtCache = ActiveWorkbook.PivotCaches.Add(SourceType:=xlDatabase, SourceData:=pvtRng)

Set pvtTbl = pvtCache.CreatePivotTable(TableDestination:=ws2.Cells(1, 1), TableName:="PivotTable1")

pvtTbl.ManualUpdate = True

pvtTbl.AddFields RowFields:=Array("G7USEDF7", "COV1", "PLAN1"), ColumnFields:="Data"

With pvtTbl.PivotFields("Deduction Amount")
    .Orientation = xlDataField
    .Function = xlSum
    .Position = 1
    .NumberFormat = "###.#0"
    .Name = "Deduction Amount "
End With

With pvtTbl.PivotFields("Fringe Amount")
    .Orientation = xlDataField
    .Function = xlSum
    .Position = 2
    .NumberFormat = "###.00"
    .Name = "Fringe Amount "
End With

pvtTbl.ManualUpdate = False

With pvtTbl
    .ColumnGrand = False
    .RowGrand = False
End With

pvtTbl.PivotFields("G7USEDF7").Subtotals(1) = False
pvtTbl.PivotFields("COV1").Subtotals(1) = False

pvtTbl.TableRange2.Offset(1, 0).Copy
ws3.Cells(1, 1).PasteSpecial xlPasteValues



'===================================
'Fills in blank cells with ID number
'===================================
ws3.Activate
ws3.Columns("A:A").SpecialCells(xlCellTypeBlanks).Select
Selection.FormulaR1C1 = "=R[-1]C"
ws3.Columns("A:A").Copy
ws3.Cells(1, 1).PasteSpecial xlPasteValues

'==================================
'adds column headers for final file
'==================================
ws3.Cells(1, 1).Value = "MemberID"
ws3.Cells(1, 2).Value = "Type of Change"
ws3.Cells(1, 3).Value = "Benefit Code"
ws3.Cells(1, 4).Value = "Benefit Deduction Amount"
ws3.Cells(1, 5).Value = "Benefit Deduction Adjustment Amount"
ws3.Cells(1, 6).Value = "Benefit Fringe (City) Amount"
ws3.Cells(1, 7).Value = "Benefit Fringe Adjustment Amount"
ws3.Cells(1, 8).Value = "Effective Date"
ws3.Cells(1, 9).Value = "Origination"
ws3.Cells(1, 10).Value = "Member Name"

'====================================
'adds Deduction Code based on amount.
'====================================
lr = ws3.Cells(Rows.Count, 1).End(xlUp).Row
For i = 2 To lr
    Select Case ws3.Cells(i, 4).Value
        Case "23.76"
            ws3.Cells(i, 3).Value = "GGSQA120"
        Case "47.52"
            ws3.Cells(i, 3).Value = "GGSQA220"
        Case "71.28"
            ws3.Cells(i, 3).Value = "GGSQA320"
        Case "40.55"
            ws3.Cells(i, 3).Value = "EFSQA120"
        Case "81.1"
            ws3.Cells(i, 3).Value = "EFSQA220"
        Case "121.65"
            ws3.Cells(i, 3).Value = "EFSQA320"
        Case "13.5"
            ws3.Cells(i, 3).Value = "'00040010"
        Case "35.77"
            ws3.Cells(i, 3).Value = "'00050030"
        Case "70.82"
            ws3.Cells(i, 3).Value = "'00050031"
        Case "119.07"
            ws3.Cells(i, 3).Value = "'00050032"
        Case "29.76"
            ws3.Cells(i, 3).Value = "'00050035"
        Case "56.04"
            ws3.Cells(i, 3).Value = "'00050036"
        Case "96.76"
            ws3.Cells(i, 3).Value = "'00050037"
        Case Else
            ws3.Cells(i, 3).Value = "AMOUNT NOT FOUND"
    End Select
Next i

'=======================================
'Fills in Origination column with value.
'=======================================
ws3.Range("I2:I" & lr).Value = "PFVEBA"
ws3.Columns("D:G").NumberFormat = "0.00"

'======================
'deletes spouse records
'======================
lr = ws.Cells(Rows.Count, 1).End(xlUp).Row
For i = lr To 2 Step -1
    If ws.Cells(i, 9).Value = "S" Or ws.Cells(i, 9).Value = "D" Then
        ws.Cells(i, 2).Value = ""
    End If
Next i

'=============
'combines name
'=============
lr = ws.Cells(Rows.Count, 1).End(xlUp).Row
For i = 2 To lr
    ws.Cells(i, 3).Value = Trim(ws.Cells(i, 3).Value) & ", " & Trim(ws.Cells(i, 4).Value) & " " & Trim(ws.Cells(i, 5).Value)
Next i

'================================
'adds Retirees Name to final file
'================================
lr = ws.Cells(Rows.Count, 1).End(xlUp).Row
lr2 = ws3.Cells(Rows.Count, 1).End(xlUp).Row
On Error Resume Next
For i = 2 To lr2
    ws3.Cells(i, 10).Value = Application.WorksheetFunction.Index(ws.Range("C:C"), Application.WorksheetFunction.Match(ws3.Cells(i, 1).Value, ws.Range("B:B"), 0))
Next i

ws3.Rows(ws3.Cells(Rows.Count, 1).End(xlUp).Row).EntireRow.Delete
ws3.Range("E2:G" & lr2).ClearContents

'==========================================
'grabs COPSPensionDeductionFile From Access
'==========================================
wsCOPS.Activate
With wsCOPS.ListObjects.Add(SourceType:=0, Source:=Array( _
        "OLEDB;Provider=Microsoft.ACE.OLEDB.12.0;Password="""";User ID=Admin;Data Source=Z:\XXX\XXX\XXX.mdb;Mode=ReadWrite;Extended Properties="""";Jet OLE" _
        , _
        "DB:System database="""";Jet OLEDB:Registry Path="""";Jet OLEDB:Database Password="""";Jet OLEDB:Engine Type=5;Jet OLEDB:Database Locki" _
        , _
        "ng Mode=1;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="""";Jet OLEDB:" _
        , _
        "Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Wit" _
        , _
        "hout Replica Repair=False;Jet OLEDB:SFP=False;Jet OLEDB:Support Complex Data=False;Jet OLEDB:Bypass UserInfo Validation=False" _
        ), Destination:=Range("$A$1")).QueryTable
        .CommandType = xlCmdTable
        .CommandText = Array("COPSPensionDeductionFile")
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .SourceDataFile = "Z:\XXX\XXX\XXX.mdb"
        .ListObject.DisplayName = "Table_RHTPenDeductions_3"
        .Refresh BackgroundQuery:=False
End With

Range("Table_RHTPenDeductions_3[[#Headers],[PEFROMDT]]").Select
Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
Selection.Copy
ws4.Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

'===========================================
'finds how many coverages a Participant has.
'===========================================
lr = ws4.Cells(Rows.Count, 1).End(xlUp).Row
For i = 2 To lr
    If Len(ws4.Cells(i, 6).Value) > 1 Then
        If Len(ws4.Cells(i, 7).Value) = 0 Then
            ws4.Cells(i, 19).Value = "1"
            Else
            If Len(ws4.Cells(i, 8).Value) = 0 Then
                ws4.Cells(i, 19).Value = "2"
                Else
                If Len(ws4.Cells(i, 9).Value) = 0 Then
                    ws4.Cells(i, 19).Value = "3"
                    Else
                        ws4.Cells(i, 19).Value = "4"
    End If
    End If
    End If
    End If
                
Next i

'==============================
'Moves coverages to one per row
'==============================
lr = ws4.UsedRange.Rows.Count
For i = lr To 2 Step -1
If ws4.Cells(i, 19).Value > 1 Then
    If ws4.Cells(i, 19).Value = 2 Then
        Rows(i).Offset(1).Insert
        ws4.Range(Cells(i, 1), Cells(i, 5)).Copy ws4.Range("A" & i + 1)
        ws4.Cells(i, 7).Cut ws4.Range("F" & i + 1)
        ws4.Cells(i, 11).Cut ws4.Range("J" & i + 1)
        ws4.Cells(i, 15).Cut ws4.Range("N" & i + 1)
    End If
    If ws4.Cells(i, 19).Value = 3 Then
        Rows(i).Offset(1).Insert
        Rows(i).Offset(1).Insert
        ws4.Range(Cells(i, 1), Cells(i, 5)).Copy ws4.Range("A" & i + 1)
        ws4.Range(Cells(i, 1), Cells(i, 5)).Copy ws4.Range("A" & i + 2)
        ws4.Cells(i, 7).Cut ws4.Range("F" & i + 1)
        ws4.Cells(i, 11).Cut ws4.Range("J" & i + 1)
        ws4.Cells(i, 15).Cut ws4.Range("N" & i + 1)
        ws4.Cells(i, 8).Cut ws4.Range("F" & i + 2)
        ws4.Cells(i, 12).Cut ws4.Range("J" & i + 2)
        ws4.Cells(i, 16).Cut ws4.Range("N" & i + 2)
    End If
    If ws4.Cells(i, 19).Value = 4 Then
        Rows(i).Offset(1).Insert
        Rows(i).Offset(1).Insert
        Rows(i).Offset(1).Insert
        ws4.Range(Cells(i, 1), Cells(i, 5)).Copy ws4.Range("A" & i + 1)
        ws4.Range(Cells(i, 1), Cells(i, 5)).Copy ws4.Range("A" & i + 2)
        ws4.Range(Cells(i, 1), Cells(i, 5)).Copy ws4.Range("A" & i + 3)
        ws4.Cells(i, 7).Cut ws4.Range("F" & i + 1)
        ws4.Cells(i, 11).Cut ws4.Range("J" & i + 1)
        ws4.Cells(i, 15).Cut ws4.Range("N" & i + 1)
        ws4.Cells(i, 8).Cut ws4.Range("F" & i + 2)
        ws4.Cells(i, 12).Cut ws4.Range("J" & i + 2)
        ws4.Cells(i, 16).Cut ws4.Range("N" & i + 2)
        ws4.Cells(i, 9).Cut ws4.Range("F" & i + 3)
        ws4.Cells(i, 13).Cut ws4.Range("J" & i + 3)
        ws4.Cells(i, 17).Cut ws4.Range("N" & i + 3)
    End If
End If
Next i

ws4.Range("G:I,K:M,O:S").Delete

'=========================
'deletes HRA coverage rows
'=========================
lr = ws4.Cells(Rows.Count, 1).End(xlUp).Row
For i = lr To 2 Step -1
If ws4.Cells(i, 6).Value = "HRA" Then
    Rows(i).EntireRow.Delete
End If
Next i

'==============================================
'adds deduction amount based on Plan & Cov Tier
'==============================================
lr = ws4.Cells(Rows.Count, 1).End(xlUp).Row
For i = 2 To lr
Select Case ws4.Cells(i, 7).Value
    Case "COPS-DELTA-DEN-HIGH"
        Select Case ws4.Cells(i, 8).Value
            Case "P00"
                ws4.Cells(i, 9).Value = "'00050030"
                ws4.Cells(i, 10).Value = "35.77"
            Case "F00"
                ws4.Cells(i, 9).Value = "'00050031"
                ws4.Cells(i, 10).Value = "70.82"
            Case "P01"
                ws4.Cells(i, 9).Value = "'00050031"
                ws4.Cells(i, 10).Value = "70.82"
            Case "F99"
                ws4.Cells(i, 9).Value = "'00050032"
                ws4.Cells(i, 10).Value = "119.07"
            Case Else
                ws4.Cells(i, 9).Value = "NO COV MATCH IN COPS-DELTA-DEN-HIGH"
        End Select
    Case "COPS-DELTA-DEN-LOW"
        Select Case ws4.Cells(i, 8).Value
            Case "P00"
                ws4.Cells(i, 9).Value = "'00050035"
                ws4.Cells(i, 10).Value = "29.76"
            Case "F00"
                ws4.Cells(i, 9).Value = "'00050036"
                ws4.Cells(i, 10).Value = "56.04"
            Case "P01"
                ws4.Cells(i, 9).Value = "'00050036"
                ws4.Cells(i, 10).Value = "56.04"
            Case "F99"
                ws4.Cells(i, 9).Value = "'00050037"
                ws4.Cells(i, 10).Value = "96.76"
            Case Else
                ws4.Cells(i, 9).Value = "NO COV MATCH IN COPS-DELTA-DEN-LOW"
        End Select
    Case "VIS-12"
        ws4.Cells(i, 9).Value = "'00040010"
        ws4.Cells(i, 10).Value = "13.50"
    Case Else
        ws4.Cells(i, 9).Value = "NO PLAN MATCH!"
    End Select
Next i

'=============
'Rename sheets
'=============
Sheets("Sheet1").Name = "Raw RHT Data"
Sheets("Sheet2").Name = "Raw COPS Data"

Application.ScreenUpdating = True

time2 = Timer

MsgBox "Done in " & Format((time2 - time1) / 60, "0.00 \min"), vbInformation

End Sub








