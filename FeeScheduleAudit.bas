
Sub FeeScheduleAudit()

Dim sh1 As Worksheet, sh2 As Worksheet
Dim time1 As Double, time2 As Double
Dim sRange As Range, oCell As Range

time1 = Timer

'Asks User for the 4 5-digit codes that reference the Fee Schedule columns.
oneCode = InputBox("What is the 5 digit code for the SmartHealth Par Non-Facility Amount?")
twoCode = InputBox("What is the 5 digit code for the SmartHealth Par Facility Amount?")
threeCode = InputBox("What is the 5 digit code for the SmartHealth Par Non-Facility Amount pd @ 85%?")
fourCode = InputBox("What is the 5 digit code for the SmartHealth Par Facility Amount pd @ 85%?")

wSheet1 = InputBox("What is the name of the tab where you pasted the Fee Schedule into?")
wSheet2 = InputBox("What is the tab name of the Allowable Charge Detail spreadsheet?")

Set sh1 = Sheets(wSheet1)
Set sh2 = Sheets(wSheet2)
Set sRange = sh2.UsedRange
Set oCell = sh2.Cells(1, 1)


Application.ScreenUpdating = False 'Speeds up code

sh1.Activate
sh1.UsedRange.Select
Selection.ClearFormats 'Gets rid of the formatting on the Fee Schedules

sh1.Cells(1, 5).Value = "SmartHealth Par Non-Facility Amount (If Different)"
sh1.Cells(1, 7).Value = "SmartHealth Par Facility Amount (If Different)"
sh1.Cells(1, 9).Value = "SmartHealth Par Non-Facility Amount pd @ 85% (If Different)"
sh1.Cells(1, 11).Value = "SmartHealth Par Facility Amount pd @ 85% (If Different)"

lr2 = sh2.Cells(Rows.Count, 1).End(xlUp).Row

sh2.Activate
sRange.Sort Key1:=oCell, Order1:=xlAscending, Header:=xlYes

sh2.Columns(1).TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlNone, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True

sh2.Columns(4).TextToColumns Destination:=Range("D1"), DataType:=xlDelimited, _
        TextQualifier:=xlNone, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True

Set j = sh2.Range("A1:A" & lr2)

'These find the last cell that the 5 digit codes are in. You use that number later on to shorten the range that you search for.
Set c1 = j.Find(What:=oneCode, _
            LookIn:=xlValues, _
            LookAt:=xlWhole, _
            SearchDirection:=xlPrevious, _
            MatchCase:=False)

Set c2 = j.Find(What:=twoCode, _
            LookIn:=xlValues, _
            LookAt:=xlWhole, _
            SearchDirection:=xlPrevious, _
            MatchCase:=False)

Set c3 = j.Find(What:=threeCode, _
            LookIn:=xlValues, _
            LookAt:=xlWhole, _
            SearchDirection:=xlPrevious, _
            MatchCase:=False)
            
Set c4 = j.Find(What:=fourCode, _
            LookIn:=xlValues, _
            LookAt:=xlWhole, _
            SearchDirection:=xlPrevious, _
            MatchCase:=False)


lr = sh1.Cells(Rows.Count, 3).End(xlUp).Row


'Trims off any spaces in the cells that contain the CPT code + Modifier on the Sheet from DSSupport
For j = 2 To lr2
    sh2.Cells(j, 4).Value = Trim(sh2.Cells(j, 4).Value)
Next j

'Rounds the values on the Fee Schedule sheet to 2 decimal places
For i = 2 To lr
   On Error Resume Next 'Added to encompass the cells that have percentages in them.
    sh1.Cells(i, 4).Value = WorksheetFunction.Round(sh1.Cells(i, 4).Value, 2)
    sh1.Cells(i, 6).Value = WorksheetFunction.Round(sh1.Cells(i, 6).Value, 2)
    sh1.Cells(i, 8).Value = WorksheetFunction.Round(sh1.Cells(i, 8).Value, 2)
    sh1.Cells(i, 10).Value = WorksheetFunction.Round(sh1.Cells(i, 10).Value, 2)
Next i

'Makes all percentage amounts valid
For i = 2 To lr2
    If sh2.Cells(i, 7).Value = " -   " Then
        sh2.Cells(i, 7).Value = 0
    End If
Next i

For i = 2 To lr
    With sh2.Range("D2:D" & c1.Row) 'searches for amount for the first columns of values
        Set Rng = .Find(What:=sh1.Cells(i, 3).Value, _
                        After:=.Cells(.Cells.Count), _
                        LookIn:=xlValues, _
                        LookAt:=xlWhole, _
                        MatchCase:=False)
        If Not Rng Is Nothing Then
            amt = sh2.Cells(Rng.Row, 7).Value
               If amt = 0 Then
                   If CStr(sh2.Cells(Rng.Row, 8).Value) <> CStr(Left(sh1.Cells(i, 4).Value, 2)) Then 'checks if the % values match
                        sh1.Cells(i, 5).Value = "Row #: " & Rng.Row & " Amount: " & sh2.Cells(Rng.Row, 7).Value
                        'added Row number for quick searching
                    End If
                End If
                If amt <> 0 Then
                    If amt <> sh1.Cells(i, 4).Value Then 'checks if the dollar values match, and writes it if they do not.
                            sh1.Cells(i, 5).Value = "Row #: " & Rng.Row & " Amount: " & sh2.Cells(Rng.Row, 7).Value
                            'added Row number for quick searching
                    End If
                End If
        Else
            sh1.Cells(i, 5).Value = "CPT & Modifier Not Found!"
        End If
    End With
    With sh2.Range("D" & c1.Row + 1 & ":D" & c2.Row) 'searches for amount for the second columns of values
        Set Rng = .Find(What:=sh1.Cells(i, 3).Value, _
                        After:=.Cells(.Cells.Count), _
                        LookIn:=xlValues, _
                        LookAt:=xlWhole, _
                        MatchCase:=False)
        If Not Rng Is Nothing Then
            amt = sh2.Cells(Rng.Row, 7).Value
               If amt = 0 Then
                   If CStr(sh2.Cells(Rng.Row, 8).Value) <> CStr(Left(sh1.Cells(i, 6).Value, 2)) Then
                        sh1.Cells(i, 7).Value = "Row #: " & Rng.Row & " Amount: " & sh2.Cells(Rng.Row, 7).Value
                    End If
                End If
                If amt <> 0 Then
                    If amt <> sh1.Cells(i, 6).Value Then
                            sh1.Cells(i, 7).Value = "Row #: " & Rng.Row & " Amount: " & sh2.Cells(Rng.Row, 7).Value
                    End If
                End If
        Else
            sh1.Cells(i, 7).Value = "CPT & Modifier Not Found!"
        End If
   End With
   With sh2.Range("D" & c2.Row + 1 & ":D" & c3.Row) 'searches for amount for the third columns of values
        Set Rng = .Find(What:=sh1.Cells(i, 3).Value, _
                        After:=.Cells(.Cells.Count), _
                        LookIn:=xlValues, _
                        LookAt:=xlWhole, _
                        MatchCase:=False)
        If Not Rng Is Nothing Then
            amt = sh2.Cells(Rng.Row, 7).Value
               If amt = 0 Then
                   If CStr(sh2.Cells(Rng.Row, 8).Value) <> CStr(Left(sh1.Cells(i, 8).Value, 2)) Then
                        sh1.Cells(i, 9).Value = "Row #: " & Rng.Row & " Amount: " & sh2.Cells(Rng.Row, 7).Value
                    End If
                End If
                If amt <> 0 Then
                    If amt <> sh1.Cells(i, 8).Value Then
                            sh1.Cells(i, 9).Value = "Row #: " & Rng.Row & " Amount: " & sh2.Cells(Rng.Row, 7).Value
                    End If
                End If
        Else
            sh1.Cells(i, 9).Value = "CPT & Modifier Not Found!"
        End If
   End With
   With sh2.Range("D" & c3.Row + 1 & ":D" & c4.Row) 'searches for amount for the forth columns of values
        Set Rng = .Find(What:=sh1.Cells(i, 3).Value, _
                        After:=.Cells(.Cells.Count), _
                        LookIn:=xlValues, _
                        LookAt:=xlWhole, _
                        MatchCase:=False)
        If Not Rng Is Nothing Then
            amt = sh2.Cells(Rng.Row, 7).Value
               If amt = 0 Then
                   If CStr(sh2.Cells(Rng.Row, 8).Value) <> CStr(Left(sh1.Cells(i, 10).Value, 2)) Then
                        sh1.Cells(i, 11).Value = "Row #: " & Rng.Row & " Amount: " & sh2.Cells(Rng.Row, 7).Value
                    End If
                End If
                If amt <> 0 Then
                    If amt <> sh1.Cells(i, 10).Value Then
                            sh1.Cells(i, 11).Value = "Row #: " & Rng.Row & " Amount: " & sh2.Cells(Rng.Row, 7).Value
                    End If
                End If
        Else
            sh1.Cells(i, 11).Value = "CPT & Modifier Not Found!"
        End If
    End With
Next i
Application.ScreenUpdating = True

time2 = Timer

MsgBox "Done. Fee Schedule Comparison completed in " & Format((time2 - time1) / 60, "0.00 \min")
End Sub





