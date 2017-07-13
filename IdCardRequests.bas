Attribute VB_Name = "Module4"
Sub IdCardRequests()

Dim dt As String, wbNam As String
Dim OutApp As Object
Dim OutMail As Object
Dim eBody As String
Dim dataRange As Range, oneCell As Range
Dim lr As Long

wbNam = "H:\HealthX\ID Card Requests\ID_CARDs_"
dt = Format(CStr(Now - 1), "mm-dd-yy")
ChDir "H:\HealthX\ID Card Requests\"


    Columns.AutoFit
    Columns("A:E").Select
    Selection.Delete
    Columns("P:P").Select
    Selection.Delete
    Range("A1").Select
    ActiveWorkbook.SaveAs fileName:=wbNam & dt, FileFormat:=xlNormal

    With ActiveSheet.Range("C:C")
        Set dataRange = Range(.Cells(1, 1), .Cells(.Rows.Count, 1).End(xlUp))
    End With
    
    For Each oneCell In dataRange
        If 1 < Application.CountIf(dataRange, oneCell.Value) Then
            With oneCell
                MsgBox "DUPLICATE " & oneCell.Value & ". Check and Remove Duplicate if Necessary."
            End With
        End If
    Next oneCell

    
Set OutApp = CreateObject("Outlook.Application")
Set OutMail = OutApp.CreateItem(0)

eBody = "<BODY style=font-size:11pt;font-family:Calibri>Good Morning,<p><p>Can you please process the attached ID Card Requests?<p><p>Thank you,<p>Greg</BODY>"

  With OutMail
    .To = "alawson@abs-tpa.com; abeavis@abs-tpa.com; gbroadworth@abs-tpa.com; pgonzalez@abs-tpa.com"
    .CC = "AWerner@ABS-TPA.com; djones@abs-tpa.com"
    .BCC = "grodak@abs-tpa.com"
    .Subject = "ID Card Requests"
    .HTMLBody = eBody
    .Attachments.Add ActiveWorkbook.FullName
    .Display
  End With

  
  Set OutMail = Nothing
  Set OutApp = Nothing
End Sub






