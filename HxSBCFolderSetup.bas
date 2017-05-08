Sub HxSBCFolderSetup()
'Need Group ID in Column A, Division Number in Column D, PDFName in Column E

Dim sFile As String, sFolder As String, dFolder As String
Dim fileName As String, fileName2 As String
Dim lr As Long
Dim ws As Worksheet
Dim fso As Variant

Set ws = Sheets("Sheet1") 'Change if needed.
Set fso = CreateObject("Scripting.FileSystemObject")

lr = ws.Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To lr
    If Left(ws.Cells(i, 2).Value, 2) = "SG" Then
        sFolder = "Z:\xxx\xxx\xxx\"
        Else
            sFolder = "Z:\xxx\xxx\xxx\"
    End If
    sFile = "*" & ws.Cells(i, 5).Value & ".pdf"  'SBC File Name
    dFolder = "Z:\xxx\xxx\" & Trim(ws.Cells(i, 1).Value) & "\" & Trim(ws.Cells(i, 4).Value) & "\SBC\" 'Combo of Group & Div
    If Len(Dir("Z:\xxx\xxx\" & Trim(ws.Cells(i, 1).Value), vbDirectory)) = 0 Then
        MkDir ("Z:\xxx\xxx\" & Trim(ws.Cells(i, 1).Value))
    End If
    If Len(Dir("Z:\xxx\xxx\" & Trim(ws.Cells(i, 1).Value) & "\" & ws.Cells(i, 4).Value, vbDirectory)) = 0 Then
        MkDir ("Z:\xxx\xxx\" & Trim(ws.Cells(i, 1).Value) & "\" & ws.Cells(i, 4).Value)
    End If
    If Len(Dir("Z:\xxx\xxx\" & Trim(ws.Cells(i, 1).Value) & "\" & ws.Cells(i, 4).Value & "\SBC", vbDirectory)) = 0 Then
        MkDir ("Z:\xxx\xxx\" & Trim(ws.Cells(i, 1).Value) & "\" & ws.Cells(i, 4).Value & "\SBC")
    End If
    fileName = Dir(sFolder & sFile)
    olName = Dir(dFolder & Replace(sFile, "_", "-"))
    If Not fileName <> "" Then 'Checks if File exists
        ws.Cells(i, 9).Value = "File Not Found"
        ElseIf Not olName <> "" Then 'Checks if File exists in Destination Folder
            fso.CopyFile (sFolder & fileName), dFolder & Trim(ws.Cells(i, 1).Value) & "_GRP_" & Replace(fileName, "_", "-"), True
            ws.Cells(i, 9).Value = "Uploaded"
    Else
        ws.Cells(i, 9).Value = "Duplicate"
    End If
Next i

MsgBox "Done", vbInformation
End Sub

