Attribute VB_Name = "Module2"
Option Explicit

Dim PathName As String
Dim wbRef As Workbook
Dim wsRef As Worksheet
Dim wsTarget As Worksheet
Dim lastRow As Long
Dim lastRow2 As Long
Dim i As Long




Sub import_data()
    
    Dim rngRefdate As Range
    Dim rngRefbudget As Range
    Dim rngRefcontent As Range
    Dim rngRefmore As Range
    

    Dim rngTargetdate As Range
    Dim rngTargetbudget As Range
    Dim rngTargetcontent As Range
    Dim rngTargetmore As Range
    Dim answer As Variant
    
    PathName = Sheets(3).Cells(3, 5).Value
    On Error Resume Next
    If Range("E3") = "" Then
    
        MsgBox "�t�@�C�������m�F���Ă��������B"
        Exit Sub
    
    ElseIf Dir(PathName) <> "" Then
        Set wbRef = Workbooks.Open(PathName)
        
    Else
    MsgBox "�t�@�C�������m�F���Ă��������B"
    Exit Sub
    End If
    
    On Error GoTo 0

    Set wsRef = wbRef.Sheets(2)
    
    Set wsTarget = ThisWorkbook.Sheets(3)
    
    lastRow = wsTarget.Cells(wsTarget.Rows.Count, "A").End(xlUp).Row
    lastRow2 = wsRef.Cells(wsRef.Rows.Count, "A").End(xlUp).Row
    
    
    answer = MsgBox("�V�[�g�̂Q�s�ڂ���" & lastRow2 & "�s�ڂ�I�����Ă��܂��B" & vbCrLf & "�C���|�[�g���܂����H", vbQuestion + vbYesNo)
    Select Case answer
        Case vbYes
        
        Set rngRefdate = wsRef.Range("A2:A" & lastRow2)
        Set rngRefbudget = wsRef.Range("E2:E" & lastRow2)
        Set rngRefcontent = wsRef.Range("F2:F" & lastRow2)
        Set rngRefmore = wsRef.Range("G2:G" & lastRow2)
    
        Set rngTargetdate = wsTarget.Range("A" & lastRow + 1)
        Set rngTargetbudget = wsTarget.Range("F" & lastRow + 1)
        Set rngTargetcontent = wsTarget.Range("E" & lastRow + 1)
        Set rngTargetmore = wsTarget.Range("G" & lastRow + 1)
    
        rngRefdate.Copy rngTargetdate
        rngRefbudget.Copy rngTargetbudget
        rngRefcontent.Copy rngTargetcontent
        rngRefmore.Copy rngTargetmore
        
        Call TPexpense
        
        Call remove_othercost
        
        Call make_lines
    
        MsgBox "�ǂݎ�肪�������܂����B"
        
        wbRef.Close
        
        MsgBox "�f�[�^��" & lastRow + 1 & "�s�ڈȍ~�Ɋi�[����Ă��܂��B" & vbCrLf & "�m�F���Ă��������B"
        
        ActiveSheet.Cells(lastRow, 1).Select
        
        If InStr(1, wsTarget.Cells(lastRow + i - 1, 5), "�w����ʔ�") = 0 Then
            MsgBox ("�L���̕s��������܂��B" & vbCrLf & "���F���ω������������m�F���Ă��������B")
        End If
        
    End Select

End Sub

Sub OpenButton()

Dim answer As Byte
    
    answer = MsgBox("�t�@�C�����J���܂����H", vbYesNo + vbQuestion)
    
    If answer = vbYes Then
       PathName = Sheets(3).Cells(3, 5).Value
       On Error Resume Next
       
        If Range("E3") = "" Then
        
            MsgBox "�t�@�C�������m�F���Ă��������B"
        
        ElseIf Dir(PathName) <> "" Then
            
            Set wbRef = Workbooks.Open(PathName)
           
            
        Else
            MsgBox "�t�@�C�������m�F���Ă��������B"
            Exit Sub
        End If
        On Error GoTo 0
    End If
End Sub


Sub make_lines()
    Dim borderrange As Range
    Set borderrange = wsTarget.Range("A" & lastRow + 1 & ":G" & lastRow + lastRow2 - 1)
        
    With borderrange.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With

End Sub

Sub remove_othercost()

    Dim refColor As Long

    i = 2
        
    For i = 2 To lastRow2
    refColor = wsRef.Cells(i, "A").Interior.Color
        If wsTarget.Cells(lastRow + i - 1, 7).Value <> 0 Then
            wsTarget.Cells(lastRow + i - 1, 6).ClearContents
            wsTarget.Range("A" & lastRow + i - 1 & ":G" & lastRow + i - 1).Interior.Color = refColor
        End If
    Next i


End Sub

Sub TPexpense()
    i = 2
    For i = 2 To lastRow2
       
        If wsRef.Cells(i, 4).Value = "�w����ʔ�" Then
            wsTarget.Cells(lastRow + i - 1, 2).Value = "�V��"
            wsTarget.Cells(lastRow + i - 1, 4).Value = "�I�l��ʔ�"
                If InStr(1, wsTarget.Cells(lastRow + i - 1, 5), "�w����ʔ�") = 0 Then
                    wsTarget.Range("A" & lastRow + i - 1 & ":G" & lastRow + i - 1).Interior.ColorIndex = 6
                End If
        End If
    Next i

End Sub
