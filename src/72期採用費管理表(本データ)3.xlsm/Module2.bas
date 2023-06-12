Attribute VB_Name = "Module2"


Option Explicit

Sub import()

    Dim filePath As String
    Dim wbRef As Workbook
    Dim wsRef As Worksheet
    Dim wsTarget As Worksheet
    Dim lastRow1 As Long
    Dim lastRow2 As Long
    
    
    Dim rngRefdate As Range
    Dim rngRefbudget As Range
    Dim rngRefcontent As Range
    Dim rngRefref As Range
    
    Dim rngTargetdate As Range
    Dim rngTargetbudget As Range
    Dim rngTargetcontent As Range
    Dim rngTargetref As Range
    
    Dim BorderRange As Range
    Dim i As Long
    Dim Answer As VbMsgBoxResult
    
    filePath = Range("E3").Value

       If Dir(filePath) <> "" Then
            On Error Resume Next
            Set wbRef = Workbooks.Open(filePath)
            On Error GoTo 0
            
                If Not wbRef Is Nothing Then
            
                    Set wbRef = Workbooks.Open(filePath)
                    Set wsRef = wbRef.Sheets(2)
                    Set wsTarget = ThisWorkbook.Sheets("72�� ���f�[�^")
    
                    lastRow1 = wsTarget.Cells(wsTarget.Rows.Count, "A").End(xlUp).Row
                    lastRow2 = wsRef.Cells(wsRef.Rows.Count, "A").End(xlUp).Row
                    
                    wsRef.Activate
                    wsRef.Range("A1").Select
            
              Answer = MsgBox("�Q�ƃt�@�C�����J���܂����B" & vbCrLf & "2�s�ڂ���" & lastRow2 & "�s�ڂ܂Ō��f�[�^�ɃC���|�[�g���܂��B" & vbCrLf & "�X�����ł����H", vbQuestion + vbYesNo)
    
                If Answer = vbYes Then
    
            Set rngRefdate = wsRef.Range("A2:A" & lastRow2)
            Set rngRefbudget = wsRef.Range("E2:E" & lastRow2)
            Set rngRefcontent = wsRef.Range("F2:F" & lastRow2)
            Set rngRefref = wsRef.Range("G2:G" & lastRow2)
    
            Set rngTargetdate = wsTarget.Range("A" & lastRow1 + 1)
            Set rngTargetbudget = wsTarget.Range("F" & lastRow1 + 1)
            Set rngTargetcontent = wsTarget.Range("E" & lastRow1 + 1)
            Set rngTargetref = wsTarget.Range("G" & lastRow1 + 1)
 
            rngRefdate.Copy
            rngTargetdate.PasteSpecial xlPasteAll
            Application.CutCopyMode = False

            rngRefbudget.Copy
            rngTargetbudget.PasteSpecial xlPasteAll
            Application.CutCopyMode = False

            rngRefcontent.Copy
            rngTargetcontent.PasteSpecial xlPasteAll
            Application.CutCopyMode = False
    
            rngRefref.Copy
            rngTargetref.PasteSpecial xlPasteValues
            Application.CutCopyMode = False

            Set BorderRange = wsTarget.Range("A" & lastRow1 + 1 & ":G" & lastRow1 + lastRow2 - 1)
    
            With BorderRange.Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            End With
                
            For i = 2 To lastRow2
                
            Dim refColor As Long
            refColor = wsRef.Cells(i, "A").Interior.Color
        
            If refColor <> RGB(255, 255, 255) Then
                wsTarget.Range("A" & lastRow1 + i - 1 & ":G" & lastRow1 + i - 1).Interior.Color = refColor
            End If
            Next i
        
            For i = 2 To lastRow2
        
              If wsRef.Range("D" & i).Value = "�w����ʔ�" Then
              wsTarget.Range("B" & lastRow1 + i - 1).Value = "�V��"
              wsTarget.Range("D" & lastRow1 + i - 1).Value = "�I�l��ʔ�"
                  If InStr(1, wsTarget.Cells(lastRow1 + i - 1, 5), "�w����ʔ�") = 0 Then
                     wsTarget.Range("A" & lastRow1 + i - 1 & ":G" & lastRow1 + i - 1).Interior.Color = RGB(255, 255, 0)
                  End If
                
              ElseIf wsRef.Range("D" & i).Value = "���̑�" Then
              wsTarget.Range("B" & lastRow1 + i - 1).Value = ""
              wsTarget.Range("D" & lastRow1 + i - 1).Value = ""
            
                  End If
              Next i
    
            i = 2
    
            For i = 2 To lastRow2
        
            If wsTarget.Cells(lastRow1 + i - 1, 7).Value <> 0 Then
            wsTarget.Cells(lastRow1 + i - 1, 6).ClearContents
            End If
            Next i
 
            wbRef.Close False


            MsgBox lastRow2 - 1 & "�s��ǂݎ��܂����B" & vbCrLf & "�f�[�^��" & lastRow1 + 1 & "�s�ڈȍ~�Ɋi�[����Ă��܂��B" & vbCrLf & "�m�F���Ă��������B"
        
            ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Select
                    
            Else
                MsgBox "�L�����Z�����܂����B"
                wbRef.Close False
        End If
        
        Else
            MsgBox "�t�@�C���A�h���X���m�F���Ă��������B"
        End If

        Else
            MsgBox "�t�@�C���A�h���X���m�F���Ă��������B"
        End If

End Sub
    
