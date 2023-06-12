Attribute VB_Name = "Module3"
Option Explicit

Sub import2()

    Dim filePath3 As String
    Dim wbRef3 As Workbook
    Dim wsRef3 As Worksheet
    Dim wsTarget3 As Worksheet
    Dim lastRow3 As Long
    Dim lastRow4 As Long
    
    
    Dim rngRefdate3 As Range
    Dim rngRefbudget3 As Range
    Dim rngRefcontent3 As Range
    Dim rngRefref3 As Range
    
    Dim rngTargetdate3 As Range
    Dim rngTargetbudget3 As Range
    Dim rngTargetcontent3 As Range
    Dim rngTargetref3 As Range
    
    Dim BorderRange3 As Range
    Dim j As Long
    Dim Answer3 As VbMsgBoxResult
    Dim startRow3 As Long
    Dim endRow3 As Long
        
        filePath3 = Range("E3").Value
        
        If Dir(filePath3) <> "" Then
            On Error Resume Next
            Set wbRef3 = Workbooks.Open(filePath3)
            On Error GoTo 0
        
            If Not wbRef3 Is Nothing Then
                Set wsRef3 = wbRef3.Sheets(2)
                Set wsTarget3 = ThisWorkbook.Sheets("72期 元データ")
                
                wsRef3.Activate
                wsRef3.Range("A1").Select
                     
                startRow3 = Application.InputBox("参照ファイルを開きました。" & vbCrLf & "読み取るデータの初行を入力してください。" & vbCrLf & "初行番号 : ", Type:=1)
                endRow3 = Application.InputBox("読み取るデータの最終行を入力してください。" & vbCrLf & "最終行番号 : ", Type:=1)
                
                If startRow3 > 1 And endRow3 >= startRow3 Then
                    lastRow3 = wsTarget3.Cells(wsTarget3.Rows.Count, "A").End(xlUp).Row
                    lastRow4 = endRow3 - startRow3 + 1
                    
                    Set rngRefdate3 = wsRef3.Range("A" & startRow3 & ":A" & endRow3)
                    Set rngRefbudget3 = wsRef3.Range("E" & startRow3 & ":E" & endRow3)
                    Set rngRefcontent3 = wsRef3.Range("F" & startRow3 & ":F" & endRow3)
                    Set rngRefref3 = wsRef3.Range("G" & startRow3 & ":G" & endRow3)
    
                    Set rngTargetdate3 = wsTarget3.Range("A" & lastRow3 + 1)
                    Set rngTargetbudget3 = wsTarget3.Range("F" & lastRow3 + 1)
                    Set rngTargetcontent3 = wsTarget3.Range("E" & lastRow3 + 1)
                    Set rngTargetref3 = wsTarget3.Range("G" & lastRow3 + 1)
    
                    rngRefdate3.Copy
                    rngTargetdate3.PasteSpecial xlPasteAll
                    Application.CutCopyMode = False
    
                    rngRefbudget3.Copy
                    rngTargetbudget3.PasteSpecial xlPasteAll
                    Application.CutCopyMode = False
    
                    rngRefcontent3.Copy
                    rngTargetcontent3.PasteSpecial xlPasteAll
                    Application.CutCopyMode = False
    
                    rngRefref3.Copy
                    rngTargetref3.PasteSpecial xlPasteValues
                    Application.CutCopyMode = False
    
                    Set BorderRange3 = wsTarget3.Range("A" & lastRow3 + 1 & ":G" & lastRow3 + lastRow4)
    
                    With BorderRange3.Borders
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                    End With
    
                    If wsRef3.Cells(startRow3, "A").Interior.Color <> RGB(255, 255, 255) Then
                    Dim refColor4 As Long
                    refColor4 = wsRef3.Cells(startRow3, "A").Interior.Color
                    wsTarget3.Range("A" & lastRow3 + 1 & ":G" & lastRow3 + 1).Interior.Color = refColor4
                    End If
                    
                    For j = 2 To lastRow4
                        Dim refColor3 As Long
                        refColor3 = wsRef3.Cells(startRow3 + j - 1, "A").Interior.Color
    
                        If refColor3 <> RGB(255, 255, 255) Then
                            wsTarget3.Cells(lastRow3 + j, 6).ClearContents
                            wsTarget3.Range("A" & lastRow3 + j & ":G" & lastRow3 + j).Interior.Color = refColor3
                        End If
                    Next j
                    
                    For j = 2 To lastRow4 + 1
                        If wsRef3.Range("D" & startRow3 + j - 2).Value = "学生交通費" Then
                            wsTarget3.Range("B" & lastRow3 + j - 1).Value = "新卒"
                            wsTarget3.Range("D" & lastRow3 + j - 1).Value = "選考交通費"
                            If InStr(1, wsTarget3.Cells(lastRow3 + j - 1, 5), "学生交通費") = 0 Then
                                wsTarget3.Range("A" & lastRow3 + j - 1 & ":G" & lastRow3 + j - 1).Interior.Color = RGB(255, 255, 0)

                            End If
                            
                        ElseIf wsRef3.Range("D" & startRow3 + j - 1).Value = "その他" Then
                            wsTarget3.Range("B" & lastRow3 + j - 1).Value = ""
                            wsTarget3.Range("D" & lastRow3 + j - 1).Value = ""
                        End If
                    Next j
                    
                    If startRow3 = endRow3 And wsTarget3.Cells(lastRow3 + 1, 7).Value <> 0 Then
                          wsTarget3.Cells(lastRow3 + 1, 6).ClearContents
                                         
                    Else
                    
                    For j = 2 To lastRow4
                            If wsTarget3.Cells(lastRow3 + j - 1, 7).Value <> 0 Then
                            wsTarget3.Cells(lastRow3 + j - 1, 6).ClearContents
                            End If
                        
                        Next j
                    End If
                    
                    wbRef3.Close False
    
                    MsgBox "データを読み取りました。"
    
                    ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Select
            Else
                    MsgBox "入力範囲を確認してください。"
                    wbRef3.Close False
                End If
                
            Else
                MsgBox "ファイルアドレスを確認してください。"
            End If
        Else
            MsgBox "ファイルアドレスを確認してください。"
        End If

End Sub


