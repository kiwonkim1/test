Attribute VB_Name = "Module1"
Option Explicit

'GIT　テストです

Sub CommandButton1_Click()
    '選択したファイル名を表示
    Dim PathName1   As String: PathName1 = Application.GetOpenFilename("Excelブック,*.xls;*.xlsx;*.xlsm")
    Dim wb As Workbook
    If PathName1 <> "False" Then
        'ファイルが指定された場合
        Sheets("72期 元データ").Cells(3, 5).Value = PathName1
        
    Else
        'キャンセル時は何もしない
        Exit Sub

    End If

End Sub


Sub OpenButton()

Dim wb As Workbook
Dim PathName1 As String
Dim Answer As Byte
    
    Answer = MsgBox("ファイルを開きますか？", vbYesNo + vbQuestion)
    
    If Answer = vbYes Then
    PathName1 = Sheets("72期 元データ").Cells(3, 5).Value
    
        If Dir(PathName1) <> "" Then
            On Error Resume Next
            Set wb = Workbooks.Open(PathName1)
            On Error GoTo 0
            
            If Not wb Is Nothing Then
                wb.Sheets(2).Range("A1").Activate
            Else
                MsgBox "ファイルアドレスを確認してください。"
            End If
    Else
            MsgBox "ファイルアドレスを確認してください。"
        End If
    End If
End Sub

