Attribute VB_Name = "Module1"
Option Explicit



Sub CommandButton1_Click()
    '選択したファイル名を表示
    Dim PathName As String

    PathName = Application.GetOpenFilename("Excelブック,*.xls;*.xlsx;*.xlsm")
    
    If PathName <> "False" Then
        'ファイルが指定された場合
        Sheets(3).Cells(3, 5).Value = PathName
        
    Else
        'キャンセル時は何もしない
        Exit Sub
        
    End If
 
End Sub

