Attribute VB_Name = "Module1"
Option Explicit



Sub CommandButton1_Click()
    '�I�������t�@�C������\��
    Dim PathName As String

    PathName = Application.GetOpenFilename("Excel�u�b�N,*.xls;*.xlsx;*.xlsm")
    
    If PathName <> "False" Then
        '�t�@�C�����w�肳�ꂽ�ꍇ
        Sheets(3).Cells(3, 5).Value = PathName
        
    Else
        '�L�����Z�����͉������Ȃ�
        Exit Sub
        
    End If
 
End Sub

