Attribute VB_Name = "Module1"
Option Explicit

Sub CommandButton1_Click()
    '�I�������t�@�C������\��
    Dim PathName1   As String: PathName1 = Application.GetOpenFilename("Excel�u�b�N,*.xls;*.xlsx;*.xlsm")
    Dim wb As Workbook
    If PathName1 <> "False" Then
        '�t�@�C�����w�肳�ꂽ�ꍇ
        Sheets("72�� ���f�[�^").Cells(3, 5).Value = PathName1
        
    Else
        '�L�����Z�����͉������Ȃ�
        Exit Sub

    End If

End Sub


Sub OpenButton()

Dim wb As Workbook
Dim PathName1 As String
Dim Answer As Byte
    
    Answer = MsgBox("�t�@�C�����J���܂����H", vbYesNo + vbQuestion)
    
    If Answer = vbYes Then
    PathName1 = Sheets("72�� ���f�[�^").Cells(3, 5).Value
    
        If Dir(PathName1) <> "" Then
            On Error Resume Next
            Set wb = Workbooks.Open(PathName1)
            On Error GoTo 0
            
            If Not wb Is Nothing Then
                wb.Sheets(2).Range("A1").Activate
            Else
                MsgBox "�t�@�C���A�h���X���m�F���Ă��������B"
            End If
    Else
            MsgBox "�t�@�C���A�h���X���m�F���Ă��������B"
        End If
    End If
End Sub

