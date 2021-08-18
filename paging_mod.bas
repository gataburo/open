Attribute VB_Name = "Module1"
Sub pagingAll_v0()
    Dim i As Integer
    Dim n As Integer
    
    Dim asi As Integer 'Active Sheet Index
    asi = ActiveSheet.Index
    
    n = Worksheets.Count
    For i = 1 To n
        Worksheets(i).Select
        Range("a1") = i
        Range("b1") = n
    Next
    
    Worksheets(asi).Select
End Sub

Sub pagingUntilSelected_v0()
    Dim i As Integer
    Dim n As Integer
    
    n = ActiveSheet.Index
    For i = 1 To n
        Worksheets(i).Select
        Range("a1") = i
        Range("b1") = n
    Next
End Sub

'�A�C�f�A����
' While Not (ActiveSheet.Next Is Nothing)
'            ActiveSheet.Next Is SheetName

'��̃v���O�������Ɣ�\���̃��[�N�V�[�g������Ƌ@�\���Ȃ��̂�
Sub pagingAll()
    Dim sheet_obj As Object
    Dim i As Integer
    Dim n As Integer
    
    Dim asi As Integer 'Active Sheet Index
    asi = ActiveSheet.Index
    
    ' �\������Ă���V�[�g�̖����𐔂���B
    For Each sheet_obj In Sheets
        If sheet_obj.Visible Then
            n = n + 1
        End If
    Next sheet_obj
    
    ' �V�[�g�Ƀy�[�W�����L��
    i = 1
    For Each sheet_obj In Sheets
        If (sheet_obj.Visible) Then
            sheet_obj.Select
            Range("a5") = i
            Range("b5") = n
            i = i + 1
        End If
    Next sheet_obj
    
    Worksheets(asi).Select
End Sub

Sub pagingUntilSelected2()
    Dim sheet_obj As Object
    Dim i As Integer
    Dim n As Integer
    
    Dim asi As Integer 'Active Sheet Index
    asi = ActiveSheet.Index
    
    ' �\������Ă���V�[�g�̖����𐔂���B
    For Each sheet_obj In Sheets
        If sheet_obj.Visible Then
            n = n + 1
        End If
        If sheet_obj.Index = asi Then
            Exit For
        End If
    Next sheet_obj
    
    ' �V�[�g�Ƀy�[�W�����L��
    i = 1
    For Each sheet_obj In Sheets
        If (sheet_obj.Visible) Then
            sheet_obj.Select
            Range("a5") = i
            Range("b5") = n
            i = i + 1
        End If
        If sheet_obj.Index = asi Then
            Exit For
        End If
    Next sheet_obj
End Sub
