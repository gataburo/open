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

'アイデアメモ
' While Not (ActiveSheet.Next Is Nothing)
'            ActiveSheet.Next Is SheetName

'上のプログラムだと非表示のワークシートがあると機能しないので
Sub pagingAll()
    Dim sheet_obj As Object
    Dim i As Integer
    Dim n As Integer
    
    Dim asi As Integer 'Active Sheet Index
    asi = ActiveSheet.Index
    
    ' 表示されているシートの枚数を数える。
    For Each sheet_obj In Sheets
        If sheet_obj.Visible Then
            n = n + 1
        End If
    Next sheet_obj
    
    ' シートにページ数を記入
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
    
    ' 表示されているシートの枚数を数える。
    For Each sheet_obj In Sheets
        If sheet_obj.Visible Then
            n = n + 1
        End If
        If sheet_obj.Index = asi Then
            Exit For
        End If
    Next sheet_obj
    
    ' シートにページ数を記入
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
