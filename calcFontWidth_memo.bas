Attribute VB_Name = "Module1"
'win32apiの宣言

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Type TEXTMETRIC
    tmHeight As Long
    tmAscent As Long
    tmDescent As Long
    tmInternalLeading As Long
    tmExternalLeading As Long
    tmAveCharWidth As Long
    tmMaxCharWidth As Long
    tmWeight As Long
    tmOverhang As Long
    tmDigitizedAspectX As Long
    tmDigitizedAspectY As Long
    tmFirstChar As Byte
    tmLastChar As Byte
    tmDefaultChar As Byte
    tmBreakChar As Byte
    tmItalic As Byte
    tmUnderlined As Byte
    tmStruckOut As Byte
    tmPitchAndFamily As Byte
    tmCharSet As Byte
End Type

' maeke object
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, ByVal W As Long, ByVal E As Long, ByVal O As Long, ByVal W As Long, ByVal i As Long, ByVal u As Long, ByVal S As Long, ByVal C As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal F As String) As Long
' delete object
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
' format
Public Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsA" (ByVal hdc As Long, lpMetrics As TEXTMETRIC) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
' textout
Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long

' rect 計算指示
Public Const DT_CALCRECT = &H400
' 書式(bold)
Public Const FW_NORMAL = &H190
Public Const FW_BOLD = &H2BC
' キャラセット
Public Const DEFAULT_CHARSET = &H1

Sub drawEraseLineFit()
    Dim txtr As Range
    Dim width As Long
    Dim cwidth As Long
    Dim i As Integer
    Dim strlen As Integer
    Dim x As Integer
    Dim y As Integer
    Dim initx As Integer
    Dim height As Long
    
    Dim fname As String
    Dim fsize As Long
    Dim nbold As Long
    
    x = Range("b2").Left
    initx = Range("b2").Left
    y = Range("b2").Top
    Set txtr = Range("b2")
    
    width = 0
    
    strlen = txtr.Characters.Count
    While i <= strlen
        fname = txtr.Characters(i, 1).Font.Name
        fsize = txtr.Characters(i, 1).Font.Size
        j = 1
        While (fname = txtr.Characters(i + j, 1).Font.Name) And (fsize = txtr.Characters(i + j, 1).Font.Size) And (i + j <= strlen) And (InStr(txtr.Characters(i + j, 1).text, vbLf) = 0) And (InStr(txtr.Characters(i + j, 1).text, " ") = 0) And (InStr(txtr.Characters(i + j, 1).text, "　") = 0)
            j = j + 1
        Wend
        cwidth = calcCharWidth(txtr.Characters(i, j).text, fname, fsize, FW_NORMAL)
        width = width + cwidth
        Debug.Print "cwidth = " + Str(cwidth)
        Debug.Print "sum = " + Str(width)
        
        If (InStr(txtr.Characters(i + j, 1).text, vbLf) <> 0) Or (i + j >= strlen) Or (InStr(txtr.Characters(i + j, 1).text, " ") <> 0) Or (InStr(txtr.Characters(i + j, 1).text, "　") <> 0) Then
            height = fsize
            Call drawEraseLines(x, y, height, width)
            
            If InStr(txtr.Characters(i + j, 1).text, vbLf) <> 0 Then
                x = initx
                y = y + fsize
            ElseIf (InStr(txtr.Characters(i + j, 1).text, " ") <> 0) Or (InStr(txtr.Characters(i + j, 1).text, "　") <> 0) Then
                x = x + width + calcCharWidth(txtr.Characters(i + j, 1).text, fname, fsize, FW_NORMAL)
                y = y
            End If
            
            width = 0
            j = j + 1
        End If
        
        i = i + j
    Wend
    
End Sub

Function calcCharWidth(text As String, fname As String, fheight As Long, nbold As Long) As Long
    Dim hwnd As Long
    Dim hdc As Long
    Dim hcdc As Long
    Dim hfont As Long
    Dim trect As RECT
    Dim lptm As TEXTMETRIC
    
    'ハンドルのセット
    hwnd = 0 '画面全体
    hdc = GetDC(hwnd)
    hcdc = CreateCompatibleDC(hdc)
    
    'フォントのセット
    hfont = CreateFont(fheight, 0, 0, 0, nbold, 0, 0, 0, DEFAULT_CHARSET, 0, 0, 0, 0, fname)
    Call SelectObject(hdc, hfont)
    Call DeleteObject(hfont)
    
    'サイズ調整 どうやらExcelでは、tmheightに対して何かしらの倍率をかけている模様 今回は、文字の実質サイズを導出している
    'height/(ascent - internalLeading + descent)*height は全体文字高さ(余白含む)÷実際の文字高さ(余白なし)×指定したい文字高さ
    '実際の文字高さはAscent(=ベースラインより上の文字高さ)からAscent内の余白高さ(=internalLeading)を引き、Descent(=ベースラインより下の文字高さ(余白なし))を足すことで求められる。
    Call GetTextMetrics(hdc, lptm)
    hfont = CreateFont(Int(fheight * (lptm.tmHeight / (lptm.tmAscent - lptm.tmInternalLeading + lptm.tmDescent))), 0, 0, 0, nbold, 0, 0, 0, DEFAULT_CHARSET, 0, 0, 0, 0, fname)
    Call SelectObject(hdc, hfont)
    
    'サイズ計測
    Call DrawText(hdc, text, -1, trect, DT_CALCRECT)
    
    'ハンドルの解放
    Call DeleteObject(hfont)
    Call DeleteObject(hcdc)
    Call ReleaseDC(hwnd, ps)
    
    '返り値設定
    calcCharWidth = trect.Right - trect.Left
    Debug.Print trect.Right
End Function

Function drawEraseLines(x As Integer, y As Integer, height As Long, width As Long)
    Dim y1 As Long
    Dim y2 As Long
    y1 = y + height / 3
    y2 = y1 + height / 3
    With ActiveSheet.Shapes.AddLine(x, y1, x + width, y1).Line
        .ForeColor.RGB = vbBlack
    End With
    With ActiveSheet.Shapes.AddLine(x, y2, x + width, y2).Line
        .ForeColor.RGB = vbBlack
    End With
End Function
