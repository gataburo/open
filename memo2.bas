Attribute VB_Name = "Module1"
Type POINTAPI
    x As Long
    y As Long
End Type
Public Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function CreateDC Lib "gdi32.dll" Alias "CreateDCA" (ByVal lpszDrive As String, ByVal lpszDevice As String, ByVal lpszOutput As Long, lpInitData As Any) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal nLeftRect As Long, ByVal nTopRect As Long, ByVal nRightRect As Long, ByVal nBottomRect As Long) As Long
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Public Const WS_EX_TOPMOST = &H8
Public Const WS_EX_TRANSPARENT = &H20
Public Const WS_OVERLAPPED = &H0
Public Const SW_SHOW = &H5
Public Const SW_CXSCREEN = &H0
Public Const SW_CYSCREEN = &H1
Public Const WHITE_BRUSH = &H0

Sub makeWindow()
    Dim hwnd As Long
    Dim buf As Long
    
    hwnd = CreateWindowEx((WS_EX_TOPMOST Or WS_EX_TRANSPARENT), "STATIC", "Kitty on your lap", WS_OVERLAPPEDWINDOW, 100, 100, 200, 200, 0, 0, 0, 0)
    
    If (hwnd = Null) Then Exit Sub
    
    buf = ShowWindow(hwnd, SW_SHOW)
    
End Sub

Sub dyeRed()
    Dim hdc As Long
    Dim buf As Long
    
    hdc = CreateDC("DISPLAY", 0, 0, 0)
    buf = SelectObject(hdc, CreateSolidBrush(RGB(&HFF, 0, 0)))
    buf = Rectangle(hdc, 0, 0, GetSystemMetrics(SM_CXSCREEN), GetSystemMetrics(SM_CYSCREEN))
    buf = DeleteObject(SelectObject(hdc, GetStockObject(WHITE_BRUSH)))
    buf = DeleteDC(hdc)
    
End Sub

Sub cursordayo()
    Dim cursorpos As POINTAPI
    Dim buf As Long
    
    MsgBox "基準値(1)にカーソルを合わせてEnter"
    buf = GetCursorPos(cursorpos)
    Range("B2") = cursorpos.x
    Range("C2") = cursorpos.y
    MsgBox "基準値(10)にカーソルを合わせてEnter"
    buf = GetCursorPos(cursorpos)
    Range("B3") = cursorpos.x
    Range("C3") = cursorpos.y
    MsgBox "知りたい位置にカーソルを合わせてEnter"
    buf = GetCursorPos(cursorpos)
    Range("B4") = cursorpos.x
    Range("C4") = cursorpos.y
    
    ' sheetに次の式をセットする事
    ' B5に ABS($B$2-B3)それをC5, B6, C6 にオートフィル
    ' B7はグラフによって変える ※C7はオートフィル
    ' 通常のグラフ: B6/B5
    ' 対数グラフ: N^(B6/B5) ※Nは底
End Sub

Sub readChart()
    Dim cursorpos As POINTAPI
    Dim i As Integer
    Dim base_row As Integer
    Dim base_col As Integer
    Dim buf As Long
    
    base_row = 7
    base_col = 5
    
    For i = 0 To 2
        Select Case i
        Case 0
            MsgBox "基準値1にカーソルを合わせてEnter"
        Case 1
            MsgBox "基準値2にカーソルを合わせてEnter"
        Case 2
            MsgBox "知りたい位置にカーソルを合わせてEnter"
        End Select
        buf = GetCursorPos(cursorpos)
        Cells(base_row + i, base_col) = cursorpos.x
        Cells(base_row + i, base_col + 1) = cursorpos.y
        Cells(base_row + i, base_col + 2) = Abs(cursorpos.x - Cells(base_row, base_col).Value)
        Cells(base_row + i, base_col + 3) = Abs(cursorpos.y - Cells(base_row, base_col + 1).Value)
    Next i
    
    Select Case Range("C2").Value
    Case 1
        Range("c9") = "=C7+G9/G8*(C8-C7)"
        Range("d9") = "=D7+H9/H8*(D8-D7)"
    Case 4
        Range("c9") = "=C3^(G9/G8*(log(C8,C3)-log(C7,C3)))"
        Range("d9") = "=C3^(H9/H8*(log(D8,C3)-log(D7,C3)))"
    End Select
End Sub
