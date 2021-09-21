Attribute VB_Name = "Module1"
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

'
'        hdc = CreateDC(TEXT("DISPLAY"), NULL, NULL, NULL);
'        SelectObject(hdc, CreateSolidBrush(RGB(0xFF, 0, 0)));
'        Rectangle(hdc, 0, 0, GetSystemMetrics(SM_CXSCREEN), GetSystemMetrics(SM_CYSCREEN));
'        DeleteObject(SelectObject(hdc, GetStockObject(WHITE_BRUSH)));
'        DeleteDC(hdc);
