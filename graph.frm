VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} graph 
   Caption         =   "graph_form"
   ClientHeight    =   6510
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17070
   OleObjectBlob   =   "graph.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "graph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Const GWL_STYLE = (-16)
Private Const GRAPH_IMAGE As String = "E:\kodeemon\make someting\vba\Graph.bmp"

Private Sub UserForm_Initialize()
    Dim hwnd, nindex As Long
    
    hwnd = FindWindow(vbNullString, "graph_form")
    Debug.Print hwnd
    nindex = GetWindowLong(hwnd, GWL_EXSTYLE)
    Call SetWindowLong(hwnd, GWL_EXSTYLE, nindex Or WS_EX_LAYERED)
    Call SetLayeredWindowAttributes(hwnd, 0, 150, LWA_ALPHA)

    nindex = GetWindowLong(hwnd, GWL_STYLE)
    Call SetWindowLong(hwnd, GWL_STYLE, nindex Or WS_THICKFRAME)
    
    '�O���t�̑��݃`�F�b�N
    If ActiveSheet.ChartObjects.Count = 0 Then Exit Sub
    
    '�O���t���摜�Ƃ��ĕۑ�
    ActiveSheet.ChartObjects(1).Chart.Export GRAPH_IMAGE
    
    '�摜�t�@�C����Image�ɓǂݍ���
    If Len(Dir(GRAPH_IMAGE)) > 0 Then
        With Image1
            '.PictureSizeMode = fmPictureSizeModeClip      '�g��E�k���Ȃ�
            '.PictureAlignment = fmPictureAlignmentCenter  '�����z�u
            '.BorderStyle = fmBorderStyleNone              '�g�Ȃ�
            .Picture = LoadPicture(GRAPH_IMAGE)
        End With
        '�摜�t�@�C�����폜
        'Kill GRAPH_IMAGE
    End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    '�摜�t�@�C�����폜
    If Len(Dir(GRAPH_IMAGE)) > 0 Then
        Kill GRAPH_IMAGE
    End If
End Sub

Private Sub UserForm_Resize()
    Image1.Width = graph.Width
    Image1.Height = graph.Height
    'If Len(Dir(GRAPH_IMAGE)) > 0 Then
        'Image1.Picture = LoadPicture(GRAPH_IMAGE)
        'Debug.Print "resize!"
    'End If
End Sub
