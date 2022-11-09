VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TransCoordinate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Type DPOINT
'    X As Double
'    Y As Double
'End Type
'Type BOXSIZE
'    width As Double
'    height As Double
'End Type

Dim ipt As DPOINT
Dim mpt As DPOINT
Dim mbias As DPOINT
Dim ibs As BOXSIZE
Dim mbs As BOXSIZE
Dim uselogx As Boolean
Dim uselogy As Boolean
Dim lbase As DPOINT

Private Sub class_initialize()
    ipt.X = 0
    ipt.Y = 0
    mpt.X = 0
    mpt.Y = 0
    mbias.X = 0
    mbias.Y = 0
    ibs.height = 0
    ibs.width = 0
    mbs.height = 0
    mbs.width = 0
    uselogx = False
    uselogy = False
    lbase.X = 10
    lbase.Y = 10
End Sub

Function setIPt(pt As DPOINT) As DPOINT
    ipt = pt

    Call calcMPt
    setIPt = mpt
End Function

Function setIBS(bs As BOXSIZE) As DPOINT
    ibs = bs
    
    Call calcMPt
    setIBS = mpt
End Function

Function setMeanVal(bs As BOXSIZE, bias As DPOINT) As DPOINT
    mbs = bs
    mbias = bias
    
    Call calcMPt
    setMeanVal = mpt
End Function

Function setLogVal(use_logx As Boolean, use_logy As Boolean, log_base As DPOINT) As DPOINT
    uselogx = use_logx
    uselogy = use_logy
    lbase = log_base
    
    Call calcMPt
    setLogVal = mpt
End Function

Function getMPt() As DPOINT
    Call calcMPt
    getMPt = mpt
End Function

Private Function calcMPt()
    If ibs.width <= 0 Or ibs.height <= 0 Or mbs.width <= 0 Or mbs.height <= 0 Then
        mpt.X = 0
        mpt.Y = 0
    Else
        If uselogx Then
            If mbias.X <= 0 Or lbase.X <= 0 Then
                mpt.X = 0
            Else
                mpt.X = lbase.X ^ (ipt.X / ibs.width * (Log(mbs.width + mbias.X) / Log(lbase.X) - Log(mbias.X) / Log(lbase.X)))
            End If
        Else
            mpt.X = ipt.X / ibs.width * mbs.width + mbias.X
        End If
        If uselogy Then
            If mbias.Y <= 0 Or lbase.Y <= 0 Then
                mpt.Y = 0
            Else
                mpt.Y = lbase.Y ^ (ipt.Y / ibs.height * (Log(mbs.height + mbias.Y) / Log(lbase.Y) - Log(mbias.Y) / Log(lbase.Y)))
            End If
        Else
            mpt.Y = ipt.Y / ibs.height * mbs.height + mbias.Y
        End If
    End If
End Function
