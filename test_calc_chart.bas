Attribute VB_Name = "Module2"
Type DPOINT
    X As Double
    Y As Double
End Type
Type BOXSIZE
    width As Double
    height As Double
End Type

Sub testfunc()
    Dim tc As TransCoordinate
    Set tc = New TransCoordinate
    
    Dim ipt As DPOINT
    Dim bias As DPOINT
    Dim ibs As BOXSIZE
    Dim mbs As BOXSIZE
    Dim log_base As DPOINT
    
    ipt.X = 50
    ipt.Y = 10
    ibs.width = 100
    ibs.height = 200
    mbs.width = 99.99
    mbs.height = 10
    bias.X = 0.01
    bias.Y = 0
    log_base.X = 10
    log_base.Y = 10
    
    Call tc.setIPt(ipt)
    Call tc.setIBS(ibs)
    Call tc.setLogVal(True, False, log_base)
    Debug.Print tc.setMeanVal(mbs, bias).X, tc.getMPt.Y
    
End Sub
