Attribute VB_Name = "Sleep"
Public Declare Function GetTickCount Lib "kernel32" () As Long

Public Sub wait(ByVal dblmilliseconds As Double)
    Dim dblstart As Double
    Dim dblend As Double
    Dim dbltickcount As Double
    dbltickcount = GetTickCount()
    dblstart = GetTickCount()
    dblend = GetTickCount + dblmilliseconds
    
    Do
    DoEvents
    dbltickcount = GetTickCount()
    
    Loop Until dbltickcount > dblend Or dbltickcount < dblstart
End Sub

