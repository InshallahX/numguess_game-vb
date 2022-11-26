Attribute VB_Name = "Utility"
Public Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long

Function RndAns()
    Randomize
    res = Int(Rnd * 100) + 1
    RndAns = res
End Function

Function RndCh()
    Randomize
    res = Int(Rnd * (GUI.AnswerButton.Count - 1))
    RndCh = res
End Function
