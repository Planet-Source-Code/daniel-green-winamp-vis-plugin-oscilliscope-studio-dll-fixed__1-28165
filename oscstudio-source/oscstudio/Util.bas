Attribute VB_Name = "Util"
'I could have thrown this into Base.frm... but didn't.

Option Explicit

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1

Sub SetOnTop(TheForm As Form, Optional OnTop As Boolean = True)
    Dim InsertAfter As Long
    
    If OnTop Then
        InsertAfter = HWND_TOPMOST
    Else
        InsertAfter = HWND_NOTOPMOST
    End If
    
    SetWindowPos TheForm.hwnd, InsertAfter, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
End Sub
