Attribute VB_Name = "Module3"
Declare Function SetWindowPos Lib "user32" _
    (ByVal hwnd As Long, _
    ByVal hWndInsertAfter As Long, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal cx As Long, _
    ByVal cy As Long, _
    ByVal wFlags As Long) As Long
    Public Const HWND_TOPMOST = -1
    Public Const HWND_NOTOPMOST = -2

