Attribute VB_Name = "Module2"
Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)


Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long

Declare Function GetParent& Lib "user32" (ByVal hwnd As Long)
Declare Function GetCursorPos Lib "user32" (lpPoint As Where) As Long
Declare Function GetClassName Lib "user32.dll" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Declare Function GetWindowTextLength Lib "user32.dll" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal nMaxCount As Long) As Long
Declare Function EnumWindows Lib "user32.dll" (ByVal lpenumfunc As Long, ByVal lParam As Long) As Long
Declare Function EnumChildWindows Lib "user32.dll" (ByVal hwndParent As Long, ByVal lpenumfunc As Long, ByVal lParam As Long) As Long
Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long


Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long


Type Where
    Pointa As Long
    Pointb As Long
    End Type
Public Function EnumChildProc(ByVal hwnd As Long, ByVal lParam As Long) As Long
Dim slength As Long, wintext As String, slengthc As Long
Dim buffer As String, classname As String, wincap As String
Dim retval As Long
Static winnum As Integer
winnum = winnum + 1

If start = 1 Then
    Form1.List1.Clear
    start = 0
End If

classname = Space(255)
slength = GetClassName(hwnd, classname, 255)
classname = Left(classname, slength)

slengthc = GetWindowTextLength(hwnd) + 1
If slengthc > 128 Then slengthc = 128
wincap = ""
If slengthc > 1 Then
    buffer = Space(slengthc)
    retval = GetWindowText(hwnd, buffer, slengthc)
    wincap = Left(buffer, slengthc - 1)
End If

If onlycaps <> 1 Then
    'Form1.List1.AddItem classname & " : " & wincap
    Form1.List4.AddItem hwnd
    'Form1.List1.AddItem hwnd & "    " & classname & " : " & wincap
    Form1.Combo1.AddItem wincap
Else
    If wincap <> "" Then
        Form1.List1.AddItem hwnd & "    " & classname & " : " & wincap
    End If
End If
EnumChildProc = 1
End Function


Public Sub SendText(hwnd, txt)
Call SetFocusAPI(hwnd)
Call SendMessageByString(hwnd, WM_SETTEXT, 0, "")
Call SendMessageByString(hwnd, WM_SETTEXT, 0, txt)
End Sub

Public Sub SendAscii(hwnd, ascii)
Call SetFocusAPI(hwnd)
Call SendMessageByNum(hwnd, WM_CHAR, ascii, 0)
End Sub

Function WinCaption(shwnd)

Dim wintext As String
Dim slength As Long
Dim retval As Long

slength = SendMessage(shwnd, WM_GETTEXTLENGTH, ByVal CLng(0), ByVal CLng(0)) + 1

retval = SendMessage(shwnd, WM_GETTEXT, ByVal slength, ByVal wintext)

wintext = Left(wintext, retval)
WinCaption = wintext
End Function
