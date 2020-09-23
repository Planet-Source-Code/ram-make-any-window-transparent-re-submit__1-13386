VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000018&
   BorderStyle     =   0  'None
   Caption         =   "Transperent"
   ClientHeight    =   2625
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5880
   LinkTopic       =   "Form1"
   ScaleHeight     =   2625
   ScaleWidth      =   5880
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   3960
      TabIndex        =   20
      Top             =   4560
      Width           =   2055
   End
   Begin VB.ListBox List1 
      Height          =   450
      Left            =   0
      TabIndex        =   18
      Top             =   5640
      Width           =   2415
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000018&
      Caption         =   "Controls"
      Height          =   1635
      Left            =   480
      TabIndex        =   16
      Top             =   600
      Visible         =   0   'False
      Width           =   5235
      Begin Project1.AOLCmd AOLCmd5 
         Height          =   375
         Left            =   3600
         TabIndex        =   21
         Top             =   240
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         Caption         =   "Undo transperent"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "Form1.frx":0000
         ForeColor       =   0
         BackColor       =   -2147483624
         StandardColors  =   0   'False
         BackColorClick  =   12648447
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1560
         TabIndex        =   19
         Top             =   1200
         Visible         =   0   'False
         Width           =   2175
      End
      Begin Project1.AOLCmd AOLCmd4 
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         Caption         =   "Make Transperent"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "Form1.frx":001C
         ForeColor       =   0
         BackColor       =   -2147483624
         StandardColors  =   0   'False
         BackColorClick  =   12648447
      End
   End
   Begin VB.ListBox List4 
      Height          =   1425
      Left            =   3360
      TabIndex        =   15
      Top             =   3360
      Width           =   2655
   End
   Begin VB.ListBox List3 
      Height          =   1425
      Left            =   6600
      TabIndex        =   14
      Top             =   5280
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ListBox List2 
      Height          =   1425
      Left            =   2760
      TabIndex        =   13
      Top             =   5760
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H80000018&
      Caption         =   "&fade out, fade in"
      Height          =   375
      Left            =   1920
      TabIndex        =   9
      Top             =   5400
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000018&
      Caption         =   "&undo Transparency"
      Height          =   375
      Left            =   2880
      TabIndex        =   8
      Top             =   5400
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000018&
      Caption         =   "Make &transperent"
      Height          =   375
      Left            =   600
      TabIndex        =   7
      Top             =   4800
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   285
      Left            =   600
      TabIndex        =   3
      Text            =   "255"
      Top             =   3960
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000018&
      Caption         =   "Controls"
      Height          =   1635
      Left            =   480
      TabIndex        =   2
      Top             =   600
      Visible         =   0   'False
      Width           =   5235
      Begin Project1.AOLCmd AOLCmd3 
         Height          =   375
         Left            =   1920
         TabIndex        =   12
         Top             =   240
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         Caption         =   "Fade Out-In"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "Form1.frx":0038
         ForeColor       =   0
         BackColor       =   -2147483624
         StandardColors  =   0   'False
         BackColorClick  =   12648447
      End
      Begin Project1.AOLCmd AOLCmd2 
         Height          =   375
         Left            =   3600
         TabIndex        =   11
         Top             =   240
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         Caption         =   "Undo Transperent"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "Form1.frx":0054
         ForeColor       =   0
         BackColor       =   -2147483624
         StandardColors  =   0   'False
         BackColorClick  =   12648447
      End
      Begin Project1.AOLCmd AOLCmd1 
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         Caption         =   "Make Transperent"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "Form1.frx":0070
         ForeColor       =   0
         BackColor       =   -2147483624
         StandardColors  =   0   'False
         BackColorClick  =   12648447
      End
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   1560
      Top             =   4320
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H80000018&
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Text            =   "255"
      Top             =   5640
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   600
      Top             =   4320
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000018&
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   5280
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1080
      Top             =   4320
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000018&
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   6
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000018&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   5
      Top             =   360
      Width           =   135
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000018&
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   4
      Top             =   360
      Width           =   135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i
Dim started
Dim com
Dim transp

Private Sub AOLCmd1_Click()
Command1_Click
End Sub

Private Sub AOLCmd1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
AOLCmd1_Click
End If
End Sub

Private Sub AOLCmd2_Click()
Command2_Click
End Sub

Private Sub AOLCmd2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
AOLCmd2_Click
End If
End Sub

Private Sub AOLCmd3_Click()
Command3_Click
End Sub

Private Sub AOLCmd3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
AOLCmd3_Click
End If
End Sub

Private Sub AOLCmd4_Click()
    Dim Answer As String
    'Tel the client to press some keys on the keyboard
    'This function I had the most fun with... Wait when the other guy
    'works on a document to have the most fun.
    Answer = InputBox("How Transparent? (1-255)", "Transperent", "")
    Text2 = Answer
'Timer1.Enabled = False
On Error Resume Next
SetLayered Text4.Text, True, Text2.Text
End Sub

Private Sub AOLCmd5_Click()
SetLayered Text4.Text, False, Text2.Text
End Sub

Private Sub Combo1_Click()
List4.ListIndex = Combo1.ListIndex
Text4.Text = List4.Text
End Sub

Private Sub Command1_Click()
    Dim Answer As String
    'Tel the client to press some keys on the keyboard
    'This function I had the most fun with... Wait when the other guy
    'works on a document to have the most fun.
    Answer = InputBox("How Transparent? (1-255)", "Transperent", "")
    Text2 = Answer
Text4.Text = Combo1.Text
'Timer1.Enabled = False
On Error Resume Next
SetLayered Text1.Text, True, Text2.Text
End Sub

Private Sub Command2_Click()
On Error Resume Next
SetLayered Text1.Text, False, Text2.Text
End Sub

Private Sub Command3_Click()
Timer2.Enabled = True
End Sub

Private Sub Form_Load()
AlwaysOnTop Form1, True
Label3.Caption = Me.Caption
SetLayered Me.hwnd, True, 200
On Error Resume Next
    Timer1.Enabled = True
    Timer1.Interval = 100
'SetLayered 264160, True, 200
helpme
Dim a
For a = 0 To List1.ListCount - 1
On Error Resume Next
    Combo1.AddItem List1.List(a)
On Error Resume Next
Next
    Dim Answer As String
    'Tel the client to press some keys on the keyboard
    'This function I had the most fun with... Wait when the other guy
    'works on a document to have the most fun.
    Answer = InputBox("Whitch Mode (ComboBox or Mouse)", "Mode", "")
    If Answer = "ComboBox" Then
    Frame2.Visible = True
    AOLCmd4.Visible = True
    AOLCmd5.Visible = True
    Combo1.Visible = True
    End If
    If Answer = "Mouse" Then
    Frame1.Visible = True
    AOLCmd1.Visible = True
    AOLCmd2.Visible = True
    AOLCmd3.Visible = True
    End If
    If Answer = "" Then
    End
    End If

End Sub

Private Sub Label1_Click()
Me.WindowState = vbMinimized
End Sub

Private Sub Label2_Click()
End
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Select Case Button
Case "1"
FormDrag Me
Case "2"
End Select
End Sub

Private Sub List1_Click()
'List4.ListIndex = List1.ListIndex
End Sub

Private Sub List4_Click()

End Sub

Private Sub Timer1_Timer()

    Dim AnB As Where
    Call GetCursorPos(AnB)
    On Error Resume Next
    YourHWND& = WindowFromPoint(AnB.Pointa, AnB.Pointb) 'Mouse pos.


    If YourHWND& <> LasthWnd& Then 'If there no the same
        LasthWnd& = YourHWND&
        Text1 = YourHWND& 'place whatever output device here To equal hwndover.
    End If

End Sub
Public Sub AlwaysOnTop(myfrm As Form, SetOnTop As Boolean)


    If SetOnTop Then
        lFlag = HWND_TOPMOST
    Else
        lFlag = HWND_NOTOPMOST
    End If
    SetWindowPos Me.hwnd, lFlag, _
    Me.Left / Screen.TwipsPerPixelX, _
    Me.Top / Screen.TwipsPerPixelY, _
    Me.Width / Screen.TwipsPerPixelX, _
    Me.Height / Screen.TwipsPerPixelY, _
    SWP_NOACTIVATE Or SWP_SHOWWINDOW
End Sub

Private Sub Timer2_Timer()
If Text3.Text = "255" Then
started = Text1.Text
End If
Text3.Text = Text3.Text - 1
On Error GoTo Timer3:
SetLayered started, True, Text3.Text
Exit Sub
Timer3:
Timer2.Enabled = False
Timer3.Enabled = True
Timer2.Enabled = False
End Sub

Private Sub Timer3_Timer()
Text3.Text = Text3.Text + 1
On Error GoTo error:
SetLayered started, True, Text3.Text
Exit Sub
error:
Timer3.Enabled = False
Text3.Text = "255"
On Error Resume Next
SetLayered started, False, 255
Timer3.Enabled = False
End Sub
Public Function helpme()
start = 1
paren = 1
retval = EnumWindows(AddressOf EnumChildProc, 0)
onlycaps = 0

'add top level windows first
For i = 0 To List1.ListCount - 1
    tv.Nodes.Add , , "w" & Mid(List1.List(i), 1, InStr(1, List1.List(i), " ") - 1), List1.List(i)
Next i

For i = 1 To List1.ListCount - 1
    List2.AddItem Mid(List1.List(i), 1, InStr(1, List1.List(i), " ") - 1)   'move list1 hwnds to list2 for reference
Next i
List1.Clear

'add child windows
For i = 0 To List2.ListCount - 1
    List1.Clear
    paren = 0
    phwnd = List2.List(i)
    retval = EnumChildWindows(phwnd, AddressOf EnumChildProc, 0)
    If List1.ListCount <> 0 Then
        ListToTree phwnd
    End If
Next i

End Function
Sub ListToTree(ph)
For i = 0 To List1.ListCount - 1
    hh = Mid(List1.List(i), 1, InStr(1, List1.List(i), " ") - 1) 'handle of window
    On Error Resume Next
    'List3.AddItem "w" & GetParent(hh), tvwChild, "w" & hh, List1.List(i)
Next i
End Sub


