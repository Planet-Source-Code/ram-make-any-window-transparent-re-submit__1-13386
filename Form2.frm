VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Dim lReigon As Long
    Dim lReigon1 As Long
    Dim lReigon2 As Long
    Dim lResult As Long
    
        Unload Form1
        Load Form1
        
        ' Make the form big enough
        Form1.Width = 500 * Screen.TwipsPerPixelX
        Form1.Height = 300 * Screen.TwipsPerPixelY
        
        lReigon = CreateEllipticRgn(0, 0, 420, 200)
        
        ' set the window reigon
        lResult = SetWindowRgn(Form1.hWnd, lReigon, True)
        Form1.Show
    Unload Me

End Sub
