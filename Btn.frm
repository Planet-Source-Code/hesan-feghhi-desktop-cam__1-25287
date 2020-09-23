VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   75
   ClientLeft      =   3765
   ClientTop       =   1665
   ClientWidth     =   255
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   75
   ScaleWidth      =   255
   ShowInTaskbar   =   0   'False
   Begin VB.Image Command1 
      Height          =   135
      Left            =   0
      Top             =   0
      Width           =   255
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   0
      Top             =   0
      Width           =   255
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Me.Top = Me.Top - 70
 w = 255
 For i = 0 To 255
  Form1.Shape4.FillColor = RGB(w, i, 0)
  w = w - 1
  DoEvents
 Next i
 For i = 1 To 1000
  DoEvents
 Next i
 Form1.Visible = False
 Visible = False
 DoEvents
 Dim hwnd&, hdc&
 hwnd& = GetDesktopWindow()
 hdc& = GetDC(hwnd&)
 BitBlt Form3.hdc, 0, 0, Screen.Width / 15, Screen.Height / 15, hdc&, 0, 0, SRCCOPY
 DoEvents
 curnum = Val(Form1.T.Text)
 SavePicture Form3.Image, App.Path + "\HN" + Str$(curnum) + ".BMP"
 Unload Form3
 'Form3.Show
 'DoEvents
 'BitBlt Form3.hdc, 0, 0, Screen.Width / 15, Screen.Height / 15, hdc&, 0, 0, SRCCOPY
 'DoEvents
 curnum = curnum + 1
 VBA.SaveSetting "Hello", "H", "K", Str$(curnum)
 Form1.T.Text = VBA.GetSetting("Hello", "H", "K")
 Form1.Visible = True
 Visible = True
 DoEvents
End Sub

Private Sub Command1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Me.Top = Me.Top + 70
End Sub

Private Sub Form_Load()
 Form1.Show , Me
 Form1.T.Text = VBA.GetSetting("Hello", "H", "K")
 If Form1.T.Text = "" Then Form1.T.Text = "0"
End Sub
