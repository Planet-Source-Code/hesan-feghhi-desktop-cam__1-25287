VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1320
   ClientLeft      =   3885
   ClientTop       =   2355
   ClientWidth     =   2400
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00E0E0E0&
   Icon            =   "Cam2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1320
   ScaleWidth      =   2400
   Begin VB.TextBox T 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Image1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "S H O W"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   840
      Left            =   120
      TabIndex        =   2
      Top             =   420
      Width           =   255
   End
   Begin VB.Shape Shape4 
      FillColor       =   &H000000C0&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   2160
      Shape           =   3  'Circle
      Top             =   1080
      Width           =   135
   End
   Begin VB.Line Line1 
      X1              =   480
      X2              =   480
      Y1              =   0
      Y2              =   1320
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  HN"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00909090&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   1200
      Top             =   120
      Width           =   375
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000C0&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   840
      Shape           =   3  'Circle
      Top             =   480
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   1800
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_KeyPress(KeyAscii As Integer)
 If KeyAscii = 27 Then End
End Sub

Private Sub Form_Load()
 Form2.Left = Me.Left + 250
 Form2.Top = Me.Top - 90
End Sub

Private Sub Image1_Click()
 'a = Val(VBA.GetSetting("Hello", "H", "K"))
 'If Val(T.Text) = a Then T.Text = Str$(Val(T.Text) - 1)
 On Error GoTo Prev
 Form3.Picture = LoadPicture(App.Path + "\HN" + Str$(Val(T.Text)) + ".BMP")
 Form3.Show
 Me.Hide
 Form2.Hide
 Exit Sub

Prev:
  a = MsgBox("Picture isn't gotten. Do you want to see the previous picture?", vbYesNo, "Show")
  If a = vbYes Then
   Form3.Picture = LoadPicture(App.Path + "\HN" + Str$(Val(T.Text) - 1) + ".BMP")
   Form3.Show
   Me.Hide
   Form2.Hide
  End If
End Sub

Private Sub T_Change()
 a = Val(VBA.GetSetting("Hello", "H", "K"))
 If Val(T.Text) > a Then T.Text = Str$(Val(T.Text) - 1)
 If Val(T.Text) < 0 Then T.Text = Str$(0)
End Sub

Private Sub T_KeyPress(KeyAscii As Integer)
 If KeyAscii = 27 Then End
End Sub
