VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form3 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   ClientHeight    =   4725
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6510
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   ScaleHeight     =   4725
   ScaleWidth      =   6510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Prev 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   1920
      ScaleHeight     =   1665
      ScaleWidth      =   2505
      TabIndex        =   0
      Top             =   1560
      Visible         =   0   'False
      Width           =   2535
   End
   Begin MSComDlg.CommonDialog CDlg 
      Left            =   3000
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Save Capture As ..."
      Filter          =   "Bitmap Files (*.bmp)|*.bmp|Jpeg Files (*.jpg;*.jpe)|*.Jpg;*.Jpe|GIF Files (*.gif)|*.gif|All Files (*.*)|*.*"
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Drag the mouse to choose the area of the screen you want to capture."
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6495
   End
   Begin VB.Shape Selection 
      Height          =   1935
      Left            =   1800
      Top             =   1440
      Visible         =   0   'False
      Width           =   2775
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private XOrigin, YOrigin, disabled

Private Sub Command1_Click()
 Unload Me
End Sub

Private Sub Form_Load()
 Label1.Width = Screen.Width
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Not disabled Then
 Selection.Visible = True
 XOrigin = X
 YOrigin = Y
 Selection.Left = X
 Selection.Top = Y
 Selection.Width = 0
 Selection.Height = 0
 End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Not disabled Then
 If X >= XOrigin Then
  Selection.Width = X - XOrigin
 Else
  Selection.Left = X
  Selection.Width = Abs(XOrigin - X)
 End If
 If Y >= XOrigin Then
  Selection.Height = Y - YOrigin
 Else
  Selection.Top = Y
  Selection.Height = Abs(YOrigin - Y)
 End If
 End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Not disabled Then
 Selection.Visible = False
 Prev.Top = Selection.Top
 Prev.Left = Selection.Left
 Prev.Width = Selection.Width
 Prev.Height = Selection.Height
 Prev.Visible = True
 BitBlt Prev.hdc, 0, 0, Prev.Width / 15, Prev.Height / 15, Me.hdc, Prev.Left / 15, Prev.Top / 15, SRCCOPY
 Me.Picture = LoadPicture("")
 Me.Cls
 center Prev, 0, 0, Screen.Width, Screen.Height
 CDlg.ShowSave
 If Len(CDlg.FileName) > 0 Then SavePicture Prev.Image, CDlg.FileName
 disabled = True
 End If
 Unload Me
 Form1.Show
 Form2.Show
End Sub
