Attribute VB_Name = "Rayaneh"

Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Public Declare Function GetDesktopWindow Lib "user32" () As Long

Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long

Public Const SRCCOPY = &HCC0020

Public Const SRCAND = &H8800C6

Public Const SRCINVERT = &H660046

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Type PAINTSTRUCT
        hdc As Long
        fErase As Long
        rcPaint As RECT
        fRestore As Long
        fIncUpdate As Long
        rgbReserved(32) As Byte
End Type

Public Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function PaintDesktop Lib "user32" (ByVal hdc As Long) As Long

Public Sub center(obj As Object, sx, sy, swidth, sheight)
 cx = sx + swidth / 2
 cy = sy + sheight / 2
 cw = obj.Width / 2
 ch = obj.Height / 2
 X = cx - cw
 Y = cy - ch
 obj.Left = X
 obj.Top = Y
End Sub
