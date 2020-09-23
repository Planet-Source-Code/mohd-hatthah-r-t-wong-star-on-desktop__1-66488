VERSION 5.00
Begin VB.Form frmmh 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000040&
   BorderStyle     =   0  'None
   ClientHeight    =   8235
   ClientLeft      =   3015
   ClientTop       =   1920
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   549
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   439
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   40
      Left            =   480
      Top             =   6120
   End
   Begin VB.Image Im 
      Height          =   495
      Index           =   4
      Left            =   0
      Picture         =   "frmmh.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Im 
      Height          =   495
      Index           =   3
      Left            =   0
      Picture         =   "frmmh.frx":F346
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Im 
      Height          =   495
      Index           =   2
      Left            =   0
      Picture         =   "frmmh.frx":1E68C
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Im 
      Height          =   495
      Index           =   1
      Left            =   0
      Picture         =   "frmmh.frx":2D9D2
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Im 
      Height          =   495
      Index           =   0
      Left            =   0
      Picture         =   "frmmh.frx":3CD18
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image mask 
      Height          =   495
      Left            =   600
      Picture         =   "frmmh.frx":4C05E
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "frmmh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim s(50) As Star

Private Sub Form_Activate()
Dim dc As Long, r As RECT
GetClientRect hwnd, r
BitBlt hdc, 0, 0, r.Right, r.Bottom, GetDC(0), 0, 0, vbSrcCopy
Picture = Image
Timer1.Enabled = True
End Sub

Private Sub Form_Click()
End
End Sub

Private Sub Form_Resize()
Dim i As Integer, k As Integer
Timer1.Enabled = False
For i = 0 To 50
s(i).x = Rnd * Me.ScaleWidth
s(i).y = Rnd * Me.ScaleHeight
s(i).size = 5 + (Rnd * 20)
Next


End Sub

Private Sub Timer1_Timer()
Dim i As Integer, j As Integer
Cls
For i = 0 To 50
If s(i).y > Me.ScaleHeight Then s(i).y = 0
j = Rnd * 5
If j > 4 Then j = 4
Me.PaintPicture mask, s(i).x, s(i).y, s(i).size, s(i).size, , , , , vbSrcAnd
Me.PaintPicture Im(j), s(i).x, s(i).y, s(i).size, s(i).size, , , , , vbSrcPaint
s(i).y = s(i).y + 10
Next
Refresh
End Sub
