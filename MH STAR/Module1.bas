Attribute VB_Name = "Module1"


Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type



Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long

Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long

Public Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Type Star
x As Single
y As Single
size As Single
img As Integer
End Type

