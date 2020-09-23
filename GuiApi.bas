Attribute VB_Name = "GuiApi"
Option Explicit

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, _
  ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, _
  ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal xD As Long, ByVal yD As Long, _
  ByVal nWidthD As Long, ByVal nHeightD As Long, ByVal hSrcDC As Long, ByVal xS As Long, ByVal yS As Long, _
  ByVal nWidthS As Long, ByVal nHeightS As Long, ByVal dwRop As Long) As Long
Public Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, _
  ByVal crColor As Long) As Long
Public Declare Function InvertRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long

