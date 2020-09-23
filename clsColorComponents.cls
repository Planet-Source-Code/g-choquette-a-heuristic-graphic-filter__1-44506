VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsColorComponents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'***************************************************************************
' Heuristic graphic filter
' Written by G Choquette 4/4/2003,  gchoquet@hotmail.com
' Copyright 2003
'
' Provided as is for non-commercial, educational purposes.  All commercial
' rights are reserved by the author.  The use of any of this code in any
' commercial applications requires the expressed written consent of the author.
'
'***************************************************************************
Option Explicit

Private mintAverage As Integer
Private mintBlue As Integer
Private mintGreen As Integer
Private mintRed As Integer
Private mintSum As Integer
Private msngBG As Single
Private msngBlueFraction As Single
Private msngBR As Single
Private msngGB As Single
Private msngGR As Single
Private msngGreenFraction As Single
Private msngRB As Single
Private msngRedFraction As Single
Private msngRG As Single
Private msngValues() As Single

Public Event ColorUpdated()

Private Const MAX_RATIO_VALUE = 200

Public Property Get Average() As Integer
  Average = mintAverage
End Property

Public Property Get BG() As Single
  BG = msngBG
End Property

Public Property Get Blue() As Integer
  Blue = mintBlue
End Property

Public Property Let Blue(ByVal vData As Integer)
  If Not mintBlue = vData Then
    mintBlue = vData
    CalcUpdate
    RaiseEvent ColorUpdated
  End If
End Property

Public Property Get BlueFraction() As Single
  BlueFraction = msngBlueFraction
End Property

Public Property Get BR() As Single
  BR = msngBR
End Property

Private Sub CalcUpdate()
  mintSum = mintBlue + mintGreen + mintRed
  mintAverage = CInt(mintSum / 3)
  If mintSum > 0 Then
    msngRedFraction = mintRed / mintSum
    msngGreenFraction = mintGreen / mintSum
    msngBlueFraction = mintBlue / mintSum
  End If
  If mintRed = 0 Then
    msngGR = MAX_RATIO_VALUE
    msngBR = MAX_RATIO_VALUE
  Else
    msngGR = mintGreen / mintRed
    msngBR = mintBlue / mintRed
  End If
  If mintGreen = 0 Then
    msngRG = MAX_RATIO_VALUE
    msngBG = MAX_RATIO_VALUE
  Else
    msngRG = mintRed / mintGreen
    msngBG = mintBlue / mintGreen
  End If
  If mintBlue = 0 Then
    msngRB = MAX_RATIO_VALUE
    msngGB = MAX_RATIO_VALUE
  Else
    msngRB = mintRed / mintBlue
    msngGB = mintGreen / mintBlue
  End If
End Sub

Public Property Get GB() As Single
  GB = msngGB
End Property

Public Property Get GR() As Single
  GR = msngGR
End Property

Public Property Get Green() As Integer
  Green = mintGreen
End Property

Public Property Let Green(ByVal vData As Integer)
  If Not mintGreen = vData Then
    mintGreen = vData
    CalcUpdate
    RaiseEvent ColorUpdated
  End If
End Property

Public Property Get GreenFraction() As Single
  GreenFraction = msngGreenFraction
End Property

Public Property Get RB() As Single
  RB = msngRB
End Property

Public Property Get Red() As Integer
  Red = mintRed
End Property

Public Property Let Red(ByVal vData As Integer)
  If Not mintRed = vData Then
    mintRed = vData
    CalcUpdate
    RaiseEvent ColorUpdated
  End If
End Property

Public Property Get RedFraction() As Single
  RedFraction = msngRedFraction
End Property

Public Property Get RG() As Single
  RG = msngRG
End Property

Public Function SetColorComponent(Color As Long)
  mintBlue = Color \ 256 \ 256
  mintGreen = (Color - mintBlue * 65536) \ 256
  mintRed = Color And 255
  CalcUpdate
  RaiseEvent ColorUpdated
End Function

Public Property Get Sum() As Integer
  Sum = mintSum
End Property
