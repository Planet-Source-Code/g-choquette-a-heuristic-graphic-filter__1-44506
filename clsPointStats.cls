VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPointStats"
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

Private mbolIsDirty As Boolean
Private mdblAverage As Double
Private mdblStandardDeviation As Double
Private mdblSum As Double
Private mlngNumberOfSamples As Long
Private msngLowerBound As Single
Private msngStandardDeviationOrders As Single
Private msngUpperBound As Single
Private msngValues() As Single

Public Function AddPointToSample(ByVal vData)
  Dim lngIndex As Long
  Dim dblSumError As Double
  mlngNumberOfSamples = mlngNumberOfSamples + 1
  If mlngNumberOfSamples > UBound(msngValues) Then
    ReDim Preserve msngValues(1 To mlngNumberOfSamples + 10)
  End If
  msngValues(mlngNumberOfSamples) = vData
  mdblSum = mdblSum + vData
  If Not mbolIsDirty Then
    mbolIsDirty = True
  End If
End Function

Public Property Get Average() As Double
  CalcUpdate
  Average = mdblAverage
End Property

Private Sub CalcBounds()
  msngUpperBound = mdblAverage + mdblStandardDeviation * msngStandardDeviationOrders
  msngLowerBound = mdblAverage - mdblStandardDeviation * msngStandardDeviationOrders
End Sub

Private Sub CalcUpdate()
  Dim lngIndex As Long
  Dim dblSumError As Double
  If mbolIsDirty Then
    mdblAverage = mdblSum / mlngNumberOfSamples
    For lngIndex = 1 To mlngNumberOfSamples
      dblSumError = dblSumError + (msngValues(lngIndex) - mdblAverage) ^ 2
    Next lngIndex
    If mlngNumberOfSamples > 1 Then
      mdblStandardDeviation = (dblSumError / (mlngNumberOfSamples - 1)) ^ 0.5
    Else
      mdblStandardDeviation = 0
    End If
    CalcBounds
    mbolIsDirty = False
  End If
End Sub

Public Function ClearStats()
  ReDim msngValues(1 To 5)
  mlngNumberOfSamples = 0
  mdblSum = 0
  mdblAverage = 0
  msngLowerBound = 0
  msngUpperBound = 0
  mbolIsDirty = False
End Function

Public Function IsWithinBounds(ByVal TestValue As Single) As Boolean
  CalcUpdate
  If msngLowerBound < TestValue Then
    If msngUpperBound > TestValue Then
      IsWithinBounds = True
    End If
  End If
End Function

Public Property Get LowerBound() As Single
  CalcUpdate
  LowerBound = msngLowerBound
End Property

Public Property Get NumberOfSamples() As Long
  NumberOfSamples = mlngNumberOfSamples
End Property

Public Property Get StandardDeviation() As Double
  CalcUpdate
  StandardDeviation = mdblStandardDeviation
End Property

Public Property Get StandardDeviationOrders() As Single
  StandardDeviationOrders = msngStandardDeviationOrders
End Property

Public Property Let StandardDeviationOrders(ByVal vData As Single)
  If Not msngStandardDeviationOrders = vData Then
    msngStandardDeviationOrders = vData
    If mbolIsDirty Then
      CalcUpdate
    Else
      CalcBounds
    End If
  End If
End Property

Public Property Get Sum() As Double
  Sum = mdblSum
End Property

Public Property Get UpperBound() As Single
  CalcUpdate
  UpperBound = msngUpperBound
End Property

Public Property Get Values(Index As Long) As Single
  Values = msngValues(Index)
End Property

Private Sub Class_Initialize()
  ReDim msngValues(1 To 10)
  msngStandardDeviationOrders = 1
End Sub
