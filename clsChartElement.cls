VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsChartElement"
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
' Requires MS ADO 2.6 Library or compatable data access components
'***************************************************************************
Option Explicit

Private mclcColorComponentTest As clsColorComponents
Private mlngChartElement_ID As Long
Private mlngOverwriteColor As Long
Private mptsAverage As clsPointStats
Private mptsBG As clsPointStats
Private mptsBlue As clsPointStats
Private mptsBlueFraction As clsPointStats
Private mptsBR As clsPointStats
Private mptsGB As clsPointStats
Private mptsGR As clsPointStats
Private mptsGreen As clsPointStats
Private mptsGreenFraction As clsPointStats
Private mptsRB As clsPointStats
Private mptsRed As clsPointStats
Private mptsRedFraction As clsPointStats
Private mptsRG As clsPointStats
Private mptsSum As clsPointStats
Private mstrChartElementName As String
Private mstrDescription As String
Private WithEvents mclcColorComponents As clsColorComponents
Attribute mclcColorComponents.VB_VarHelpID = -1

Private Const STD_BOUND1 = 1.5
Private Const STD_BOUND2 = 2

Public Property Get Red() As clsPointStats
  Set Red = mptsRed
End Property

Public Property Get Green() As clsPointStats
  Set Green = mptsGreen
End Property

Public Property Get Blue() As clsPointStats
  Set Blue = mptsBlue
End Property

Public Property Get RedFraction() As clsPointStats
  Set RedFraction = mptsRedFraction
End Property

Public Property Get GreenFraction() As clsPointStats
  Set GreenFraction = mptsGreenFraction
End Property

Public Property Get BlueFraction() As clsPointStats
  Set BlueFraction = mptsBlueFraction
End Property

Public Property Get Average() As clsPointStats
  Set Average = mptsAverage
End Property

Public Property Get Sum() As clsPointStats
  Set Sum = mptsSum
End Property

Public Property Get GR() As clsPointStats
  Set GR = mptsGR
End Property

Public Property Get BR() As clsPointStats
  Set BR = mptsBR
End Property

Public Property Get RG() As clsPointStats
  Set RG = mptsRG
End Property

Public Property Get BG() As clsPointStats
  Set BG = mptsBG
End Property

Public Property Get RB() As clsPointStats
  Set RB = mptsRB
End Property

Public Property Get GB() As clsPointStats
  Set GB = mptsGB
End Property

Public Property Get ColorComponents() As clsColorComponents
  Set ColorComponents = mclcColorComponents
End Property

Public Property Let OverwriteColor(ByVal vData As Long)
  mlngOverwriteColor = vData
End Property

Public Property Get OverwriteColor() As Long
  OverwriteColor = mlngOverwriteColor
End Property

Public Property Let Description(ByVal vData As String)
  mstrDescription = vData
End Property

Public Property Get Description() As String
  Description = mstrDescription
End Property

Public Property Let ChartElementName(ByVal vData As String)
  mstrChartElementName = vData
End Property

Public Property Get ChartElementName() As String
  ChartElementName = mstrChartElementName
End Property

Public Property Let ChartElement_ID(ByVal vData As Long)
  mlngChartElement_ID = vData
End Property

Public Property Get ChartElement_ID() As Long
  ChartElement_ID = mlngChartElement_ID
End Property

Public Function ClearStats()
  mptsRed.ClearStats
  mptsGreen.ClearStats
  mptsBlue.ClearStats
  mptsAverage.ClearStats
  mptsSum.ClearStats
  mptsRB.ClearStats
  mptsRG.ClearStats
  mptsGR.ClearStats
  mptsGB.ClearStats
  mptsBR.ClearStats
  mptsBG.ClearStats
  mptsRedFraction.ClearStats
  mptsGreenFraction.ClearStats
  mptsBlueFraction.ClearStats
End Function

Public Function LoadFromDatabase(Connect As ADODB.Connection)
  Dim strQry As String
  Dim rsElementSample As ADODB.Recordset
  If mstrChartElementName = "" Then
    Exit Function
  End If
  ClearStats
  Set rsElementSample = New ADODB.Recordset
  strQry = "SELECT Color FROM ElementSamples WHERE ChartElement_ID = " & mlngChartElement_ID
  rsElementSample.Open strQry, Connect, adOpenKeyset, adLockReadOnly
  Do While Not rsElementSample.EOF
    mclcColorComponents.SetColorComponent rsElementSample!Color
    rsElementSample.MoveNext
  Loop
  rsElementSample.Close
  Set rsElementSample = Nothing
End Function

Public Function IsValidColor(TestPointColor As Long) As Boolean
  mclcColorComponentTest.SetColorComponent TestPointColor
  If Not mptsAverage.IsWithinBounds(CSng(mclcColorComponentTest.Average)) Then
    Exit Function
  ElseIf Not mptsSum.IsWithinBounds(CSng(mclcColorComponentTest.Sum)) Then
    Exit Function
  ElseIf Not mptsRedFraction.IsWithinBounds(mclcColorComponentTest.RedFraction) Then
    Exit Function
  ElseIf Not mptsGreenFraction.IsWithinBounds(mclcColorComponentTest.GreenFraction) Then
    Exit Function
  ElseIf Not mptsBlueFraction.IsWithinBounds(mclcColorComponentTest.BlueFraction) Then
    Exit Function
  ElseIf Not mptsRG.IsWithinBounds(mclcColorComponentTest.RG) Then
    Exit Function
  ElseIf Not mptsRB.IsWithinBounds(mclcColorComponentTest.RB) Then
    Exit Function
  ElseIf Not mptsGR.IsWithinBounds(mclcColorComponentTest.GR) Then
    Exit Function
  ElseIf Not mptsGB.IsWithinBounds(mclcColorComponentTest.GB) Then
    Exit Function
  ElseIf Not mptsBR.IsWithinBounds(mclcColorComponentTest.BR) Then
    Exit Function
  ElseIf Not mptsBG.IsWithinBounds(mclcColorComponentTest.BG) Then
    Exit Function
  ElseIf Not mptsRed.IsWithinBounds(CSng(mclcColorComponentTest.Red)) Then
    Exit Function
  ElseIf Not mptsGreen.IsWithinBounds(CSng(mclcColorComponentTest.Green)) Then
    Exit Function
  ElseIf Not mptsBlue.IsWithinBounds(CSng(mclcColorComponentTest.Blue)) Then
    Exit Function
  End If
  IsValidColor = True
End Function

Public Function SaveToDatabase(Connect As ADODB.Connection)
  Dim strQry As String
  Dim rsElementSample As ADODB.Recordset
  Dim lngIndexMaxRecords As Long
  Dim lngIndex As Long
  Dim lngIndexCount As Long
  If mstrChartElementName = "" Then
    Exit Function
  End If
  Set rsElementSample = New ADODB.Recordset
  strQry = "SELECT * FROM ElementSamples WHERE ChartElement_ID = " & mlngChartElement_ID
  rsElementSample.Open strQry, Connect, adOpenKeyset, adLockOptimistic
  lngIndexMaxRecords = mptsRed.NumberOfSamples
  Do While Not rsElementSample.EOF
    lngIndexCount = lngIndexCount + 1
    If lngIndexCount > lngIndexMaxRecords Then
      'fewer records than exist in the database, delete extra records
      rsElementSample.Delete
    Else
      'replace color data in db w/ recorded values
      rsElementSample!Color = RGB(mptsRed.Values(lngIndexCount), mptsGreen.Values(lngIndexCount), mptsBlue.Values(lngIndexCount))
    End If
    rsElementSample.Update
    rsElementSample.MoveNext
  Loop
  'now check to see if additional records were added that need to be saved
  For lngIndex = lngIndexCount + 1 To lngIndexMaxRecords
    rsElementSample.AddNew
    rsElementSample!SampleIndex = lngIndex
    rsElementSample!ChartElement_ID = mlngChartElement_ID
    rsElementSample!Color = RGB(mptsRed.Values(lngIndex), mptsGreen.Values(lngIndex), mptsBlue.Values(lngIndex))
    rsElementSample.Update
  Next lngIndex
  rsElementSample.Close
  Set rsElementSample = Nothing
End Function

Private Sub Class_Initialize()
  Set mclcColorComponentTest = New clsColorComponents
  Set mclcColorComponents = New clsColorComponents
  Set mptsRed = New clsPointStats
  mptsRed.StandardDeviationOrders = STD_BOUND2
  Set mptsGreen = New clsPointStats
  mptsGreen.StandardDeviationOrders = STD_BOUND2
  Set mptsBlue = New clsPointStats
  mptsBlue.StandardDeviationOrders = STD_BOUND2
  Set mptsRedFraction = New clsPointStats
  mptsRedFraction.StandardDeviationOrders = STD_BOUND2
  Set mptsGreenFraction = New clsPointStats
  mptsGreenFraction.StandardDeviationOrders = STD_BOUND2
  Set mptsBlueFraction = New clsPointStats
  mptsBlueFraction.StandardDeviationOrders = STD_BOUND2
  Set mptsAverage = New clsPointStats
  mptsAverage.StandardDeviationOrders = STD_BOUND1
  Set mptsSum = New clsPointStats
  mptsSum.StandardDeviationOrders = STD_BOUND1
  Set mptsRG = New clsPointStats
  mptsRG.StandardDeviationOrders = STD_BOUND2
  Set mptsRB = New clsPointStats
  mptsRB.StandardDeviationOrders = STD_BOUND2
  Set mptsGR = New clsPointStats
  mptsGR.StandardDeviationOrders = STD_BOUND2
  Set mptsGB = New clsPointStats
  mptsGB.StandardDeviationOrders = STD_BOUND2
  Set mptsBR = New clsPointStats
  mptsBR.StandardDeviationOrders = STD_BOUND2
  Set mptsBG = New clsPointStats
  mptsBG.StandardDeviationOrders = STD_BOUND2
End Sub

Private Sub Class_Terminate()
  Set mclcColorComponentTest = Nothing
  Set mclcColorComponents = Nothing
  Set mptsRed = Nothing
  Set mptsGreen = Nothing
  Set mptsBlue = Nothing
  Set mptsRedFraction = Nothing
  Set mptsGreenFraction = Nothing
  Set mptsBlueFraction = Nothing
  Set mptsAverage = Nothing
  Set mptsSum = Nothing
  Set mptsRG = Nothing
  Set mptsRB = Nothing
  Set mptsGR = Nothing
  Set mptsGB = Nothing
  Set mptsBR = Nothing
  Set mptsBG = Nothing
End Sub

Private Sub mclcColorComponents_ColorUpdated()
  With mclcColorComponents
    mptsAverage.AddPointToSample .Average
    mptsSum.AddPointToSample .Sum
    mptsRed.AddPointToSample .Red
    mptsGreen.AddPointToSample .Green
    mptsBlue.AddPointToSample .Blue
    mptsRedFraction.AddPointToSample .RedFraction
    mptsGreenFraction.AddPointToSample .GreenFraction
    mptsBlueFraction.AddPointToSample .BlueFraction
    mptsGR.AddPointToSample .GR
    mptsBR.AddPointToSample .BR
    mptsRG.AddPointToSample .RG
    mptsBG.AddPointToSample .BG
    mptsRB.AddPointToSample .RB
    mptsGB.AddPointToSample .GB
  End With
End Sub
