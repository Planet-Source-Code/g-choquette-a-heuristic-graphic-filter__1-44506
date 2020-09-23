VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsChartElements"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"clsChartElement"
Attribute VB_Ext_KEY = "Member0" ,"clsChartElement"
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

Private mCol As Collection

Public Function Add(ChartElement_ID As Long, ChartElementName As String, Description As String, OverwriteColor As Long) As clsChartElement
  Dim objNewMember As clsChartElement
  Set objNewMember = New clsChartElement
  objNewMember.ChartElement_ID = ChartElement_ID
  objNewMember.ChartElementName = ChartElementName
  objNewMember.Description = Description
  objNewMember.OverwriteColor = OverwriteColor
  mCol.Add objNewMember, ChartElementName
  Set Add = objNewMember
  Set objNewMember = Nothing
End Function

Public Property Get Item(vntIndexKey As Variant) As clsChartElement
Attribute Item.VB_UserMemId = 0
  Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
  Count = mCol.Count
End Property

Public Function Clear()
  Set mCol = Nothing
  Set mCol = New Collection
End Function

Public Function LoadFromDatabase(Connect As ADODB.Connection)
  Dim rsChartElement As ADODB.Recordset
  Dim strQry As String
  Dim cheChartElement As clsChartElement
  Set rsChartElement = New ADODB.Recordset
  strQry = "SELECT * FROM ChartElement"
  rsChartElement.Open strQry, Connect, adOpenKeyset, adLockOptimistic
  If mCol.Count > 0 Then
    Clear
  End If
  Do While Not rsChartElement.EOF
    Set cheChartElement = Add(rsChartElement!ChartElement_ID, rsChartElement!ChartElementName, rsChartElement!Description, rsChartElement!OverwriteColor)
    cheChartElement.LoadFromDatabase Connect
    rsChartElement.MoveNext
  Loop
  rsChartElement.Close
  Set rsChartElement = Nothing
End Function

Public Sub Remove(vntIndexKey As Variant)
  mCol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
  Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
  Set mCol = Nothing
End Sub

