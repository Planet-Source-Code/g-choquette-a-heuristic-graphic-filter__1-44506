VERSION 5.00
Begin VB.Form frmChartProcess 
   Caption         =   "Chart Filter"
   ClientHeight    =   14430
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   19080
   LinkTopic       =   "Form1"
   ScaleHeight     =   14430
   ScaleWidth      =   19080
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSetFilteredAsBase 
      Caption         =   "Set Filtered as Base"
      Height          =   375
      Left            =   11640
      TabIndex        =   73
      Top             =   11040
      Width           =   2655
   End
   Begin VB.CommandButton cmdRestoreElement 
      Caption         =   "Restore Element"
      Height          =   375
      Left            =   8640
      TabIndex        =   72
      Top             =   11040
      Width           =   2655
   End
   Begin VB.CommandButton cmdSaveCurrentElement 
      Caption         =   "Save Element Data"
      Height          =   375
      Left            =   8640
      TabIndex        =   71
      Top             =   10080
      Width           =   2655
   End
   Begin VB.ComboBox cmbChartElement 
      Height          =   315
      Left            =   8640
      TabIndex        =   69
      Text            =   "Combo1"
      Top             =   9600
      Width           =   2655
   End
   Begin VB.CommandButton cmdSubFromOrigImage 
      Caption         =   "Subtract from Original"
      Height          =   375
      Left            =   11640
      TabIndex        =   68
      Top             =   10560
      Width           =   2655
   End
   Begin VB.CommandButton cmdRestoreImage 
      Caption         =   "Restore Original Image"
      Height          =   375
      Left            =   11640
      TabIndex        =   67
      Top             =   9600
      Width           =   2655
   End
   Begin VB.CommandButton cmdClearStats 
      Caption         =   "Clear Element Data"
      Height          =   375
      Left            =   8640
      TabIndex        =   34
      Top             =   10560
      Width           =   2655
   End
   Begin VB.CommandButton cmdPerformFilter 
      Caption         =   "Perform Filter"
      Height          =   375
      Left            =   11640
      TabIndex        =   33
      Top             =   10080
      Width           =   2655
   End
   Begin VB.PictureBox picColorUnderMouse 
      Height          =   255
      Left            =   5640
      ScaleHeight     =   195
      ScaleWidth      =   555
      TabIndex        =   3
      Top             =   9480
      Width           =   615
   End
   Begin VB.PictureBox picZoom 
      AutoRedraw      =   -1  'True
      Height          =   5055
      Left            =   0
      ScaleHeight     =   333
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   349
      TabIndex        =   2
      Top             =   9360
      Width           =   5295
   End
   Begin VB.PictureBox picFiltered 
      AutoRedraw      =   -1  'True
      Height          =   9345
      Left            =   9540
      ScaleHeight     =   619
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   632
      TabIndex        =   1
      Top             =   0
      Width           =   9540
   End
   Begin VB.PictureBox picOrigChart 
      AutoRedraw      =   -1  'True
      Height          =   9345
      Left            =   0
      ScaleHeight     =   619
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   632
      TabIndex        =   0
      Top             =   0
      Width           =   9540
   End
   Begin VB.Label lbl 
      Caption         =   "Chart Element"
      Height          =   255
      Index           =   13
      Left            =   8640
      TabIndex        =   70
      Top             =   9360
      Width           =   1575
   End
   Begin VB.Label lblGreen 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Index           =   2
      Left            =   7680
      TabIndex        =   66
      Top             =   10320
      Width           =   735
   End
   Begin VB.Label lblBlue 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Index           =   2
      Left            =   7680
      TabIndex        =   65
      Top             =   10560
      Width           =   735
   End
   Begin VB.Label lblAvg 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Index           =   2
      Left            =   7680
      TabIndex        =   64
      Top             =   10800
      Width           =   735
   End
   Begin VB.Label lblSum 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Index           =   2
      Left            =   7680
      TabIndex        =   63
      Top             =   11040
      Width           =   735
   End
   Begin VB.Label lblGR 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Index           =   2
      Left            =   7680
      TabIndex        =   62
      Top             =   11280
      Width           =   735
   End
   Begin VB.Label lblBR 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Index           =   2
      Left            =   7680
      TabIndex        =   61
      Top             =   11520
      Width           =   735
   End
   Begin VB.Label lblRG 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Index           =   2
      Left            =   7680
      TabIndex        =   60
      Top             =   11760
      Width           =   735
   End
   Begin VB.Label lblBG 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Index           =   2
      Left            =   7680
      TabIndex        =   59
      Top             =   12000
      Width           =   735
   End
   Begin VB.Label lblRB 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Index           =   2
      Left            =   7680
      TabIndex        =   58
      Top             =   12240
      Width           =   735
   End
   Begin VB.Label lblGB 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Index           =   2
      Left            =   7680
      TabIndex        =   57
      Top             =   12480
      Width           =   735
   End
   Begin VB.Label lblBlueF 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Index           =   2
      Left            =   7680
      TabIndex        =   56
      Top             =   13200
      Width           =   735
   End
   Begin VB.Label lblGreenF 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Index           =   2
      Left            =   7680
      TabIndex        =   55
      Top             =   12960
      Width           =   735
   End
   Begin VB.Label lblRedF 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Index           =   2
      Left            =   7680
      TabIndex        =   54
      Top             =   12720
      Width           =   735
   End
   Begin VB.Label lblGreen 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Index           =   1
      Left            =   6840
      TabIndex        =   53
      Top             =   10320
      Width           =   735
   End
   Begin VB.Label lblBlue 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Index           =   1
      Left            =   6840
      TabIndex        =   52
      Top             =   10560
      Width           =   735
   End
   Begin VB.Label lblAvg 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Index           =   1
      Left            =   6840
      TabIndex        =   51
      Top             =   10800
      Width           =   735
   End
   Begin VB.Label lblSum 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Index           =   1
      Left            =   6840
      TabIndex        =   50
      Top             =   11040
      Width           =   735
   End
   Begin VB.Label lblGR 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Index           =   1
      Left            =   6840
      TabIndex        =   49
      Top             =   11280
      Width           =   735
   End
   Begin VB.Label lblBR 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Index           =   1
      Left            =   6840
      TabIndex        =   48
      Top             =   11520
      Width           =   735
   End
   Begin VB.Label lblRG 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Index           =   1
      Left            =   6840
      TabIndex        =   47
      Top             =   11760
      Width           =   735
   End
   Begin VB.Label lblBG 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Index           =   1
      Left            =   6840
      TabIndex        =   46
      Top             =   12000
      Width           =   735
   End
   Begin VB.Label lblRB 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Index           =   1
      Left            =   6840
      TabIndex        =   45
      Top             =   12240
      Width           =   735
   End
   Begin VB.Label lblGB 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Index           =   1
      Left            =   6840
      TabIndex        =   44
      Top             =   12480
      Width           =   735
   End
   Begin VB.Label lblBlueF 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Index           =   1
      Left            =   6840
      TabIndex        =   43
      Top             =   13200
      Width           =   735
   End
   Begin VB.Label lblGreenF 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Index           =   1
      Left            =   6840
      TabIndex        =   42
      Top             =   12960
      Width           =   735
   End
   Begin VB.Label lblRedF 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Index           =   1
      Left            =   6840
      TabIndex        =   41
      Top             =   12720
      Width           =   735
   End
   Begin VB.Label lblBlueF 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Index           =   0
      Left            =   6000
      TabIndex        =   35
      Top             =   13200
      Width           =   735
   End
   Begin VB.Label lblGreenF 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Index           =   0
      Left            =   6000
      TabIndex        =   37
      Top             =   12960
      Width           =   735
   End
   Begin VB.Label lblRedF 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Index           =   0
      Left            =   6000
      TabIndex        =   39
      Top             =   12720
      Width           =   735
   End
   Begin VB.Label lblGB 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Index           =   0
      Left            =   6000
      TabIndex        =   25
      Top             =   12480
      Width           =   735
   End
   Begin VB.Label lblRB 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Index           =   0
      Left            =   6000
      TabIndex        =   23
      Top             =   12240
      Width           =   735
   End
   Begin VB.Label lblBG 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Index           =   0
      Left            =   6000
      TabIndex        =   21
      Top             =   12000
      Width           =   735
   End
   Begin VB.Label lblRG 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Index           =   0
      Left            =   6000
      TabIndex        =   19
      Top             =   11760
      Width           =   735
   End
   Begin VB.Label lblBR 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Index           =   0
      Left            =   6000
      TabIndex        =   17
      Top             =   11520
      Width           =   735
   End
   Begin VB.Label lblGR 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Index           =   0
      Left            =   6000
      TabIndex        =   15
      Top             =   11280
      Width           =   735
   End
   Begin VB.Label lbl 
      Caption         =   "RedF:"
      Height          =   255
      Index           =   17
      Left            =   5400
      TabIndex        =   40
      Top             =   12720
      Width           =   495
   End
   Begin VB.Label lbl 
      Caption         =   "GrnF:"
      Height          =   255
      Index           =   16
      Left            =   5400
      TabIndex        =   38
      Top             =   12960
      Width           =   495
   End
   Begin VB.Label lbl 
      Caption         =   "BluF:"
      Height          =   255
      Index           =   15
      Left            =   5400
      TabIndex        =   36
      Top             =   13200
      Width           =   495
   End
   Begin VB.Label lbl 
      Caption         =   "G/B:"
      Height          =   255
      Index           =   10
      Left            =   5400
      TabIndex        =   24
      Top             =   12480
      Width           =   495
   End
   Begin VB.Label lbl 
      Caption         =   "R/B:"
      Height          =   255
      Index           =   9
      Left            =   5400
      TabIndex        =   22
      Top             =   12240
      Width           =   495
   End
   Begin VB.Label lbl 
      Caption         =   "B/G:"
      Height          =   255
      Index           =   8
      Left            =   5400
      TabIndex        =   20
      Top             =   12000
      Width           =   495
   End
   Begin VB.Label lbl 
      Caption         =   "R/G:"
      Height          =   255
      Index           =   7
      Left            =   5400
      TabIndex        =   18
      Top             =   11760
      Width           =   495
   End
   Begin VB.Label lbl 
      Caption         =   "B/R:"
      Height          =   255
      Index           =   6
      Left            =   5400
      TabIndex        =   16
      Top             =   11520
      Width           =   495
   End
   Begin VB.Label lbl 
      Caption         =   "G/R:"
      Height          =   255
      Index           =   5
      Left            =   5400
      TabIndex        =   14
      Top             =   11280
      Width           =   495
   End
   Begin VB.Label lblNumSamples 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   7320
      TabIndex        =   32
      Top             =   9480
      Width           =   495
   End
   Begin VB.Label lbl 
      Caption         =   "# Samples:"
      Height          =   255
      Index           =   12
      Left            =   6360
      TabIndex        =   31
      Top             =   9480
      Width           =   855
   End
   Begin VB.Label lblRed 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Index           =   2
      Left            =   7680
      TabIndex        =   29
      Top             =   10080
      Width           =   735
   End
   Begin VB.Label lblRed 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Index           =   1
      Left            =   6840
      TabIndex        =   27
      Top             =   10080
      Width           =   735
   End
   Begin VB.Label lblSum 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Index           =   0
      Left            =   6000
      TabIndex        =   13
      Top             =   11040
      Width           =   735
   End
   Begin VB.Label lbl 
      Caption         =   "Sum:"
      Height          =   255
      Index           =   4
      Left            =   5400
      TabIndex        =   12
      Top             =   11040
      Width           =   495
   End
   Begin VB.Label lblAvg 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Index           =   0
      Left            =   6000
      TabIndex        =   11
      Top             =   10800
      Width           =   735
   End
   Begin VB.Label lbl 
      Caption         =   "Avg:"
      Height          =   255
      Index           =   3
      Left            =   5400
      TabIndex        =   10
      Top             =   10800
      Width           =   495
   End
   Begin VB.Label lblBlue 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Index           =   0
      Left            =   6000
      TabIndex        =   9
      Top             =   10560
      Width           =   735
   End
   Begin VB.Label lbl 
      Caption         =   "Blue:"
      Height          =   255
      Index           =   2
      Left            =   5400
      TabIndex        =   8
      Top             =   10560
      Width           =   495
   End
   Begin VB.Label lblGreen 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Index           =   0
      Left            =   6000
      TabIndex        =   7
      Top             =   10320
      Width           =   735
   End
   Begin VB.Label lbl 
      Caption         =   "Green:"
      Height          =   255
      Index           =   1
      Left            =   5400
      TabIndex        =   6
      Top             =   10320
      Width           =   495
   End
   Begin VB.Label lblRed 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Index           =   0
      Left            =   6000
      TabIndex        =   5
      Top             =   10080
      Width           =   735
   End
   Begin VB.Label lbl 
      Caption         =   "Red:"
      Height          =   255
      Index           =   0
      Left            =   5400
      TabIndex        =   4
      Top             =   10080
      Width           =   495
   End
   Begin VB.Label lbl 
      Caption         =   "Spot"
      Height          =   255
      Index           =   11
      Left            =   6000
      TabIndex        =   26
      Top             =   9840
      Width           =   495
   End
   Begin VB.Label lbl 
      Caption         =   "Avg"
      Height          =   255
      Index           =   23
      Left            =   6840
      TabIndex        =   28
      Top             =   9840
      Width           =   495
   End
   Begin VB.Label lbl 
      Caption         =   "Std"
      Height          =   255
      Index           =   35
      Left            =   7680
      TabIndex        =   30
      Top             =   9840
      Width           =   495
   End
End
Attribute VB_Name = "frmChartProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private mlngColor As Long
Private mconConnection As Connection
Private mchsChartElements As clsChartElements
Private mcheActiveChartElement As clsChartElement
Private mclcColorUnderMouse As clsColorComponents

Const DB_CONNECT_DRIVER = "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ="
Const DB_CONNECT_NAME = "\ChartPenData.mdb"
Const CHART_FILE = "8Chart24BitColor632.bmp"

Private Sub FilterImage()
  Dim lngBackground As Long
  Dim lngXIndex As Long
  Dim lngYIndex As Long
  Dim lngIndexCount As Long
  Dim lngIndexTotal As Long
  Dim lngTestColor As Long
  Dim strCmdText As String
  MousePointer = vbHourglass
  lngBackground = RGB(255, 255, 255)
  strCmdText = cmdPerformFilter.Caption
  With picFiltered
    lngIndexTotal = .ScaleWidth * .ScaleHeight
    For lngXIndex = 0 To .ScaleWidth
      For lngYIndex = 0 To .ScaleHeight
        lngTestColor = GetPixel(.hdc, lngXIndex, lngYIndex)
        If Not mcheActiveChartElement.IsValidColor(lngTestColor) Then
          SetPixelV .hdc, lngXIndex, lngYIndex, lngBackground
        End If
        lngIndexCount = lngIndexCount + 1
        DoEvents
      Next lngYIndex
      cmdPerformFilter.Caption = Format(lngIndexCount / lngIndexTotal * 100, "0.0") & " %"
      .Refresh
    Next lngXIndex
  End With
  cmdPerformFilter.Caption = strCmdText
  MousePointer = vbDefault
End Sub

Private Sub PopulateLabelsWithValues()
  With mclcColorUnderMouse
    lblRed(0).Caption = Format(.Red, "0")
    lblGreen(0).Caption = Format(.Green, "0")
    lblBlue(0).Caption = Format(.Blue, "0")
    lblAvg(0).Caption = Format(.Average, "0")
    lblSum(0).Caption = Format(.Sum, "0")
    lblRB(0).Caption = Format(.RB, "0.000")
    lblRG(0).Caption = Format(.RG, "0.000")
    lblGR(0).Caption = Format(.GR, "0.000")
    lblGB(0).Caption = Format(.GB, "0.000")
    lblBR(0).Caption = Format(.BR, "0.000")
    lblBG(0).Caption = Format(.BG, "0.000")
    lblRedF(0).Caption = Format(.RedFraction, "0.000")
    lblGreenF(0).Caption = Format(.GreenFraction, "0.000")
    lblBlueF(0).Caption = Format(.BlueFraction, "0.000")
  End With
End Sub

Private Sub SetZoomPicture(PicBox As PictureBox, X As Single, Y As Single)
  Dim lngLeft As Long
  Dim lngTop As Long
  Dim lngWidth As Long
  Dim lngHeight As Long
  Dim intZoomFactor As Integer
  intZoomFactor = 8
  With picZoom
    lngWidth = .ScaleWidth
    lngHeight = .ScaleHeight
  End With
  lngLeft = X - lngWidth / (intZoomFactor * 2)
  lngTop = Y - lngHeight / (intZoomFactor * 2)
  With PicBox
    If lngTop < 0 Then
      lngTop = 0
    ElseIf lngTop + lngHeight / intZoomFactor > .ScaleHeight Then
      lngTop = .ScaleHeight - lngHeight / intZoomFactor
    End If
    If lngLeft < 0 Then
      lngLeft = 0
    ElseIf lngLeft + lngWidth / intZoomFactor > .ScaleWidth Then
      lngLeft = .ScaleWidth - lngWidth / intZoomFactor
    End If
  End With
  StretchBlt picZoom.hdc, 0, 0, lngWidth * intZoomFactor, lngHeight * intZoomFactor, PicBox.hdc, lngLeft, _
    lngTop, lngWidth, lngHeight, vbSrcCopy
  picZoom.Refresh
End Sub

Private Sub UpdateStats()
  With mcheActiveChartElement
    lblNumSamples.Caption = .Red.NumberOfSamples
    lblRed(1).Caption = Format(.Red.Average, "0")
    lblGreen(1).Caption = Format(.Green.Average, "0")
    lblBlue(1).Caption = Format(.Blue.Average, "0")
    lblAvg(1).Caption = Format(.Average.Average, "0")
    lblSum(1).Caption = Format(.Sum.Average, "0")
    lblRB(1).Caption = Format(.RB.Average, "0.000")
    lblRG(1).Caption = Format(.RG.Average, "0.000")
    lblGR(1).Caption = Format(.GR.Average, "0.000")
    lblGB(1).Caption = Format(.GB.Average, "0.000")
    lblBR(1).Caption = Format(.BR.Average, "0.000")
    lblBG(1).Caption = Format(.BG.Average, "0.000")
    lblRedF(1).Caption = Format(.RedFraction.Average, "0.000")
    lblGreenF(1).Caption = Format(.GreenFraction.Average, "0.000")
    lblBlueF(1).Caption = Format(.BlueFraction.Average, "0.000")
    lblRed(2).Caption = Format(.Red.StandardDeviation, "0")
    lblGreen(2).Caption = Format(.Green.StandardDeviation, "0")
    lblBlue(2).Caption = Format(.Blue.StandardDeviation, "0")
    lblAvg(2).Caption = Format(.Average.StandardDeviation, "0")
    lblSum(2).Caption = Format(.Sum.StandardDeviation, "0")
    lblRB(2).Caption = Format(.RB.StandardDeviation, "0.000")
    lblRG(2).Caption = Format(.RG.StandardDeviation, "0.000")
    lblGR(2).Caption = Format(.GR.StandardDeviation, "0.000")
    lblGB(2).Caption = Format(.GB.StandardDeviation, "0.000")
    lblBR(2).Caption = Format(.BR.StandardDeviation, "0.000")
    lblBG(2).Caption = Format(.BG.StandardDeviation, "0.000")
    lblRedF(2).Caption = Format(.RedFraction.StandardDeviation, "0.000")
    lblGreenF(2).Caption = Format(.GreenFraction.StandardDeviation, "0.000")
    lblBlueF(2).Caption = Format(.BlueFraction.StandardDeviation, "0.000")
  End With
End Sub

Private Sub cmbChartElement_Click()
  Set mcheActiveChartElement = mchsChartElements(cmbChartElement.List(cmbChartElement.ListIndex))
  UpdateStats
End Sub

Sub cmdClearStats_Click()
  mcheActiveChartElement.ClearStats
  UpdateStats
End Sub

Private Sub cmdPerformFilter_Click()
  BitBlt picFiltered.hdc, 0, 0, picFiltered.ScaleWidth, picFiltered.ScaleHeight, picOrigChart.hdc, 0, 0, vbSrcCopy
  picFiltered.Refresh
  FilterImage
End Sub

Private Sub cmdRestoreElement_Click()
  mcheActiveChartElement.LoadFromDatabase mconConnection
  UpdateStats
End Sub

Private Sub cmdRestoreImage_Click()
  picOrigChart.Picture = LoadPicture(App.Path & "\" & CHART_FILE)
End Sub

Private Sub cmdSaveCurrentElement_Click()
  mcheActiveChartElement.SaveToDatabase mconConnection
End Sub

Private Sub cmdSetFilteredAsBase_Click()
  BitBlt picOrigChart.hdc, 0, 0, picOrigChart.ScaleWidth, picOrigChart.ScaleHeight, picFiltered.hdc, 0, 0, vbSrcCopy
  picOrigChart.Refresh
End Sub

Private Sub cmdSubFromOrigImage_Click()
  Dim recTemp As RECT
  recTemp.Left = 0
  recTemp.Top = 0
  recTemp.Bottom = picFiltered.ScaleHeight
  recTemp.Right = picFiltered.ScaleWidth
  BitBlt picFiltered.hdc, 0, 0, picFiltered.ScaleWidth, picFiltered.ScaleHeight, picOrigChart.hdc, 0, 0, vbSrcInvert
  InvertRect picFiltered.hdc, recTemp
  picFiltered.Refresh
End Sub

Private Sub Form_Initialize()
  Dim cheChartElement As clsChartElement
  Set mconConnection = New ADODB.Connection
  mconConnection.Open DB_CONNECT_DRIVER & App.Path & DB_CONNECT_NAME
  Set mchsChartElements = New clsChartElements
  mchsChartElements.LoadFromDatabase mconConnection
  For Each cheChartElement In mchsChartElements
    cmbChartElement.AddItem cheChartElement.ChartElementName
  Next
  If cmbChartElement.ListCount > 0 Then
    cmbChartElement.ListIndex = 1 'default to first item in the list
  End If
  Set mclcColorUnderMouse = New clsColorComponents
  picOrigChart.Picture = LoadPicture(App.Path & "\" & CHART_FILE)
  picFiltered.PaintPicture picOrigChart, 0, 0
End Sub

Private Sub Form_Load()
  BitBlt picFiltered.hdc, 0, 0, picFiltered.ScaleWidth, picFiltered.ScaleHeight, picOrigChart.hdc, 0, 0, vbSrcCopy
End Sub

Private Sub Form_Terminate()
  Set mchsChartElements = Nothing
  Set mclcColorUnderMouse = Nothing
  mconConnection.Close
  Set mconConnection = Nothing
End Sub

Private Sub picFiltered_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 1 Then 'left button pressed, zoom the filtered image
    SetZoomPicture picFiltered, X, Y
  ElseIf Button = 2 Then 'right button pressed, zoom the original image
    SetZoomPicture picOrigChart, X, Y
  End If
End Sub

Private Sub picOrigChart_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  SetZoomPicture picOrigChart, X, Y
End Sub

Private Sub picOrigChart_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  picOrigChart.ToolTipText = "(" & X & ", " & Y & ")"
End Sub

Private Sub picZoom_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  mcheActiveChartElement.ColorComponents.SetColorComponent mlngColor
  UpdateStats
End Sub

Private Sub picZoom_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  mlngColor = picZoom.Point(X, Y)
  picColorUnderMouse.BackColor = mlngColor
  mclcColorUnderMouse.SetColorComponent mlngColor
  PopulateLabelsWithValues
End Sub
