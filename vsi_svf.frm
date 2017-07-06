VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{572FF236-2066-11D4-8ED4-00E07D815373}#1.0#0"; "MBMsgEx.ocx"
Begin VB.Form frmLineOfSight 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "VSI/SVF v0.25 04/08/12"
   ClientHeight    =   4620
   ClientLeft      =   240
   ClientTop       =   480
   ClientWidth     =   4500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   4500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame frameOutput 
      Caption         =   "Output"
      Height          =   860
      Left            =   60
      TabIndex        =   41
      Top             =   3720
      Width           =   2775
      Begin VB.CommandButton CmdSav2CSV_File 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2220
         TabIndex        =   45
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox txtCsvName 
         Height          =   285
         Left            =   120
         TabIndex        =   44
         Top             =   480
         Width           =   1995
      End
      Begin VB.CheckBox chkSave2CSV 
         Caption         =   ".CSV file"
         Height          =   195
         Left            =   120
         TabIndex        =   43
         Top             =   240
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.CheckBox chkStore 
         Caption         =   "Point Shapefile"
         Height          =   195
         Left            =   1320
         TabIndex        =   42
         Top             =   240
         Width           =   1395
      End
   End
   Begin VB.Frame frameMode 
      Caption         =   "Mode"
      Height          =   1035
      Left            =   60
      TabIndex        =   37
      Top             =   2640
      Width           =   4395
      Begin VB.CheckBox chkShowSeg 
         Caption         =   "Show results of VoS per segment"
         Height          =   195
         Left            =   120
         TabIndex        =   40
         Top             =   720
         Width           =   2895
      End
      Begin VB.CheckBox chkShowMsg 
         Caption         =   "Pause for each point with ScreenLog"
         Height          =   195
         Left            =   120
         TabIndex        =   39
         Top             =   480
         Width           =   3135
      End
      Begin VB.CheckBox chkViz 
         Caption         =   "Show segment graphic (disable to speed up)"
         Height          =   195
         Left            =   120
         TabIndex        =   38
         Top             =   240
         Width           =   3615
      End
   End
   Begin VB.Frame frameConfig 
      Caption         =   "Configuration"
      Height          =   1755
      Left            =   60
      TabIndex        =   18
      Top             =   840
      Width           =   4395
      Begin VB.TextBox txtAlphaCoverDeg 
         Height          =   315
         Left            =   3480
         TabIndex        =   35
         Text            =   "Text1"
         Top             =   1380
         Width           =   615
      End
      Begin VB.TextBox txtBetaDeg 
         Height          =   285
         Left            =   1860
         TabIndex        =   34
         Text            =   "Text1"
         Top             =   1380
         Width           =   615
      End
      Begin VB.TextBox txtAlphaInitDeg 
         Height          =   285
         Left            =   3480
         TabIndex        =   30
         Text            =   "Text1"
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox txtRadius 
         Height          =   285
         Left            =   3480
         TabIndex        =   26
         Text            =   "Text1"
         Top             =   780
         Width           =   615
      End
      Begin VB.TextBox txtiNLoS 
         Height          =   285
         Left            =   3480
         TabIndex        =   25
         Text            =   "Text1"
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox txtTarOffset 
         Height          =   285
         Left            =   3480
         TabIndex        =   24
         Text            =   "Text1"
         Top             =   180
         Width           =   615
      End
      Begin VB.TextBox txtObsOffset 
         Height          =   285
         Left            =   1860
         TabIndex        =   22
         Text            =   "Text1"
         Top             =   180
         Width           =   615
      End
      Begin VB.Label Label10 
         Caption         =   "°"
         Height          =   195
         Left            =   4140
         TabIndex        =   36
         Top             =   1440
         Width           =   195
      End
      Begin VB.Label Label8 
         Caption         =   "° H(max360°)"
         Height          =   195
         Left            =   2520
         TabIndex        =   33
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "SightSpan: V(max90°)"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "°"
         Height          =   195
         Left            =   4140
         TabIndex        =   31
         Top             =   1140
         Width           =   135
      End
      Begin VB.Label Label3 
         Caption         =   "Initial Orientation (max.360°):"
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   1140
         Width           =   2295
      End
      Begin VB.Label Label14 
         Caption         =   "m"
         Height          =   255
         Left            =   4140
         TabIndex        =   28
         Top             =   840
         Width           =   135
      End
      Begin VB.Label Label2 
         Caption         =   "m"
         Height          =   195
         Left            =   4140
         TabIndex        =   27
         Top             =   240
         Width           =   195
      End
      Begin VB.Label Label7 
         Caption         =   "m     Target"
         Height          =   255
         Left            =   2580
         TabIndex        =   23
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Initial Radius (max.3500m):"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "No. of Vol.Of.Sight (max.360):"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   540
         Width           =   2175
      End
      Begin VB.Label Label12 
         Caption         =   "Z offset: Observer"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   1395
      End
   End
   Begin VB.Frame frameInputPointFeature 
      Caption         =   "Source of Point Feature"
      Height          =   555
      Left            =   60
      TabIndex        =   14
      Top             =   240
      Width           =   4395
      Begin VB.CommandButton cmdBrInputSF 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3960
         TabIndex        =   17
         Top             =   180
         Width           =   330
      End
      Begin VB.OptionButton optSFFC 
         Caption         =   "Select a ShapeFile"
         Height          =   255
         Left            =   2160
         TabIndex        =   16
         Top             =   240
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.OptionButton optMapFC 
         Caption         =   "Select a Layer"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1455
      End
   End
   Begin LOSTool.TransMsg TransMsg1 
      Height          =   360
      Left            =   1500
      TabIndex        =   13
      Top             =   -60
      Visible         =   0   'False
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   635
   End
   Begin VB.CheckBox chkCurv 
      Caption         =   "Curvature and refraction correction"
      Height          =   255
      Left            =   3120
      TabIndex        =   0
      Top             =   3960
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton cmdLogToggle 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   12
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   3900
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4440
      Width           =   495
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2460
      Top             =   -240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton CmdAbout 
      Caption         =   "About"
      Height          =   375
      Left            =   2940
      TabIndex        =   9
      Top             =   4200
      Width           =   735
   End
   Begin VB.Frame FrameLog 
      Caption         =   "Log"
      Height          =   3060
      Left            =   60
      TabIndex        =   5
      Top             =   4620
      Visible         =   0   'False
      Width           =   4395
      Begin VB.CommandButton CmdReset 
         Caption         =   "Reset"
         Height          =   495
         Left            =   3660
         TabIndex        =   8
         Top             =   2340
         Width           =   615
      End
      Begin VB.CommandButton CmdCopy 
         Caption         =   "Copy All"
         Height          =   495
         Left            =   3660
         TabIndex        =   7
         Top             =   1740
         Width           =   615
      End
      Begin VB.TextBox TxtLog 
         Height          =   2655
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Text            =   "VSI_SVF.frx":0000
         Top             =   240
         Width           =   3435
      End
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3780
      TabIndex        =   4
      Top             =   3780
      Width           =   675
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2940
      TabIndex        =   3
      Top             =   3780
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      Height          =   495
      Left            =   3540
      Picture         =   "VSI_SVF.frx":003E
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   1
      Top             =   -180
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3000
      Top             =   -240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VSI_SVF.frx":0348
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Height          =   375
      Left            =   2940
      TabIndex        =   11
      Top             =   3780
      Visible         =   0   'False
      Width           =   735
   End
   Begin MBMsgBoxEx.MsgBoxEx MsgBoxEx1 
      Left            =   1920
      Top             =   -120
      _ExtentX        =   847
      _ExtentY        =   847
      Foreground      =   -1  'True
      Position        =   0
      CustomIcon      =   "VSI_SVF.frx":069A
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Log"
      Height          =   255
      Left            =   3840
      TabIndex        =   12
      Top             =   4200
      Width           =   615
   End
   Begin VB.Label Label17 
      Caption         =   "Please make sure a TIN or Raster layer is loaded."
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   3855
   End
End
Attribute VB_Name = "frmLineOfSight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public m_obsOffset As Double ' These are module level so the tool remembers the _
values
Public m_tarOffset As Double ' when it's unselected then selected again.
Public m_iNLoS As Integer
Public m_Radius As Double
Public m_BetaDeg As Double
Public m_AlphaInitDeg As Double
Public m_AlphaCoverDeg As Double

Private m_inSFFC As IFeatureClass 'v0.23
Private m_inSFnumFeat As Long 'v0.23
Private m_pApp As IApplication 'addbutton
Private m_bFeedback_Start As Boolean 'addbutton
Private m_pActiveView As IActiveView 'addbutton
Private m_pDispTrans As IDisplayTransformation 'addbutton
Private m_pDispFeedback As IDisplayFeedback 'addbutton
Private m_pNewLineFeedback As INewLineFeedback 'addbutton
Private m_pScenePoints As IPointCollection 'addbutton
Private m_pDDDExtConfig As IExtensionConfig 'addbutton
Private m_Stop As Boolean
Private m_StopfrPause As Boolean
Private MsgEx As New CMsgBoxEx 'v0.14
Private frmHeightOri As Long 'v0.14
Private pWindPos As IWindowPosition 'v0.14
' Variables used by the Error handler function - DO NOT REMOVE
Const c_sModuleFileName As String = "K:\Library\Progm\v0.21\VSI_SVF.frm"


' Variables used by the Error handler function - DO NOT REMOVE
' Variables used by the Error handler function - DO NOT REMOVE
' This should only be called once, when the command is created

Public Sub Init(pApp As IApplication, obsOffset As Double, tarOffset As Double, _
      iNLoS As Integer, radius As Double, BetaDeg As Double, AlphaInitDeg As Double, _
AlphaCoverDeg As Double)
    On Error GoTo ErrorHandler
    
40:     Set m_pApp = pApp 'addbutton
    '  If (TypeOf m_pApp Is IMxApplication) Then 'addbutton
    '    Dim pMxDoc As IMxDocument 'addbutton
    '    Set pMxDoc = m_pApp.Document 'addbutton
    '    Set m_pMap = pMxDoc.FocusMap 'addbutton
    '  ElseIf (TypeOf m_pApp Is ISxApplication) Then 'addbutton
    '    Dim pSxDoc As ISxDocument 'addbutton
    '    Set pSxDoc = m_pApp.Document 'addbutton
    '    Set m_pMap = pSxDoc.Scene 'addbutton
    '   Exit Sub 'addbutton
    '  End If 'addbutton
    
52:     m_obsOffset = obsOffset
53:     m_tarOffset = tarOffset
54:     m_iNLoS = iNLoS
55:     m_Radius = radius
56:     m_BetaDeg = BetaDeg
57:     m_AlphaInitDeg = AlphaInitDeg
58:     m_AlphaCoverDeg = AlphaCoverDeg
    


  Exit Sub
ErrorHandler:
  HandleError True, "Init " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Public Function GetObserverOffset() As Double
    On Error GoTo EH
69:     GetObserverOffset = CDbl(txtObsOffset.Text)
    Exit Function
EH:
72:     MsgBox Err.Description
End Function

Public Function GetTargetOffset() As Double
    On Error GoTo EH
77:     GetTargetOffset = CDbl(txtTarOffset.Text)
    Exit Function
EH:
80:     MsgBox Err.Description
End Function

Public Function GetiNLoS() As Double
    On Error GoTo EH
85:     GetiNLoS = CDbl(txtiNLoS.Text)
    Exit Function
EH:
88:     MsgBox Err.Description
End Function

Public Function GetRadius() As Double
    On Error GoTo EH
93:     GetRadius = CDbl(txtRadius.Text)
    Exit Function
EH:
96:     MsgBox Err.Description
End Function

Public Function GetBetaDeg() As Double
    On Error GoTo EH
101:     GetBetaDeg = CDbl(txtBetaDeg.Text)
    Exit Function
EH:
104:     MsgBox Err.Description
End Function

Public Function GetAlphaInitDeg() As Double
    On Error GoTo EH
109:     GetAlphaInitDeg = CDbl(txtAlphaInitDeg.Text)
    Exit Function
EH:
112:     MsgBox Err.Description
End Function

Public Function GetAlphaCoverDeg() As Double
    On Error GoTo EH
117:     GetAlphaCoverDeg = CDbl(txtAlphaCoverDeg.Text)
    Exit Function
EH:
120:     MsgBox Err.Description
End Function

Public Sub CurvatureEnabled(bEnabled As Boolean)
  On Error GoTo ErrorHandler

    
127:     chkCurv.Enabled = bEnabled
    


  Exit Sub
ErrorHandler:
  HandleError True, "CurvatureEnabled " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Public Function GetCurvatureEnabled() As Boolean
  On Error GoTo ErrorHandler

    
140:     GetCurvatureEnabled = (chkCurv.Value = 1)
    


  Exit Function
ErrorHandler:
  HandleError True, "GetCurvatureEnabled " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Public Sub VizEnabled(bEnabled As Boolean)
  On Error GoTo ErrorHandler

    
153:     chkViz.Enabled = bEnabled
    


  Exit Sub
ErrorHandler:
  HandleError True, "VizEnabled " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Public Function GetVizEnabled() As Boolean
  On Error GoTo ErrorHandler

    
166:     GetVizEnabled = (chkViz.Value = 1)
    


  Exit Function
ErrorHandler:
  HandleError True, "GetVizEnabled " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Public Sub ShowMsgEnabled(bEnabled As Boolean)
  On Error GoTo ErrorHandler

    
179:     chkShowMsg.Enabled = bEnabled
    


  Exit Sub
ErrorHandler:
  HandleError True, "ShowMsgEnabled " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Public Function GetShowMsgEnabled() As Boolean
  On Error GoTo ErrorHandler

    
192:     GetShowMsgEnabled = (chkShowMsg.Value = 1)
    


  Exit Function
ErrorHandler:
  HandleError True, "GetShowMsgEnabled " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Public Sub ShowSegEnabled(bEnabled As Boolean)
  On Error GoTo ErrorHandler

    
205:     chkShowSeg.Enabled = bEnabled
    


  Exit Sub
ErrorHandler:
  HandleError True, "ShowSegEnabled " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Public Function GetShowSegEnabled() As Boolean
  On Error GoTo ErrorHandler

    
218:     GetShowSegEnabled = (chkShowSeg.Value = 1)
    


  Exit Function
ErrorHandler:
  HandleError True, "GetShowSegEnabled " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Public Sub StoreEnabled(bEnabled As Boolean)
  On Error GoTo ErrorHandler

    
231:     chkStore.Enabled = bEnabled
    


  Exit Sub
ErrorHandler:
  HandleError True, "StoreEnabled " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Public Function GetStoreEnabled() As Boolean
  On Error GoTo ErrorHandler

    
244:     GetStoreEnabled = (chkStore.Value = 1)
    


  Exit Function
ErrorHandler:
  HandleError True, "GetStoreEnabled " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Private Sub cmdPntFrmLayer_Click()
  On Error GoTo ErrorHandler

    
    'Dim bPntFrmLayer as Boolean
    'bPntFrmLayer = Yes
    
    'Get selected layer
    'If selected layer is point feature class then go
    'add to point collection
    


  Exit Sub
ErrorHandler:
  HandleError False, "cmdPntFrmLayer_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub chkSave2CSV_Click()
  On Error GoTo ErrorHandler

    
275:     If chkSave2CSV.Value = 1 Then
276:         txtCsvName.Enabled = 1
             txtCsvName.BackColor = &H80000005
277:         CmdSav2CSV_File.Enabled = 1
278:     Else
279:         txtCsvName.Enabled = 0
             txtCsvName.BackColor = &H8000000F
280:         CmdSav2CSV_File.Enabled = 0
281:         txtCsvName.Text = "" 'v0.14
282:     End If
    


  Exit Sub
ErrorHandler:
  HandleError True, "chkSave2CSV_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub CmdAbout_Click()
  On Error GoTo ErrorHandler

    
295:     ShowMsgInfo "About", ("Copyright © 2003-2004 NUS" & Chr(13) & "Yang, Putra and Li")
    


  Exit Sub
ErrorHandler:
  HandleError True, "CmdAbout_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub cmdBrInputSF_Click()
  On Error GoTo ErrorHandler
  Dim pBrowser As IGxDialog
307:   Set pBrowser = New GxDialog
  Dim pEnumGX As IEnumGxObject

  Dim pObjectFilter As IGxObjectFilter
311:   Set pObjectFilter = New GxFilterShapefiles

  ' Open browser
  Dim blnFlag As Boolean
315:   pBrowser.Title = "Select ShapeFile"
316:   pBrowser.ButtonCaption = "Select"
317:   pBrowser.AllowMultiSelect = False
318:   Set pBrowser.ObjectFilter = pObjectFilter
319:   blnFlag = pBrowser.DoModalOpen(frmLineOfSight.hwnd, pEnumGX)
  
  If blnFlag = False Then Exit Sub
  
323:   Screen.MousePointer = vbHourglass
  
  ' Get the first GXObject in the enumeration
326:   pEnumGX.Reset
  Dim pGXObject As IGxObject
328:   Set pGXObject = pEnumGX.Next
  
  Dim pName As IName
331:   Set pName = pGXObject.InternalObjectName
  'SFfullname = pGXObject.FullName
  
'  If TypeOf pName Is IFeatureClassName Then
335:     Set m_inSFFC = pName.Open
336:     m_inSFnumFeat = m_inSFFC.FeatureCount(Nothing)
'  End If

339:   Set pBrowser = Nothing
340:   Screen.MousePointer = vbDefault
  Exit Sub

ErrorHandler:
  HandleError False, "m_cmdBrowseOutput_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
  ' Pressed Cancel on the GXDialog
346:    Screen.MousePointer = vbDefault

End Sub

Private Sub CmdCopy_Click()
  On Error GoTo ErrorHandler

    
354:     If Not (TxtLog = Empty) Then
355:         Clipboard.Clear
356:         Clipboard.SetText TxtLog.Text
357:         Else: ShowMsgCritical "Warning", ("Logging TextBox is empty now," & (Chr(13)) & _
        "please run the program first!")
359:     End If
    


  Exit Sub
ErrorHandler:
  HandleError True, "CmdCopy_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub cmdLogToggle_Click()
  On Error GoTo ErrorHandler

    
372:     If cmdLogToggle.Caption = "6" Then
373:         cmdLogToggle.Caption = "5"
374:         cmdLogToggle.BackColor = &H80000010
375:         frmLineOfSight.Height = FrameLog.Top + FrameLog.Height + 395
376:         FrameLog.Visible = True
377:     Else
378:         cmdLogToggle.Caption = "6"
379:         cmdLogToggle.BackColor = &H8000000F
380:         frmLineOfSight.Height = frmHeightOri
381:         FrameLog.Visible = False
382:     End If
    


  Exit Sub
ErrorHandler:
  HandleError True, "cmdLogToggle_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub CmdOK_Click()
  On Error GoTo ErrorHandler

    
395:     CmdOK.Visible = False
396:     cmdStop.Visible = True
    
    'Check if Stop button has been pressed before this session starts 'v0.13
399:     If m_Stop = True Then
400:         m_Stop = False
401:     End If
    
    'If not (bPntFrmLayer = Yes) then
    
    Dim pFCursor As IFeatureCursor '27/02/04
    Dim pFeature As IFeature '27/02/04
    Dim pFClass As IFeatureClass 'ArcMAP 27/02/04
    Dim pPnt As IPoint
    
    Dim TxtCsv As String
411:     TxtCsv = ""
    
413:     If (TypeOf m_pApp Is IMxApplication) Then 'ArcMAP
414:         If (m_bFeedback_Start = False) Then
            
            Dim pMxDoc As IMxDocument
417:             Set pMxDoc = m_pApp.Document
418:             Set m_pActiveView = pMxDoc.ActiveView
            
420:             Set m_pNewLineFeedback = New NewLineFeedback
421:             Set m_pDispFeedback = m_pNewLineFeedback
422:             Set m_pDispFeedback.Display = m_pActiveView.ScreenDisplay
            
            Dim pDisp As IDisplay
425:             Set pDisp = m_pActiveView.ScreenDisplay
426:             Set m_pDispTrans = pDisp.DisplayTransformation
            
428:             m_bFeedback_Start = True
            
            'Observer Point from Point Class
            
432:             If optMapFC.Value = True Then
                Dim pFlayerM As IFeatureLayer 'ArcMAP
            '      Dim pFClassM As IFeatureClass 'ArcMAP 27/02/04
435:             Set pFlayerM = pMxDoc.SelectedLayer 'ArcMAP
            
437:             If pFlayerM Is Nothing Then
                
439:                 ShowMsgCritical "Warning", "Select a single point shapefile!"
                
441:                 cmdStop.Visible = False
442:                 CmdOK.Visible = True
                Exit Sub
444:             End If
445:             Set pFClass = pFlayerM.FeatureClass '27/02/04
446:             Else
447:                 If m_inSFnumFeat = 0 Then
448:                     ShowMsgCritical "Warning", "Select a single point shapefile!"
                        Exit Sub
450:                 Else
451:                         Set pFClass = m_inSFFC
452:                 End If
453:             End If
            
455:             If pFClass.ShapeType <> esriGeometryPoint Then  '27/02/04
456:                 ShowMsgCritical "Warning", "Select a single point shapefile!"
457:                 cmdStop.Visible = False
458:                 CmdOK.Visible = True
                Exit Sub
460:             End If
            
            '      Dim pFCursor As IFeatureCursor
463:             Set pFCursor = pFClass.Search(Nothing, False) 'ArcMAP 27/02/04
            
            '      Dim pFeature As IFeature
466:             Set pFeature = pFCursor.NextFeature 'ArcMAP
            'Dim pGeomInput As IGeometry
            'Dim pClone As IClone
            
            '      Dim pPnt As IPoint '27/02/04
            
            'Set pPnt = m_pDispTrans.ToMapPoint(x, y)
            
            'm_pNewLineFeedback.Start pPnt
            
            'm_mouseDownX = x
            'm_mouseDownY = y
            
            'Exit Sub
            'Else
            'Set pPnt = m_pDispTrans.ToMapPoint(x, y)
            
            'm_pDispFeedback.MoveTo pPnt
            'm_pNewLineFeedback.AddPoint pPnt
            
            'Dim pPolyline As IPolyline
            'Set pPolyline = m_pNewLineFeedback.Stop
488:         End If
489:     Else
        Dim pSxDoc As ISxDocument 'ArcScene
491:         Set pSxDoc = m_pApp.Document
        Dim pSG As ISceneGraph
493:         Set pSG = pSxDoc.Scene.SceneGraph
        
        Dim pOwner As stdole.IUnknown
        Dim pObject As stdole.IUnknown
        
        'pSG.Locate pSG.ActiveViewer, x, y, esriScenePickGeography, True, pPnt, pOwner, _
        pObject
        
        'Input Point from selected Point Shapefile
502:         If optMapFC.Value = True Then
            Dim pFlayerS As IFeatureLayer 'ArcScene
        '    Dim pFClassS As IFeatureClass 'ArcScene 27/02/04
        
506:         Set pFlayerS = pSxDoc.SelectedLayer 'ArcScene
        
508:         If pFlayerS Is Nothing Then
509:             ShowMsgCritical "Warning", "Select a single point shapefile!"
            
            Exit Sub
512:         End If
513:         Set pFClass = pFlayerS.FeatureClass '27/02/04
        
515:         If pFClass.ShapeType <> esriGeometryPoint Then '27/02/04
516:             ShowMsgCritical "Warning", "Select a single point shapefile!"
            
            Exit Sub
519:         End If
520:         Else
521:                 If m_inSFnumFeat = 0 Then
522:                     ShowMsgCritical "Warning", "Select a single point shapefile!"
                   Exit Sub
524:                      Else
525:                         Set pFClass = m_inSFFC
526:                 End If
527:         End If
        'Dim pFCursor As IFeatureCursor
529:         Set pFCursor = pFClass.Search(Nothing, False) 'ArcScene 27/02/04
530:         Set pFeature = pFCursor.NextFeature 'ArcScene
        
532:     End If 'ArcScene
    
    'MAKE NEW FIELDS if none START
    'If (TypeOf m_pApp Is IMxApplication) Then 'ArcMAP 27/02/04
536:     If frmLineOfSight.chkStore = 1 Then 'Optional Store Value 27/02/04
        Dim indexRadius As Long
538:         indexRadius = pFClass.FindField("UserRadius")
539:         If indexRadius < 0 Then
            Dim pFieldRadius As IFieldEdit
541:             Set pFieldRadius = New Field
542:             With pFieldRadius
543:                 .Type = esriFieldTypeDouble
544:                 .Name = "UserRadius"
545:             End With
546:             pFClass.AddField pFieldRadius
547:         End If
        
        Dim indexTotVol As Long
550:         indexTotVol = pFClass.FindField("VOLUME")
551:         If indexTotVol < 0 Then 'If there is no the field of TotVol, indexTotVol =-1
            Dim pFieldTotVol As IFieldEdit
553:             Set pFieldTotVol = New Field
554:             With pFieldTotVol
555:                 .Type = esriFieldTypeDouble
556:                 .Name = "VOLUME"
557:             End With
558:             pFClass.AddField pFieldTotVol
559:         End If
        
        Dim indexVSI As Long
562:         indexVSI = pFClass.FindField("VSI")
563:         If indexVSI < 0 Then
            Dim pFieldVSI As IFieldEdit
565:             Set pFieldVSI = New Field
566:             With pFieldVSI
567:                 .Type = esriFieldTypeDouble
568:                 .Name = "VSI"
569:             End With
570:             pFClass.AddField pFieldVSI
571:         End If
        
        Dim indexVSISTD As Long
574:         indexVSISTD = pFClass.FindField("VSISTD")
575:         If indexVSISTD < 0 Then
            Dim pFieldVSISTD As IFieldEdit
577:             Set pFieldVSISTD = New Field
578:             With pFieldVSISTD
579:                 .Type = esriFieldTypeDouble
580:                 .Name = "VSISTD"
581:             End With
582:             pFClass.AddField pFieldVSISTD
583:         End If
        
        Dim indexMaxRad As Long
586:         indexMaxRad = pFClass.FindField("Max_Radius")
587:         If indexMaxRad < 0 Then
            Dim pFieldMaxRad As IFieldEdit
589:             Set pFieldMaxRad = New Field
590:             With pFieldMaxRad
591:                 .Type = esriFieldTypeDouble
592:                 .Name = "Max_Radius"
593:             End With
594:             pFClass.AddField pFieldMaxRad
595:         End If
        
        Dim indexVSImax As Long
598:         indexVSImax = pFClass.FindField("VSImax")
599:         If indexVSImax < 0 Then
            Dim pFieldVSImax As IFieldEdit
601:             Set pFieldVSImax = New Field
602:             With pFieldVSImax
603:                 .Type = esriFieldTypeDouble
604:                 .Name = "VSImax"
605:             End With
606:             pFClass.AddField pFieldVSImax
607:         End If
        
        Dim indexVSImaxSTD As Long
610:         indexVSImaxSTD = pFClass.FindField("VSImaxSTD")
611:         If indexVSImaxSTD < 0 Then
            Dim pFieldVSImaxSTD As IFieldEdit
613:             Set pFieldVSImaxSTD = New Field
614:             With pFieldVSImaxSTD
615:                 .Type = esriFieldTypeDouble
616:                 .Name = "VSImaxSTD"
617:             End With
618:             pFClass.AddField pFieldVSImaxSTD
619:         End If
        
        Dim indexAveRad As Long
622:         indexAveRad = pFClass.FindField("Ave_Radius")
623:         If indexAveRad < 0 Then
            Dim pFieldAveRad As IFieldEdit
625:             Set pFieldAveRad = New Field
626:             With pFieldAveRad
627:                 .Type = esriFieldTypeDouble
628:                 .Name = "Ave_Radius"
629:             End With
630:             pFClass.AddField pFieldAveRad
631:         End If
        
        Dim indexVSIave As Long
634:         indexVSIave = pFClass.FindField("VSIave")
635:         If indexVSIave < 0 Then
            Dim pFieldVSIave As IFieldEdit
637:             Set pFieldVSIave = New Field
638:             With pFieldVSIave
639:                 .Type = esriFieldTypeDouble
640:                 .Name = "VSIave"
641:             End With
642:             pFClass.AddField pFieldVSIave
643:         End If
        
        Dim indexVSIaveSTD As Long
646:         indexVSIaveSTD = pFClass.FindField("VSIaveSTD")
647:         If indexVSIaveSTD < 0 Then
            Dim pFieldVSIaveSTD As IFieldEdit
649:             Set pFieldVSIaveSTD = New Field
650:             With pFieldVSIaveSTD
651:                 .Type = esriFieldTypeDouble
652:                 .Name = "VSIaveSTD"
653:             End With
654:             pFClass.AddField pFieldVSIaveSTD
655:         End If
        
        Dim indexMinRad As Long
658:         indexMinRad = pFClass.FindField("Min_Radius")
659:         If indexMinRad < 0 Then
            Dim pFieldMinRad As IFieldEdit
661:             Set pFieldMinRad = New Field
662:             With pFieldMinRad
663:                 .Type = esriFieldTypeDouble
664:                 .Name = "Min_Radius"
665:             End With
666:             pFClass.AddField pFieldMinRad
667:         End If
        
        Dim indexVSImin As Long
670:         indexVSImin = pFClass.FindField("VSImin")
671:         If indexVSImin < 0 Then
            Dim pFieldVSImin As IFieldEdit
673:             Set pFieldVSImin = New Field
674:             With pFieldVSImin
675:                 .Type = esriFieldTypeDouble
676:                 .Name = "VSImin"
677:             End With
678:             pFClass.AddField pFieldVSImin
679:         End If
        
        Dim indexVSIminSTD As Long
682:         indexVSIminSTD = pFClass.FindField("VSIminSTD")
683:         If indexVSIminSTD < 0 Then
            Dim pFieldVSIminSTD As IFieldEdit
685:             Set pFieldVSIminSTD = New Field
686:             With pFieldVSIminSTD
687:                 .Type = esriFieldTypeDouble
688:                 .Name = "VSIminSTD"
689:             End With
690:             pFClass.AddField pFieldVSIminSTD
691:         End If
        
        Dim indexSVF As Long
694:         indexSVF = pFClass.FindField("SVF")
695:         If indexSVF < 0 Then
            Dim pFieldSVF As IFieldEdit
697:             Set pFieldSVF = New Field
698:             With pFieldSVF
699:                 .Type = esriFieldTypeDouble
700:                 .Name = "SVF"
701:             End With
702:             pFClass.AddField pFieldSVF
703:         End If
        
        Dim indexSVFSTD As Long 'rocky
706:         indexSVFSTD = pFClass.FindField("SVFSTD")
707:         If indexSVFSTD < 0 Then
            Dim pFieldSVFSTD As IFieldEdit
709:             Set pFieldSVFSTD = New Field
710:             With pFieldSVFSTD
711:                 .Type = esriFieldTypeDouble
712:                 .Name = "SVFSTD"
713:             End With
714:             pFClass.AddField pFieldSVFSTD
715:         End If
        
        
        Dim indexBetaMax As Long
719:         indexBetaMax = pFClass.FindField("BetaMax")
720:         If indexBetaMax < 0 Then
            Dim pFieldBetaMax As IFieldEdit
722:             Set pFieldBetaMax = New Field
723:             With pFieldBetaMax
724:                 .Type = esriFieldTypeDouble
725:                 .Name = "BetaMax"
726:             End With
727:             pFClass.AddField pFieldBetaMax
728:         End If
        
        Dim indexBetaAve As Long
731:         indexBetaAve = pFClass.FindField("BetaAve")
732:         If indexBetaAve < 0 Then
            Dim pFieldBetaAve As IFieldEdit
734:             Set pFieldBetaAve = New Field
735:             With pFieldBetaAve
736:                 .Type = esriFieldTypeDouble
737:                 .Name = "BetaAve"
738:             End With
739:             pFClass.AddField pFieldBetaAve
740:         End If
        
        Dim indexVisArea As Long
743:         indexVisArea = pFClass.FindField("VisArea")
744:         If indexVisArea < 0 Then
            Dim pFieldVisArea As IFieldEdit
746:             Set pFieldVisArea = New Field
747:             With pFieldVisArea
748:                 .Type = esriFieldTypeDouble
749:                 .Name = "VisArea"
750:             End With
751:             pFClass.AddField pFieldVisArea
752:         End If
        
        Dim indexVisPeri As Long
755:         indexVisPeri = pFClass.FindField("VisPeri")
756:         If indexVisPeri < 0 Then
            Dim pFieldVisPeri As IFieldEdit
758:             Set pFieldVisPeri = New Field
759:             With pFieldVisPeri
760:                 .Type = esriFieldTypeDouble
761:                 .Name = "VisPeri"
762:             End With
763:             pFClass.AddField pFieldVisPeri
764:         End If
        
        Dim indexLongestAxis As Long
767:         indexLongestAxis = pFClass.FindField("LongAxis")
768:         If indexLongestAxis < 0 Then
            Dim pFieldLongAxis As IFieldEdit
770:             Set pFieldLongAxis = New Field
771:             With pFieldLongAxis
772:                 .Type = esriFieldTypeDouble
773:                 .Name = "LongAxis"
774:             End With
775:             pFClass.AddField pFieldLongAxis
776:         End If
        
        Dim indexCompact As Long
779:         indexCompact = pFClass.FindField("Compact")
780:         If indexCompact < 0 Then
            Dim pFieldCompact As IFieldEdit
782:             Set pFieldCompact = New Field
783:             With pFieldCompact
784:                 .Type = esriFieldTypeDouble
785:                 .Name = "Compact"
786:             End With
787:             pFClass.AddField pFieldCompact
788:         End If
        
        Dim indexFractal As Long
791:         indexFractal = pFClass.FindField("Fractal")
792:         If indexFractal < 0 Then
            Dim pFieldFractal As IFieldEdit
794:             Set pFieldFractal = New Field
795:             With pFieldFractal
796:                 .Type = esriFieldTypeDouble
797:                 .Name = "Fractal"
798:             End With
799:             pFClass.AddField pFieldFractal
800:         End If
        
        Dim indexConvo As Long
803:         indexConvo = pFClass.FindField("Convolut")
804:         If indexConvo < 0 Then
            Dim pFieldConvo As IFieldEdit
806:             Set pFieldConvo = New Field
807:             With pFieldConvo
808:                 .Type = esriFieldTypeDouble
809:                 .Name = "Convolut"
810:             End With
811:             pFClass.AddField pFieldConvo
812:         End If
        
        Dim indexElliptic As Long
815:         indexElliptic = pFClass.FindField("Elliptic")
816:         If indexElliptic < 0 Then
            Dim pFieldElliptic As IFieldEdit
818:             Set pFieldElliptic = New Field
819:             With pFieldElliptic
820:                 .Type = esriFieldTypeDouble
821:                 .Name = "Elliptic"
822:             End With
823:             pFClass.AddField pFieldElliptic
824:         End If
        
        Dim indexRMinMax As Long
827:         indexRMinMax = pFClass.FindField("RMinMax")
828:         If indexRMinMax < 0 Then
            Dim pFieldRMinMax As IFieldEdit
830:             Set pFieldRMinMax = New Field
831:             With pFieldRMinMax
832:                 .Type = esriFieldTypeDouble
833:                 .Name = "RMinMax"
834:             End With
835:             pFClass.AddField pFieldRMinMax
836:         End If
        
        Dim indexRAveMax As Long
839:         indexRAveMax = pFClass.FindField("RAveMax")
840:         If indexRAveMax < 0 Then
            Dim pFieldRAveMax As IFieldEdit
842:             Set pFieldRAveMax = New Field
843:             With pFieldRAveMax
844:                 .Type = esriFieldTypeDouble
845:                 .Name = "RAveMax"
846:             End With
847:             pFClass.AddField pFieldRAveMax
848:         End If
        
        Dim indexRMinAve As Long
851:         indexRMinAve = pFClass.FindField("RMinAve")
852:         If indexRMinAve < 0 Then
            Dim pFieldRMinAve As IFieldEdit
854:             Set pFieldRMinAve = New Field
855:             With pFieldRMinAve
856:                 .Type = esriFieldTypeDouble
857:                 .Name = "RMinAve"
858:             End With
859:             pFClass.AddField pFieldRMinAve
860:         End If
        
862:     End If ' Optional Store value 27/02/04
    'End If 'ArcMAP 27/02/04
    
    'MAKE NEW FIELDS if none STOP
    
    'Start getting point features
    'Dim pFeature As IFeature
    'Set pFeature = pFCursor.NextFeature
    'Dim pGeomInput As IGeometry
    '    Dim pPntCol As IPointCollection
    
    'Else (bPntFrmLayer)
    'For p = 1 to pPointColl.Count
    'pPnt = pPointColl(p)
    'Next p
    
    'Dim Point1 As IPoint, Point2 As IPoint
    'Dim r As Double
    'Set Point1 = pPolyline.FromPoint
    'Set Point2 = pPolyline.ToPoint
    'r = Sqr((Point1.x - Point2.x) ^ 2 + (Point1.y - Point2.y) ^ 2)
    'x = x - (r * (Cos(frmLineOfSight.GetRotateDegree)))
    'y = y + (r * (Sin(frmLineOfSight.GetRotateDegree)))
    'x = x - (r * (Cos(30)))
    'y = y + (r * (Sin(30)))
    'pPnt.x = x
    'pPnt.y = y
    
    
    'Flashing
    'Dim pFlashPoint As IPoint
    'Dim pClone As IClone
    'Set pClone = pPnt
    'Set pFlashPoint = pClone.Clone
    'pFlashPoint.Z = pFlashPoint.Z / pSG.VerticalExaggeration
    'Set pFlashPoint.SpatialReference = pSG.Scene.SpatialReference
    
    'Dim pDisplay As IDisplay3D
    'Set pDisplay = pSG
    'pDisplay.FlashLocation pFlashPoint
    
    Dim pPolyline As IPolyline
    Dim pGeomInput As IGeometry
    Dim pClone As IClone
    
    Dim iPntCount As Integer 'v0.12
908:     iPntCount = 0 'v0.12
    
910:     If chkSave2CSV.Value = 1 Then
911:         If txtCsvName.Text = Empty Then 'v0.12
912:             ShowMsgCritical "Warning", "Please select path and filename of new CSV file"
913:         End If 'v0.12
914:         Else: txtCsvName.Text = ""
915:     End If

    Dim pSurf As ISurface 'moved
918:     Set pSurf = MiscUtil.GetCurrentSurface(m_pApp) 'moved

        'User input
Dim iNLoS As Integer
Dim dRadius As Double
Dim iNVizEdge As Integer
Dim dBetaDeg As Double
Dim dBetaRad As Double
Dim dVisAlphaRad As Double
Dim dVisBetaRad As Double
Dim dAlphaInitDeg As Double
Dim dAlphaInitRad As Double
Dim dAlphaCoverDeg As Double
Dim dAlphaCoverRad As Double
Dim dSegVolTot As Double
Dim iNoObstrCount As Integer
        
Dim dVisVol(1 To 360) As Double 'rocky
Dim dinVisVol(1 To 360) As Double 'rocky
        
Dim PI As Double
939:  PI = 4 * Atn(1)
 
     'Array definitions
    Dim dVisLengthArr(400) As Double
    Dim dVisHeightArr(400) As Double
    Dim dInVisBlkArr(400, 200) As Double
    Dim dInVisRadArr(400, 200) As Double
    Dim dVisBetaRadArr(400) As Double

948:         If frmLineOfSight.GetiNLoS < 361 Then
949:             iNLoS = frmLineOfSight.GetiNLoS 'get this
950:         End If
951:         If (frmLineOfSight.GetRadius < 3501) And (frmLineOfSight.GetRadius > 51) Then _
        '08/06/04
953:         dRadius = frmLineOfSight.GetRadius 'get this
954:     End If
955:     If (frmLineOfSight.GetBetaDeg < 91) And (frmLineOfSight.GetBetaDeg > 1) Then
956:         dBetaDeg = frmLineOfSight.GetBetaDeg 'get this
957:     End If
958:     If (frmLineOfSight.GetAlphaInitDeg < 361) And (frmLineOfSight.GetAlphaInitDeg >= 0) _
        Then
960:         dAlphaInitDeg = frmLineOfSight.GetAlphaInitDeg
961:     End If
962:     If (frmLineOfSight.GetAlphaCoverDeg < 361) And (frmLineOfSight.GetAlphaCoverDeg >= _
        0) Then
964:         dAlphaCoverDeg = frmLineOfSight.GetAlphaCoverDeg
965:     End If
    
    
968:     dBetaRad = dBetaDeg * (PI / 180) 'check again
969:     dAlphaInitRad = dAlphaInitDeg * (PI / 180)
970:     dAlphaCoverRad = dAlphaCoverDeg * (PI / 180)
971:     dVisAlphaRad = (dAlphaCoverRad) / iNLoS '10/03/04 AlphaCover
    
973:     dSegVolTot = ((dAlphaCoverRad * (dRadius ^ 3) * Sin(dBetaRad)) / (3 * iNLoS)) _
    '10/03/04 AlphaCover
    
    'While-Wend input Point
977:     While Not pFeature Is Nothing
978:         Set pGeomInput = pFeature.Shape
        
980:         If (pGeomInput Is Nothing) Then
981:             Beep
            Exit Sub
983:         End If
        'If (Not pOwner Is MiscUtil.GetCurrentSurfaceLayer(m_pApp)) Then
        '    Beep
        '    Exit Sub
        'End If
        
989:         iPntCount = iPntCount + 1 'v0.12
        
        'Check if the process is stopped in Pause Msgbox
992:         If m_StopfrPause = True Then
993:             m_StopfrPause = False
            
995:             ShowMsgCritical "Warning", "Interrupted by user!"
996:             cmdStop.Visible = False
997:             CmdOK.Visible = True
            Exit Sub
999:         End If
        
1001:         Set pPnt = pGeomInput
        '        pPntCol.AddPoint pGeomInput
        
        'Start Routine
        ' Dim hpoints As Integer
        Dim i As Integer

           



    'For hpoints = 0 To (pPntCol.PointCount - 1)
1013:     For i = 1 To iNLoS
        '  If Not (TypeOf m_pApp Is IMxApplication) Then 'ArcScene calc starts
1015:         If (m_pScenePoints Is Nothing) Then '1
1016:             Set m_pScenePoints = New Polyline
            Dim pGeom As IGeometry
1018:             Set pGeom = m_pScenePoints
1019:             If (TypeOf m_pApp Is IMxApplication) Then
1020:                 Set pGeom.SpatialReference = m_pDispTrans.SpatialReference
1021:             End If
1022:             If (TypeOf m_pApp Is ISxApplication) Then
1023:                 Set pGeom.SpatialReference = pSG.Scene.SpatialReference
1024:             End If
1025:         End If '1
        
        '    pPntCol.QueryPoint hpoints, pPnt
1028:         m_pScenePoints.AddPoint pPnt
        
        'Circle
1031:         If (m_pScenePoints.PointCount = 1) Then '2
            
            Dim pCircArc As ICircularArc
1034:             Set pCircArc = New CircularArc
            
1036:             pCircArc.PutCoordsByAngle pPnt, dAlphaInitRad, dAlphaCoverRad, dRadius
            
            'Extract point
            Dim pToPoint As IPoint
1040:             Set pToPoint = New Point
            
1042:             pCircArc.QueryPoint esriNoExtension, i / iNLoS, True, pToPoint
            
1044:             If (Not pToPoint Is Nothing) Then
1045:                 m_pScenePoints.AddPoint pToPoint
1046:             End If
1047:         End If '2
        'End Circle
        
1050:         If (m_pScenePoints.PointCount = 2) Then '3
1051:             Set pClone = m_pScenePoints
1052:             Set pPolyline = pClone.Clone
1053:             Set m_pScenePoints = Nothing
1054:         Else
1055:             Set pPolyline = Nothing
1056:             Set m_pScenePoints = Nothing
1057:         End If '3
        '  End If
        
1060:         If (Not pPolyline Is Nothing) Then '(1) Not Polyline is Nothing
1061:             m_bFeedback_Start = False
1062:             Set m_pScenePoints = Nothing
            
'            Dim pSurf As ISurface 'moved
'             Set pSurf = MiscUtil.GetCurrentSurface(m_pApp) 'moved
            
            Dim fPoint As IPoint, tPoint As IPoint
1068:             Set fPoint = pPolyline.FromPoint
1069:             fPoint.Z = pSurf.GetElevation(fPoint)
1070:             Set tPoint = pPolyline.ToPoint
1071:             tPoint.Z = pSurf.GetElevation(tPoint)
            
1073:             If (pSurf.IsVoidZ(fPoint.Z) Or pSurf.IsVoidZ(tPoint.Z)) Then '(2)
1074:                 Beep
                Exit Sub
1076:             End If '(2)
            
1078:             fPoint.Z = fPoint.Z + frmLineOfSight.GetObserverOffset
1079:             tPoint.Z = tPoint.Z + frmLineOfSight.GetTargetOffset
            
            Dim pObstruct As IPoint
            Dim pVisPolyline As IPolyline
            Dim pInVisPolyline As IPolyline
            Dim bIsVis As Boolean
            Dim dVisColor As Double 'for changing color representation 09/03/04
            
1087:             pSurf.GetLineOfSight fPoint, tPoint, pObstruct, pVisPolyline, pInVisPolyline, _
            bIsVis, frmLineOfSight.GetCurvatureEnabled, frmLineOfSight.GetCurvatureEnabled
            
1090:             If (Not pObstruct Is Nothing) Then '(2) Not pObstruct is Nothing
                Dim dVisHeight As Double
                Dim dVisLength As Double
                '        Dim dVisVol As Double
                Dim dInVol As Double
                Dim dVisBetaRadMax As Double
                Dim dVisBetaDegMax As Double
                Dim dVisBetaDeg As Double
                '        Dim dInVisVol As Double
                Dim dInVisBlk As Double
                Dim dVisSurf As Double 'Visible surface 15/03/04
                Dim dTotSurf As Double 'Total Visible surface 15/03/04
                
                Dim dTotVol As Double
                Dim dTotInVol As Double
                Dim dObstRadMax As Double
                Dim dObstRadTot As Double
                Dim dObstRadMin As Double
                Dim pDeltaObst As IVector3D
1109:                 Set pDeltaObst = New Vector3D
                Dim pInVisGeomColl As IGeometryCollection
                Dim pInVisPointColl As IPointCollection
                Dim pInVisPntFirst As IPoint
                Dim pPrevObstruct As IPoint
                
                'Array calculation for Ave Obstruction Radius
                'Dim dVisLengthArr(101) As Double
                'Dim dVisHeightArr(101) As Double
                
                '        dVisHeight = pObstruct.Z '- fPoint.Z '05/03/04
1120:                 If (Not pVisPolyline Is Nothing) Then '(3)
                    'Calculation if Length = sqr(x2+y2)
                    'If (pVisPolyline.length > dVisHeight) Then
                    'dVisLength = Sqr(Abs((pVisPolyline.length) ^ 2) - (dVisHeight ^ 2))
                    'dVisLengthArr(i) = dVisLength
                    'End If
                    'Else
                    'New calculation based on distance between 2D pObstruct and fPoint
                    '            Set pClone = pObstruct '05/03/04
                    '            pObstruct.Z = pSurf.GetElevation(pObstruct) '05/03/04
                    '            fPoint.Z = 0 '05/03/04
1131:                     If (Not pInVisPolyline Is Nothing) Then '(4) 05/03/04
1132:                         Set pInVisGeomColl = pInVisPolyline
1133:                         Set pInVisPointColl = pInVisGeomColl.Geometry(0)
1134:                         Set pInVisPntFirst = pInVisPointColl.Point(0)
1135:                     End If '(4) 05/03/04
                    
                    'Reassessment of Obstruction point '28/05/04
1138:                     If (Not pInVisPntFirst Is Nothing) Then
1139:                         If pObstruct.Z >= pInVisPntFirst.Z Then '28/05/04
1140:                             Set pInVisPntFirst = pObstruct
1141:                         End If '28/05/04
                        'Visual Edge
1143:                         If pInVisPntFirst.Z > fPoint.Z Then '24/05/04
1144:                             iNVizEdge = iNVizEdge + 1
1145:                         End If '24/05/04
1146:                     End If '28/05/04
                    
1148:                     pDeltaObst.ConstructDifference pInVisPntFirst, fPoint '05/03/04
1149:                     dVisHeight = pDeltaObst.ZComponent '05/03/04
1150:                     dVisLength = Sqr((pDeltaObst.XComponent ^ 2) + (pDeltaObst.YComponent ^ 2)) '05/03/04
                    '        dVisLength = pVisPolyline.length 'Previous calculation
1152:                     dVisHeightArr(i) = dVisHeight '05/03/04
1153:                     dVisLengthArr(i) = dVisLength
                    '        Set pObstruct = pClone.Clone '05/03/04
1155:                     dVisBetaRad = pDeltaObst.Inclination
1156:                     dVisBetaRadArr(i) = dVisBetaRad
                    'Maximum Beta 02/03/04
1158:                     dVisBetaDeg = dVisBetaRadArr(i) * (180 / PI) '02/03/04
                    
1160:                     If dVisBetaRadMax < dVisBetaRad Then '(4)
1161:                         dVisBetaRadMax = dVisBetaRad
1162:                     End If '(4)
                    
                    'Cumulative area
                    'dVisArea = (dVisHeight * dVisLength) / 2
1166:                     dVisVol(i) = dVisHeight * (dVisLength ^ 2) * Sin(dVisAlphaRad) / 3
1167:                     dinVisVol(i) = 0 'if no invisible volume
                    
                    'Count Invisible Polyline before pInVisPntFirst
                    Dim k As Integer
                    Dim l(400) As Integer
                    Dim pGeomCollInVis As IGeometryCollection
1173:                     Set pGeomCollInVis = pVisPolyline '06/03/04 change from pInVisPolyline
                    Dim pPCollInVis As IPointCollection
                    Dim pPInVisCounter As IPoint
                    
                    'Major change to use pVisPolyline alone 06/03/04
1178:                     If (Not pGeomCollInVis.GeometryCount = 0) Then '(4) Check InVis Point 06/03/04
1179:                         Set pPCollInVis = pGeomCollInVis '06/03/04 from inside loop
1180:                         For k = 0 To (pPCollInVis.PointCount - 1)
                            'Dim pClone As IClone
                            '            Set pClone = pPCollInVis.Point(k)
                            '            Set pPInVisCounter = pClone.Clone
                            '            If pPInVisCounter.Z <= pObstruct.Z Then
                            'Set pPInVisCounter.Geometry(k) = pGeomCollInVis.Geometry(k)
1186:                             If (l(i) < pPCollInVis.PointCount) Then '06/03/04
1187:                                 l(i) = l(i) + 1
1188:                             End If
                            '            End If
1190:                         Next k
                        
                        'Invisible volume
                        Dim j As Integer
                        Dim pPOrigin As IPoint
1195:                         Set pPOrigin = fPoint
                        
1197:                         If (Not l(i) = 0) And (l(i) < pPCollInVis.PointCount) Then '(5) Check InVis Point _
                            06/03/04
1199:                             For j = 1 To l(i) '06/03/04
                                Dim dInVisRad As Double
                                Dim dInVisL As Double
                                Dim dInVisH As Double
                                Dim dInVisX As Double
                                Dim pPInVis2 As IPoint
                                Dim pPInVis1 As IPoint
                                '            Set pPCollInVis = pGeomCollInVis '06/03/04
1207:                                 If (j > 0) Then '(6)
                                    
                                    Dim pDeltaX As IVector3D
1210:                                     Set pDeltaX = New Vector3D
                                    Dim pDeltaRad As IVector3D
1212:                                     Set pDeltaRad = New Vector3D
1213:                                     Set pClone = pPCollInVis.Point(j)
1214:                                     Set pPInVis2 = pClone.Clone
1215:                                     Set pClone = pPCollInVis.Point(j - 1)
1216:                                     Set pPInVis1 = pClone.Clone
                                    'dInVisX = segment depth, dInVisL = segment width
1218:                                     dInVisH = pPInVis1.Z
                                    'pPInVis1.Z = 0 'can be neglected '15/04/04
                                    'pPInVis2.Z = 0 'can be neglected 15/04/04
1221:                                     pPOrigin.Z = 0
1222:                                     pDeltaX.ConstructDifference pPInVis1, pPInVis2
1223:                                     pDeltaRad.ConstructDifference pPInVis1, pPOrigin
1224:                                     dInVisX = Sqr((pDeltaX.XComponent ^ 2) + (pDeltaX.YComponent ^ 2)) '15/04/04 _
                                    originally pDeltaX.Magnitude
1226:                                     dInVisRad = Sqr((pDeltaRad.XComponent ^ 2) + (pDeltaRad.YComponent ^ 2)) '15/04/04 _
                                    originally pDeltaRad.Magnitude
1228:                                     dInVisRadArr(i, j) = dInVisRad
1229:                                     dInVisL = 2 * dInVisRad * Sin(dVisAlphaRad / 2)
1230:                                     If dInVisRad < dVisLength Then '(7)
1231:                                         dInVisBlk = dInVisL * dInVisH * dInVisX
1232:                                         dInVisBlkArr(i, j) = dInVisBlk
                                        'Surface calculation 15/04/04
1234:                                         If (pPInVis2.Z > pPInVis1.Z) Then '(8) 15/04/04
1235:                                             dVisSurf = (pDeltaX.Magnitude * dInVisL)
1236:                                         Else
1237:                                             dVisSurf = 0 '15/04/04
1238:                                         End If '(8) 15/04/04
1239:                                     Else '(7)
1240:                                         dInVisBlkArr(i, j) = 0
1241:                                         dVisSurf = 0
1242:                                     End If '(7)
                                    
1244:                                 End If '(6)
                                
1246:                                 dinVisVol(i) = dinVisVol(i) + dInVisBlk
1247:                                 dInVisBlk = 0
1248:                                 dTotSurf = dTotSurf + dVisSurf '15/04/04
1249:                             Next j
1250:                         End If '(5) Check InVis Point 06/03/04
1251:                     End If '(4) Check InVis Point 06/03/04
                    
1253:                     If (dVisVol(i) > dSegVolTot) Then '(4) Visible volume check '10/03/04
1254:                         dVisVol(i) = dSegVolTot
1255:                     End If '(4)
1256:                     If (dinVisVol(i) >= dVisVol(i)) Then '(4) Invisible volume check '06/03/04
1257:                         dinVisVol(i) = 0
1258:                     End If '(4)
1259:                     dTotVol = dTotVol + (dVisVol(i) - dinVisVol(i))
                    
                    'Invisible occluded volume 24/05/04
1262:                     If (dInVol <= ((dVisAlphaRad * (dRadius ^ 3) * Sin(dVisBetaRadArr(i))) / 3)) Then _
                    '24/05/04
1264:                     dInVol = ((dVisAlphaRad * (dRadius ^ 3) * Sin(dVisBetaRadArr(i))) / 3) - (dVisVol(i) _
                    - dinVisVol(i)) '24/05/04
1266:                 Else
1267:                     dInVol = ((dVisAlphaRad * (dRadius ^ 3) * Sin(dVisBetaRadArr(i))) / 3)
1268:                 End If
1269:                 dTotInVol = dTotInVol + dInVol
                
1271:                 If (dTotSurf < (2 * dVisLength * Sin(dVisAlphaRad / 2) * dVisHeight)) Then '(4) _
                    Visible surface check 15/04/04
1273:                     dTotSurf = (2 * dVisLength * Sin(dVisAlphaRad / 2) * dVisHeight)
1274:                 End If '(4) Visible surface check 15/04/04
                
1276:                 dVisColor = ((255 * dVisVol(i)) / dSegVolTot) '09/03/04
                
                'dVisVol(i) = 0
                'dinVisVol(i) = 0
1280:                 dInVol = 0
                
                'Max Obstruction Radius
1283:                 If (dObstRadMax < dVisLength) Then '(4)
1284:                     dObstRadMax = dVisLength
1285:                 End If '(4)
                'Min Obstruction Radius
1287:                 If (dObstRadMin = 0) Then '(4) 19/06/04
1288:                     dObstRadMin = dVisLength
1289:                 End If '(4) 19/06/04
1290:                 If ((dObstRadMin > dVisLength) And (Not dVisLength = 0)) Then '(4)
                    'If ((dVisLength + 1) > 1) And (dVisHeight > 10) Then '(5)
1292:                     dObstRadMin = dVisLength
                    'End If '(5)
1294:                 End If '(4)
1295:             End If '(3)
            'Total Obstruction Radius
1297:             dObstRadTot = dObstRadTot + dVisLength
1298:             dVisLength = 0
            
            'Perimeter calculation '08/03/04
            
            Dim dSegPerimeter As Double
            Dim dTotPerimeter As Double
            Dim pDeltaPeri As IVector3D
1305:             Set pDeltaPeri = New Vector3D
            
            '    If (pInVisPntFirst Is Nothing) Then '(3)
            '        Set pPrevObstruct = tPoint
            '    Else '(3)
            '        Set pPrevObstruct = pInVisPntFirst
            '    End If '(3)
1312:             If (Not pPrevObstruct Is Nothing) And (Not i = 1) Then '(3) 28/05/04
1313:                 If (pInVisPntFirst Is Nothing) Then '(4)
1314:                     pDeltaPeri.ConstructDifference tPoint, pPrevObstruct
1315:                     Set pPrevObstruct = tPoint
1316:                 Else '(4)
1317:                     pDeltaPeri.ConstructDifference pInVisPntFirst, pPrevObstruct
1318:                     Set pPrevObstruct = pInVisPntFirst
1319:                 End If '(4)
1320:                 dSegPerimeter = Sqr((pDeltaPeri.XComponent ^ 2) + (pDeltaPeri.YComponent ^ 2)) _
                '08/03/04
1322:             Else '(3)
1323:                 If (pInVisPntFirst Is Nothing) Then '(4)
1324:                     Set pPrevObstruct = tPoint
1325:                 Else '(4)
1326:                     Set pPrevObstruct = pInVisPntFirst
1327:                 End If '(4)
1328:             End If '(3)
            
1330:             dTotPerimeter = dTotPerimeter + dSegPerimeter
1331:             dSegPerimeter = 0
            
1333:         Else '(2) Not pObstruct is Nothing
            
1335:             dVisBetaRadArr(i) = 0
1336:             iNoObstrCount = iNoObstrCount + 1
            
1338:         End If '(2) Not pObstruct is Nothing
        'End Perimeter calculation 08/03/04
        
        Dim pSym As ISimpleLineSymbol
1342:         Set pSym = New SimpleLineSymbol
        Dim pColor As IRgbColor
1344:         Set pColor = New RgbColor
1345:         pSym.Width = 2
1346:         pSym.Style = esriSLSSolid
        
        Dim pFillSym As ISimpleFillSymbol
1349:         Set pFillSym = New SimpleFillSymbol
        
        Dim bVisAdded As Boolean
1352:         bVisAdded = False
        
1354:         If (TypeOf m_pApp Is IMxApplication) Then '(2)
1355:             If frmLineOfSight.chkViz = 1 Then '(3) Graphic optional visualization
1356:                 If (Not pVisPolyline Is Nothing) Then '(4)
1357:                     pColor.Green = 255
1358:                     pSym.Color = pColor
1359:                     AddGraphic m_pApp, pVisPolyline, pSym
1360:                     bVisAdded = True
1361:                 End If '(4)
                
1363:                 If (Not pInVisPolyline Is Nothing) Then '(4)
1364:                     pColor.Green = 0
1365:                     pColor.Red = 255
1366:                     pSym.Color = pColor
                    'AddGraphic m_pApp, pInVisPolyline, pSym, bVisAdded '15/04/04 No red lines
1368:                     If (bVisAdded) Then '(5)
1369:                         MiscUtil.GroupSelectedGraphics m_pApp
1370:                     End If '(5)
1371:                 End If '(4)
1372:             End If '(3) Graphic optional visualization
1373:         End If '(2) Mx Application 10/03/04
        
1375:         If (TypeOf m_pApp Is ISxApplication) Then '(2)
            Dim pVisPatch As IMultiPatch
1377:             Set pVisPatch = New MultiPatch
1378:             Set pGeom = pVisPatch
1379:             Set pGeom.SpatialReference = pSG.Scene.SpatialReference
            Dim pInVisPatch As IMultiPatch
1381:             Set pInVisPatch = New MultiPatch
1382:             Set pGeom = pInVisPatch
1383:             Set pGeom.SpatialReference = pSG.Scene.SpatialReference
            Dim dTargetHeightForVis As Double
1385:             If frmLineOfSight.chkViz = 1 Then       '(3) Graphic optional visualization
1386:                 geomutil.CreateVerticalLOSPatches bIsVis, fPoint, tPoint, pVisPolyline, _
                pInVisPolyline, pVisPatch, pInVisPatch, dTargetHeightForVis
1388:                 If (Not pVisPatch.IsEmpty) Then '(4)
                    'Graphic color representing Visible Beta 09/03/04
1390:                     If (dVisBetaDeg > 45) Then '(5)
1391:                         pColor.Green = 0
1392:                         pColor.Red = 255
                        'pColor.Transparency = 255
1394:                         pFillSym.Color = pColor
1395:                         ElseIf (dVisBetaDeg > 27) Then '(5)10/03/04
1396:                         pColor.Green = 255
1397:                         pColor.Red = 255
1398:                         pFillSym.Color = pColor
1399:                     Else '(5)
1400:                         pColor.Green = 255 '- (10 * dVisColor) '09/03/04
1401:                         pColor.Red = 0 '+ (10 * dVisColor) '09/03/04
                        'pColor.Transparency = 120 '09/03/04
1403:                         pFillSym.Color = pColor
1404:                     End If '(5)
1405:                     AddGraphic m_pApp, pVisPatch, pFillSym, False, False
1406:                 End If '(4)
1407:                 If (Not pInVisPatch.IsEmpty) Then '(4)
1408:                     pColor.Green = 0
1409:                     pColor.Red = 255
1410:                     pFillSym.Color = pColor
                    'AddGraphic m_pApp, pInVisPatch, pFillSym, False, False 'if viewsphere then off
1412:                 End If '(4)
1413:             End If '(3)
1414:         End If '(2) Sx application 10/03/04
        
        'Msgbox check obstruction degree '05/03/04
1417:         If frmLineOfSight.chkShowSeg = 1 Then '(2) Optional Show Segments 06/03/04
1418:             If (Not pObstruct Is Nothing) Then '(3) Not pObstruct is Nothing2 06/03/04
                Dim sfPointX As String
                Dim sfPointY As String
                Dim sfPointZ As String
                Dim sObstructX As String
                Dim sObstructY As String
                Dim sObstructZ As String
                Dim sVisLengthArr As String
                Dim sVisHeightArr As String
                Dim sVisBetaDeg As String
1428:                 sfPointX = Format(fPoint.x, "####.####")
1429:                 sfPointY = Format(fPoint.y, "####.####")
1430:                 sfPointZ = Format(fPoint.Z, "####.####")
1431:                 sObstructX = Format(pInVisPntFirst.x, "####.####")
1432:                 sObstructY = Format(pInVisPntFirst.y, "####.####")
1433:                 sObstructZ = Format(pInVisPntFirst.Z, "####.####")
1434:                 sVisHeightArr = Format(dVisHeightArr(i), "###.######")
1435:                 sVisLengthArr = Format(dVisLengthArr(i), "###.######")
1436:                 sVisBetaDeg = Format(dVisBetaDeg, "###.####")
                '        If frmLineOfSight.chkShowMsg = 1 Then 'Optional Show Message Box
1438:                 ShowMsgInfo "Data per Segment", ("Observ = (" & sfPointX & "," & sfPointY & "," & _
                sfPointZ & ")" & (Chr(13)) & _
                " ; Obstr = (" & sObstructX & "," & sObstructY & "," & sObstructZ & ")" & (Chr(13)) _
                & _
                " ; Height = " & sVisHeightArr & (Chr(13)) & _
                " ; Length = " & sVisLengthArr & (Chr(13)) & _
                " ; Beta deg = " & sVisBetaDeg)
                '        End If
1446:             End If '(3) Not pObstruct is Nothing2 06/03/04
1447:         End If '(2) Optional Show Segments 06/03/04
        'End Msgbox check 05/03/04
        
        ' notify doc it's been changed - so it knows to ask the user if they want to save on _
        exit
        Dim pDoc As IBasicDocument
1453:         Set pDoc = m_pApp.Document
1454:         pDoc.UpdateContents
        
1456:         DoEvents ' since the added graphic gets selected the doc sends a message to the _
        status bar and that
        ' won't happen until after this routine returns - effectively overwriting this tool's
        ' message (see next section). DoEvents flushes the doc's message so the overwrite _
        problem
        ' goes away.
        
        ' Write results to status bar
        '    If (bIsVis) Then
        '      m_pApp.StatusBar.Message(0) = "Target is visible"
        '    Else
        '      If (TypeOf m_pApp Is IMxApplication) Then
        'm_pApp.StatusBar.Message(0) = "Target is not visible"
        '      Else
        '        Dim sTarZ As String
        '        sTarZ = Format(tPoint.Z, "#.##")
        '        Dim sTarVisZ As String
        '        sTarVisZ = Format(dTargetHeightForVis, "#.##")
        '        'm_pApp.StatusBar.Message(0) = "Target is not visible (z = " & sTarZ & ", z _
        required: " & sTarVisZ & ")"
        '      End If
        '    End If
        
1479:     End If '(1) Not Polyline is Nothing - ArcScene calc ends
1480: Next i

'Write calculation
Dim sTarVizVol As String
Dim sTotSurf As String
Dim sUserRadius As String
Dim sVizEdge As String
Dim sViewsphere As String
Dim sViewsphereSTD As String 'rocky
Dim dViewsphere As Double
Dim dViewsphereSTD As Double 'rocky
Dim sObstRadMax As String
Dim sObstRadAve As String
Dim sObstRadMin As String
Dim dObstRadAve As Double
Dim dVSImax As Double
Dim dVSImaxSTD As Double 'rocky
Dim dVSIave As Double
Dim dVSIaveSTD As Double 'rocky
Dim dVSImin As Double
Dim dVSIminSTD As Double 'rocky
Dim sVSImax As String
Dim sVSImaxSTD As String 'rocky
Dim sVSIave As String
Dim sVSIaveSTD As String 'rocky
Dim sVSImin As String
Dim sVSIminSTD As String 'rocky
Dim sSVF As String
Dim sSVFSTD As String 'rocky
'24/05/04
Dim sTotInVol As String
Dim sInVSI As String
Dim sInVSImax As String
Dim sInVSIave As String
Dim sOcclRatMax As String
Dim sOcclRatAve As String

Dim sBetaDegMax As String
Dim sBetaDegAve As String
Dim sDistanceHeightMax As String
Dim sDistanceHeightAve As String
Dim sVisArea As String
Dim sPerimeter As String
Dim sLongestAxis As String
Dim sCompact As String
Dim sFractal As String
Dim sConvolution As String
Dim sPattonDiv As String
Dim sElliptic As String

Dim sEnclosFull As String
Dim sEnclosThres As String
Dim sEnclosMin As String
Dim sEnclosLoose As String
Dim sTypology As String

Dim dCheck As Double
Dim sCheck As String

'    dCheck = Sin(90)
'    sCheck = Format(dCheck, "#.##")

'Array calculation for VSIave
1543: dObstRadAve = dObstRadTot / iNLoS
1544: If dObstRadMin = 0 Then
1545:     dObstRadMin = 1
1546: End If

Dim dVizEdgePerc As Double
Dim dVisVolMax(1 To 360) As Double
Dim dInVisVolMax(1 To 360) As Double
Dim dTotVolMax As Double
Dim dSegVolMax As Double
Dim dVisVolAve(1 To 360) As Double
Dim dInVisVolAve(1 To 360) As Double
Dim dTotVolAve As Double
Dim dSegVolAve As Double
Dim dVisVolMin(1 To 360) As Double
Dim dTotVolMin As Double
Dim dSegVolMin As Double
Dim dSegSVF(1 To 360) As Double
Dim dTotSVF As Double
Dim dSVF As Double
Dim dSVFSTD As Double 'rocky
'24/05/04
Dim dInVSI As Double
Dim dTotInVolMax As Double
Dim dInVolMax As Double
Dim dInVSImax As Double
Dim dOcclRatioMax As Double
Dim dTotInVolAve As Double
Dim dInVolAve As Double
Dim dInVSIave As Double
Dim dOcclRatioAve As Double

Dim dTotBetaRad As Double
Dim dVisBetaDegAve As Double
Dim dVisArea As Double
Dim dTotVisArea As Double
Dim dCompact As Double
Dim dFractal As Double
Dim dLongAxis As Double
Dim dLongestAxis As Double
Dim dConvolution As Double
Dim dPattonDiv As Double
Dim dElliptic As Double

Dim dEnclosureFull As Double
Dim dEnclosureThreshold As Double
Dim dEnclosureMin As Double
Dim dEnclosureLoose As Double
Dim dDistanceHeightMax As Double
Dim dDistanceHeightAve As Double
Dim dRatioMinMax As Double
Dim dRatioAveMax As Double
Dim dRatioMinAve As Double

'Start calculation
1598: dSegVolMax = ((dAlphaCoverRad * (dObstRadMax ^ 3) * Sin(dBetaRad)) / (3 * iNLoS)) _
'28/05/04
1600: dSegVolAve = ((dAlphaCoverRad * (dObstRadAve ^ 3) * Sin(dBetaRad)) / (3 * iNLoS)) _
'10/03/04
1602: dSegVolMin = ((dAlphaCoverRad * (dObstRadMin ^ 3) * Sin(dBetaRad)) / (3 * iNLoS)) _
'10/03/04

1605: If (dSegVolMax = 0) Or (dSegVolAve = 0) Or (dSegVolMin = 0) Then 'Check against _
    overflow '10/03/04
1607:     dSegVolMax = dSegVolTot '28/05/04
1608:     dSegVolAve = dSegVolTot
1609:     dSegVolMin = dSegVolTot
1610: End If

1612: If (iNoObstrCount < iNLoS) Then '(1) 10/03/04
    'Main VSI calculation
1614:     For i = 1 To iNLoS
        'Check if Stop button has been pressed in the middle of this process 'v0.13
1616:         If m_Stop = True Then
1617:             m_Stop = False
1618:             ShowMsgCritical "Warning", "Interrupted by user!"
1619:             cmdStop.Visible = False
1620:             CmdOK.Visible = True
            Exit Sub
1622:         End If
        'VSImax '28/05/04
1624:         If (dVisLengthArr(i) < dObstRadMax) Then '(2)
1625:             dVisVolMax(i) = (dVisHeightArr(i) * (dVisLengthArr(i) ^ 2) * Sin(dVisAlphaRad)) / 3
1626:         Else '(2)
1627:             dVisVolMax(i) = ((dVisAlphaRad) * (dObstRadMax ^ 3) * Sin(dVisBetaRadArr(i))) / 3 _
            '10/03/04
1629:         End If '(2)
        
        'VSIave
1632:         If (dVisLengthArr(i) < dObstRadAve) Then '(2)
1633:             dVisVolAve(i) = (dVisHeightArr(i) * (dVisLengthArr(i) ^ 2) * Sin(dVisAlphaRad)) / 3
1634:         Else '(2)
1635:             dVisVolAve(i) = ((dVisAlphaRad) * (dObstRadAve ^ 3) * Sin(dVisBetaRadArr(i))) / 3 _
            '10/03/04
1637:         End If '(2)
        
        'Invisible part of VSImax & VSIave
1640:         If (Not j = 0) Then '(2) 28/05/04
1641:             For j = 1 To l(i) 'Invisible volume calculation for VSImax & VSIave 28/05/04
1642:                 If dInVisRadArr(i, j) < dObstRadMax Then '(3)
1643:                     dInVisVolMax(i) = dInVisVolMax(i) + dInVisBlkArr(i, j)
1644:                 End If '(3)
1645:                 If dInVisRadArr(i, j) < dObstRadAve Then '(3)
1646:                     dInVisVolAve(i) = dInVisVolAve(i) + dInVisBlkArr(i, j)
1647:                 End If '(3)
1648:             Next j
1649:         End If '(2) 06/03/04
        
        'VSImax check
1652:         If (dVisVolMax(i) > dSegVolMax) Then '(2) Vis vol check '28/05/04
1653:             dVisVolMax(i) = dSegVolMax
1654:         End If '(2)
1655:         If (dInVisVolMax(i) >= dVisVolMax(i)) Then '(2) Invis vol check '28/05/04
1656:             dInVisVolMax(i) = 0
1657:         End If '(2)
1658:         dTotVolMax = dTotVolMax + (dVisVolMax(i) - dInVisVolMax(i))
        
        'Invisible occluded volume average 24/05/04
1661:         dInVolMax = ((dVisAlphaRad * (dObstRadMax ^ 3) * Sin(dVisBetaRadArr(i))) / 3) - _
        (dVisVolMax(i) - dInVisVolMax(i)) '24/05/04
1663:         dTotInVolMax = dTotInVolMax + dInVolMax
        'dVisVolMax(i) = 0
        'dInVisVolMax(i) = 0
1666:         dInVolMax = 0
        
        'VSIave check
1669:         If (dVisVolAve(i) > dSegVolAve) Then '(2) Vis vol check '10/03/04
1670:             dVisVolAve(i) = dSegVolAve
1671:         End If '(2)
1672:         If (dInVisVolAve(i) >= dVisVolAve(i)) Then '(2) Invis vol check '06/03/04
1673:             dInVisVolAve(i) = 0
1674:         End If '(2)
1675:         dTotVolAve = dTotVolAve + (dVisVolAve(i) - dInVisVolAve(i))
        
        'Invisible occluded volume average 24/05/04
1678:         dInVolAve = ((dVisAlphaRad * (dObstRadAve ^ 3) * Sin(dVisBetaRadArr(i))) / 3) - _
        (dVisVolAve(i) - dInVisVolAve(i)) '24/05/04
1680:         dTotInVolAve = dTotInVolAve + dInVolAve
        'dVisVolAve(i) = 0
        'dInVisVolAve(i) = 0
1683:         dInVolAve = 0
        
        'VSImin
1686:         If (dVisLengthArr(i) > 0) Then '(2)
1687:             dVisVolMin(i) = ((dVisAlphaRad) * (dObstRadMin ^ 3) * Sin(dVisBetaRadArr(i))) / 3 _
            '10/03/04
1689:         End If '(2)
        
1691:         If (dVisVolMin(i) > dSegVolMin) Then '(2) Vis vol check '10/03/04
1692:             dVisVolMin(i) = dSegVolMin
1693:         End If '(2)
        
1695:         dTotVolMin = dTotVolMin + dVisVolMin(i)
        'dVisVolMin(i) = 0
        
        'Sky View Factor
1699:         dSegSVF(i) = (1 - Sin(dVisBetaRadArr(i)))
1700:         dTotSVF = dTotSVF + dSegSVF(i)
        'dSegSVF(i) = 0
        
        'Average Visible Beta 08/03/04
1704:         dTotBetaRad = dTotBetaRad + dVisBetaRadArr(i)
        
        'Area of visible urban space 08/03/04
1707:         If (dVisLengthArr(i) > 0) Then '(2)
1708:             dVisArea = (dVisAlphaRad * (dVisLengthArr(i) ^ 2)) / 2 '08/03/04
1709:         Else '(2)
1710:             dVisArea = (dVisAlphaRad * (dRadius ^ 2)) / 2 '08/03/04
1711:         End If '(2)
1712:         dTotVisArea = dTotVisArea + dVisArea '08/03/04
1713:         dVisArea = 0
        
        'Longest Axis '08/03/04
1716:         If (dAlphaCoverDeg >= 180) And ((i - (0.5 * iNLoS)) > 0) Then '(2)
1717:             If dVisLengthArr(i) > 0 Then '(3)
1718:                 If dVisLengthArr(i - (0.5 * iNLoS)) > 0 Then '(4)
1719:                     dLongAxis = dVisLengthArr(i) + dVisLengthArr(i - (0.5 * iNLoS))
1720:                 Else '(4)
1721:                     dLongAxis = dVisLengthArr(i) + dRadius
1722:                 End If '(4)
1723:             Else '(3)
1724:                 If dVisLengthArr(i - (0.5 * iNLoS)) > 0 Then
1725:                     dLongAxis = dRadius + dVisLengthArr(i - (0.5 * iNLoS))
1726:                 Else
1727:                     dLongAxis = dRadius + dRadius
1728:                 End If '(4)
1729:             End If '(3)
1730:         End If '(2)
1731:         If (dAlphaCoverDeg < 180) Then '(2)
1732:             If dVisLengthArr(i) > 0 Then
1733:                 dLongAxis = dVisLengthArr(i)
1734:             Else
1735:                 dLongAxis = dRadius
1736:             End If
1737:         End If '(2)
1738:         If dLongestAxis < dLongAxis Then '(2)
1739:             dLongestAxis = dLongAxis
1740:         End If '(2)
        'End Longest Axis '08/03/04
        
        'Degree of Enclosure by Spreiregen, Lynch, Ashihara '09/03/04
1744:         dVisBetaDeg = dVisBetaRadArr(i) * (180 / PI)
1745:         If (dVisBetaDeg > 45) Then '(2)
1746:             dEnclosureFull = dEnclosureFull + (dVisAlphaRad * (180 / PI))
1747:             ElseIf (dVisBetaDeg > 27) Then '(2)
1748:             dEnclosureThreshold = dEnclosureThreshold + (dVisAlphaRad * (180 / PI))
1749:             ElseIf (dVisBetaDeg > 18) Then '(2)
1750:             dEnclosureMin = dEnclosureMin + (dVisAlphaRad * (180 / PI))
1751:         Else '(2)
1752:             dEnclosureLoose = dEnclosureLoose + (dVisAlphaRad * (180 / PI))
1753:         End If '(2)
1754:     Next i
    'End of Main VSI calculation
    
1757:     If (Not iNVizEdge > iNLoS) Then '(2)
1758:         dVizEdgePerc = (iNVizEdge / iNLoS) * 100 '24/05/04
1759:     End If '(2)
    'Check Longest Axis 28/05/04
1761:     If (dVizEdgePerc >= 100) Then
1762:         If (dLongestAxis > (2 * dObstRadMax)) Then
1763:             dLongestAxis = 2 * dObstRadMax
1764:         End If
1765:     Else
1766:         If (dLongestAxis > (2 * dRadius)) Then
1767:             dLongestAxis = 2 * dRadius
1768:         End If
1769:     End If
    
    'dViewsphere = dTotArea / (iNLoS * 0.25 * PI * (dRadius ^ 2)) 'segment area
    'dViewsphere = dTotVol / ((2 * PI * (dRadius ^ 3)) / 3) 'full hemisphere
1773:     dViewsphere = dTotVol / ((dAlphaCoverRad * (dRadius ^ 3) * Sin(dBetaRad)) / 3) 'ok _
    10/03/04
    'dVSImax = dTotArea / (iNLoS * 0.25 * PI * (dObstRadMax ^ 2))
    'dVSImax = dTotVol / ((2 * PI * (dObstRadMax ^ 3)) / 3)
1777:     dInVSI = dTotInVol / ((dAlphaCoverRad * (dRadius ^ 3) * Sin(dBetaRad)) / 3) '24/05/04
    
1779:     If (dObstRadMax > 0) And (dObstRadAve > 0) And (dObstRadMin > 0) Then '(2)
        'MAX 'latest 28/05/04
1781:         dVSImax = dTotVolMax / ((dAlphaCoverRad * (dObstRadMax ^ 3) * Sin(dBetaRad)) / 3) _
        'ok 10/03/04
1783:         dInVSImax = dTotInVolMax / ((dAlphaCoverRad * (dObstRadMax ^ 3) * Sin(dBetaRad)) / _
        3) '24/05/04
1785:         dOcclRatioMax = dTotInVolMax / dTotVolMax '24/05/04
        'dVSIave = dTotVolArea / (iNLoS * 0.25 * PI * (dObstRadAve ^ 2))
        'dVSIave = dTotVolAve / ((2 * PI * (dObstRadAve ^ 3)) / 3)
        
        'AVE
1790:         dVSIave = dTotVolAve / ((dAlphaCoverRad * (dObstRadAve ^ 3) * Sin(dBetaRad)) / 3) _
        'ok 10/03/04
1792:         dInVSIave = dTotInVolAve / ((dAlphaCoverRad * (dObstRadAve ^ 3) * Sin(dBetaRad)) / _
        3) '24/05/04
1794:         dOcclRatioAve = dTotInVolAve / dTotVolAve '24/05/04
        
        'MIN
1797:         dVSImin = dTotVolMin / ((dAlphaCoverRad * (dObstRadMin ^ 3) * Sin(dBetaRad)) / 3) _
        'ok 10/03/04
        
1800:     End If '(2)
    
1802:     dSVF = dTotSVF / iNLoS
1803:     dVisBetaDegMax = dVisBetaRadMax * (180 / PI) '08/03/04
1804:     dVisBetaDegAve = (dTotBetaRad / iNLoS) * (180 / PI) '08/03/04
    'Distance-Height proportion
1806:     dDistanceHeightMax = 1 / (Tan(dVisBetaRadMax)) '08/04/04
1807:     dDistanceHeightAve = 1 / (Tan(dTotBetaRad / iNLoS)) '08/04/04
    
    
    Dim p1 As Integer
    Dim p2 As Integer
    Dim p3 As Integer
    Dim p4 As Integer
    Dim p5 As Integer
    
    'rocky
1817:     dViewsphereSTD = 0
1818:     For p1 = 1 To iNLoS
1819:         dViewsphereSTD = dViewsphereSTD + (((dVisVol(p1) - dinVisVol(p1)) - (dTotVol / _
        iNLoS)) ^ 2) / ((iNLoS - 1) * ((dAlphaCoverRad * (dRadius ^ 3) * Sin(dBetaRad)) / 3) _
^ 2)
        'dViewsphereSTD = dViewsphereSTD + (((dVisVol(p1) - dinVisVol(p1)) - (dTotVol / _
        iNLoS)) ^ 2) / (iNLoS - 1)
1824:     Next p1
1825:     dViewsphereSTD = dViewsphereSTD ^ (1 / 2)
    
1827:     dVSImaxSTD = 0
1828:     For p2 = 1 To iNLoS
1829:         dVSImaxSTD = dVSImaxSTD + ((dVisVolMax(p2) - dInVisVolMax(p2)) - (dTotVolMax / _
        iNLoS)) ^ 2 / ((iNLoS - 1) * ((dAlphaCoverRad * (dObstRadMax ^ 3) * Sin(dBetaRad)) / _
3) ^ 2)
        'dVSImaxSTD = dVSImaxSTD + ((dVisVolMax(p2) - dInVisVolMax(p2)) - (dTotVolMax / _
        iNLoS)) ^ 2 / (iNLoS - 1)
1834:     Next p2
1835:     dVSImaxSTD = dVSImaxSTD ^ (1 / 2)
    
1837:     dVSIaveSTD = 0
1838:     For p3 = 1 To iNLoS
1839:         dVSIaveSTD = dVSIaveSTD + ((dVisVolAve(p3) - dInVisVolAve(p3)) - (dTotVolAve / _
        iNLoS)) ^ 2 / ((iNLoS - 1) * ((dAlphaCoverRad * (dObstRadAve ^ 3) * Sin(dBetaRad)) / _
3) ^ 2)
        'dVSIaveSTD = dVSIaveSTD + ((dVisVolAve(p3) - dInVisVolAve(p3)) - (dTotVolAve / _
        iNLoS)) ^ 2 / (iNLoS - 1)
1844:     Next p3
1845:     dVSIaveSTD = dVSIaveSTD ^ (1 / 2)
    
1847:     dVSIminSTD = 0
1848:     For p4 = 1 To iNLoS
1849:         dVSIminSTD = dVSIminSTD + (dVisVolMin(p4) - (dTotVolMin / iNLoS)) ^ 2 / ((iNLoS - 1) _
        * ((dAlphaCoverRad * (dObstRadMin ^ 3) * Sin(dBetaRad)) / 3) ^ 2)
        'dVSIminSTD = dVSIminSTD + (dVisVolMin(p4) - (dTotVolMin / iNLoS)) ^ 2 / (iNLoS - 1)
1852:     Next p4
1853:     dVSIminSTD = dVSIminSTD ^ (1 / 2)
    
1855:     dSVFSTD = 0
1856:     For p5 = 1 To iNLoS
1857:         dSVFSTD = dSVFSTD + ((dSegSVF(p5) - dSVF) ^ 2) / (iNLoS - 1)
1858:     Next p5
1859:     dSVFSTD = dSVFSTD ^ (1 / 2)
    
1861: Else '(1) iNoObstrCount < iNLoS '10/03/04
    
1863:     dTotVol = 0
1864:     dTotInVol = 0
1865:     dTotInVolMax = 0
1866:     dTotInVolAve = 0
1867:     dTotSurf = 0
1868:     dVizEdgePerc = 0
1869:     dViewsphere = 0
1870:     dViewsphereSTD = 0 'rocky
1871:     dVSImax = 0
1872:     dVSImaxSTD = 0 'rocky
1873:     dVSIave = 0
1874:     dVSIaveSTD = 0 'rocky
1875:     dVSImin = 0
1876:     dVSIminSTD = 0 'rocky
1877:     dInVSI = 0
1878:     dInVSImax = 0
1879:     dInVSIave = 0
1880:     dOcclRatioMax = 0
1881:     dOcclRatioAve = 0
1882:     dSVF = 1
1883:     dSVFSTD = 0 'rocky
1884:     dVisBetaDegMax = 0
1885:     dVisBetaDegAve = 0
1886:     dDistanceHeightMax = 0
1887:     dDistanceHeightAve = 0
1888:     dTotVisArea = PI * (dRadius ^ 2)
1889:     dTotPerimeter = 2 * PI * dRadius
1890:     If (dAlphaCoverDeg < 180) Then '(2)
1891:         dLongestAxis = dRadius
1892:     Else '(2)
1893:         dLongestAxis = 2 * dRadius
1894:     End If '(2)
1895:     dEnclosureLoose = dAlphaCoverDeg
    
1897: End If '(1) iNoObstrCount < iNLoS '10/03/04

'2D shape indicators
1900: If (dTotPerimeter > 0) Then '(1) Check against overflow 10/03/04
1901:     dCompact = 2 * ((Sqr(PI * dTotVisArea)) / dTotPerimeter) '08/03/04
1902: End If '(1)
1903: If (dTotVisArea > 0) Then '(1)
1904:     dFractal = 2 * ((Log(dTotPerimeter)) / (Log(dTotVisArea))) '08/03/04
1905: End If '(1)
1906: If (dLongestAxis > 0) Then '(1)
1907:     dConvolution = dTotPerimeter / dLongestAxis '08/03/04
1908: End If '(1)
1909: If (dTotVisArea > 0) Then 'Patton Diversity (1)
1910:     dPattonDiv = dTotPerimeter / (2 * Sqr(PI * dTotVisArea)) '27/05/04
1911: End If '(1)
1912: If (dTotVisArea > 0) Then 'Ellipticity (1)
1913:     dElliptic = (0.5 * PI * (dLongestAxis ^ 2)) / dTotVisArea '27/05/04
1914: End If '(1)

'VSIs Ratio Calculation '29/04/04
1917: If Not ((dVSImax = 0) Or (dVSIave = 0) Or (dVSImin = 0)) Then '(1)
1918:     dRatioMinMax = dVSImin / dVSImax
1919:     dRatioAveMax = dVSIave / dVSImax
1920:     dRatioMinAve = dVSImin / dVSIave
1921: Else '(1)
1922:     dRatioMinMax = 0
1923:     dRatioAveMax = 0
1924:     dRatioMinAve = 0
1925: End If '(1)

'Plaza P '29/04/04
1928: If ((dRatioMinMax >= 36.2) And (dRatioMinMax <= 50.8)) And _
    ((dRatioAveMax >= 14.9) And (dRatioAveMax <= 26)) Then 'Or
    '((dRatioMinAve > 1.56) And (dRatioMinAve < 2)) Then
1931:     sTypology = "Plaza"
    'Intersection IT '29/04/04
1933:     ElseIf ((dRatioMinMax >= 59.7) And (dRatioMinMax <= 390.4)) Then 'Or
    '((dRatioAveMax > 33.2) And (dRatioAveMax < 73.1)) Or
    '((dRatioMinAve > 1.84) And (dRatioMinAve < 3.19)) Then
1936:     sTypology = "Intersection"
    'Street ST '29/04/04
1938:     ElseIf ((dRatioMinMax >= 139.5) And (dRatioMinMax <= 1020)) Then 'Or
    '((dRatioAveMax > 65.5) And (dRatioAveMax < 87.5)) 'Or
    '((dRatioMinAve > 3.52) And (dRatioMinAve < 6.11)) Then
1941:     sTypology = "Street"
    'Street/Intersection-next-to-plaza (P-ST/IT) '29/04/04
1943: Else
1944:     sTypology = "Plaza nearby"
1945: End If


'result output
1949: sTarVizVol = Format(dTotVol, "#.##")
1950: sTotInVol = Format(dTotInVol, "#.##")
1951: sTotSurf = Format(dTotSurf, "#.##")
1952: sUserRadius = Format(dRadius, "#.##")
1953: sVizEdge = Format(dVizEdgePerc, "#.##")
1954: sViewsphere = Format(dViewsphere, "#.####")
1955: sViewsphereSTD = Format(dViewsphereSTD, "#.##########") 'rocky
1956: sInVSI = Format(dInVSI, "#.####")
1957: sObstRadMax = Format(dObstRadMax, "#.###")
1958: sObstRadAve = Format(dObstRadAve, "#.###")
1959: sObstRadMin = Format(dObstRadMin, "#.###")
1960: sVSImax = Format(dVSImax, "#.####")
1961: sVSImaxSTD = Format(dVSImaxSTD, "#.##########") 'rocky
1962: sVSIave = Format(dVSIave, "#.####")
1963: sVSIaveSTD = Format(dVSIaveSTD, "#.##########") 'rocky
1964: sVSImin = Format(dVSImin, "#.####")
1965: sVSIminSTD = Format(dVSIminSTD, "#.##########") 'rocky
1966: sInVSImax = Format(dInVSImax, "#.####")
1967: sInVSIave = Format(dInVSIave, "#.####")
1968: sOcclRatMax = Format(dOcclRatioMax, "#.###")
1969: sOcclRatAve = Format(dOcclRatioAve, "#.###")
1970: sSVF = Format(dSVF, "#.####")
1971: sSVFSTD = Format(dSVFSTD, "#.##########") 'rocky
1972: sBetaDegMax = Format(dVisBetaDegMax, "##.####")
1973: sBetaDegAve = Format(dVisBetaDegAve, "##.####")
1974: sDistanceHeightMax = Format(dDistanceHeightMax, "##.####")
1975: sDistanceHeightAve = Format(dDistanceHeightAve, "##.####")
1976: sVisArea = Format(dTotVisArea, "#.##")
1977: sPerimeter = Format(dTotPerimeter, "#.##")
1978: sLongestAxis = Format(dLongestAxis, "#.##")
1979: sCompact = Format(dCompact, "#.###")
1980: sFractal = Format(dFractal, "#.###")
1981: sConvolution = Format(dConvolution, "#.###")
1982: sPattonDiv = Format(dPattonDiv, "#.###")
1983: sElliptic = Format(dElliptic, "#.###")
1984: sEnclosFull = Format(dEnclosureFull, "#.##")
1985: sEnclosThres = Format(dEnclosureThreshold, "#.##")
1986: sEnclosMin = Format(dEnclosureMin, "#.##")
1987: sEnclosLoose = Format(dEnclosureLoose, "#.##")

'Status bar
1990: m_pApp.StatusBar.Message(0) = "Total vol. = " & sTarVizVol & _
" Viewsphere (VSI) = " & sViewsphere & _
" MaxRad = " & sObstRadMax & _
" VSImax = " & sVSImax & _
" AveRad = " & sObstRadAve & _
" VSIave = " & sVSIave & _
" MinRad = " & sObstRadMin & _
" VSImin = " & sVSImin & _
" Sky View Factor = " & sSVF & _
" Sky View Factor STD = " & sSVFSTD & _
" MaxBeta = " & sBetaDegMax

Dim TxtLogString As String
'2003: TxtLogString = "Visible vol. = " & sTarVizVol & vbNewLine & "Invisible vol. = " & sTotInVol & vbNewLine & _
        "Visible Surf. = " & sTotSurf & vbNewLine & "Visual Edge = " & sVizEdge & "%" & vbNewLine & _
        "UserRad = " & sUserRadius & vbNewLine & "VSI = " & sViewsphere & vbNewLine & _
        "VSI STD = " & sViewsphereSTD & vbNewLine & "MaxRad = " & sObstRadMax & vbNewLine & _
        "VSImax = " & sVSImax & vbNewLine & "VSImax STD = " & sVSImaxSTD & vbNewLine & _
        "AveRad = " & sObstRadAve & vbNewLine & "VSIave = " & sVSIave & vbNewLine & _
        "VSIave STD = " & sVSIaveSTD & vbNewLine & "MinRad = " & sObstRadMin & vbNewLine & _
        "VSImin = " & sVSImin & vbNewLine & "VSImin STD = " & sVSIminSTD & vbNewLine & _
        "inVSI = " & sInVSI & vbNewLine & "inVSImax = " & sInVSImax & vbNewLine & _
        "inVSIave = " & sInVSIave & vbNewLine & "Max Occlusion Ratio = " & sOcclRatMax & vbNewLine & _
        "Ave Occlusion Ratio = " & sOcclRatAve & vbNewLine & "SVF = " & sSVF & vbNewLine & _
        "SVF STD = " & sSVFSTD & vbNewLine & "Area = " & sVisArea & vbNewLine & "Perimeter = " & sPerimeter & vbNewLine & _
        "Longest Axis = " & sLongestAxis & vbNewLine & "Compactness = " & sCompact & vbNewLine & _
        "Convolution = " & sConvolution & vbNewLine & "Fractal = " & sFractal & vbNewLine & _
        "Patton Diversity = " & sPattonDiv & vbNewLine & "Ellipticity = " & sElliptic & vbNewLine & _
        "MaxBeta = " & sBetaDegMax & vbNewLine & "Max D/H = " & sDistanceHeightMax & vbNewLine & _
        "AveBeta = " & sBetaDegAve & vbNewLine & "Ave D/H = " & sDistanceHeightAve & vbNewLine & _
        "FullEnclos = " & sEnclosFull & vbNewLine & "ThresEnclos = " & sEnclosThres & vbNewLine & _
        "MinEnclos = " & sEnclosMin & vbNewLine & "LooseEnclos = " & sEnclosLoose & vbNewLine & _
        "VSImin/VSImax = " & dRatioMinMax & vbNewLine & "VSIave/VSImax = " & dRatioAveMax & vbNewLine & _
        "VSImin/VSIave = " & dRatioMinAve & vbNewLine & "Typology = " & sTypology
2003: TxtLogString = "SVF = " & sSVF & vbNewLine & "SVF STD = " & sSVFSTD

'logging in textbox
2026: If FrameLog.Visible = True Then
    'txtbox cannot be too long, it's safer to reset it on every 30th point.
2028: If (iPntCount - 1) / 30 = Int((iPntCount - 1) / 30) Then
2029:     TxtLog.Text = ""
2030: End If
2031:     TxtLog.SelStart = Len(TxtLog.Text)
2032:     TxtLog.SelText = "Point No. " & iPntCount & " Started." & vbNewLine & TxtLogString & _
vbNewLine & "Point No. " & iPntCount & " completed." & vbNewLine & vbNewLine

2035: End If

'Output as CSV file
Dim TxtCsvTitle As String
Dim iCsv As Long
Dim sCsv As String

2042: If Not (txtCsvName.Text = Empty) Then
2043:     TxtCsvTitle = "Point No.," & "Visible vol.," & "Invisible vol.," & "Visible Surf.," & "Visual Edge," & _
                "UserRad," & "VSI," & "VSI STD," & "MaxRad," & "VSImax," & _
                "VSImax STD," & "AveRad," & "VSIave," & "VSIave STD," & _
                "MinRad," & "VSImin," & "VSImin STD," & "inVSI," & _
                "inVSImax," & "inVSIave," & "Max Occlusion Ratio," & "Ave Occlusion Ratio," & _
                "Sky View Factor (SVF)," & "Sky View Factor (SVF) STD," & "Area," & "Perimeter," & _
                "Longest Axis," & "Compactness," & "Convolution," & "Fractal," & _
                "Patton Diversity," & "Ellipticity," & "MaxBeta," & "Max D/H," & _
                "AveBeta," & "Ave D/H," & "FullEnclos," & "ThresEnclos," & _
                "MinEnclos," & "LooseEnclos," & "VSImin/VSImax," & "VSIave/VSImax," & _
                "VSImin/VSIave," & "Typology," & vbNewLine
                
'2055:         iCsv = Int((iPntCount - 1) / 200) + 1
'2056:         If (iPntCount - 1) / 200 = iCsv - 1 Then
'2057:             TxtCsv = ""
'2058:         End If
2059:     TxtCsv = iPntCount & "," & sTarVizVol & "," & sTotInVol & "," & sTotSurf & "," & sVizEdge & "%" & "," & _
                sUserRadius & "," & sViewsphere & "," & sViewsphereSTD & "," & sObstRadMax & "," & _
                sVSImax & "," & sVSImaxSTD & "," & sObstRadAve & "," & sVSIave & "," & _
                sVSIaveSTD & "," & sObstRadMin & "," & sVSImin & "," & sVSIminSTD & "," & _
                sInVSI & "," & sInVSImax & "," & sInVSIave & "," & sOcclRatMax & "," & _
                sOcclRatAve & "," & sSVF & "," & sSVFSTD & "," & sVisArea & "," & sPerimeter & "," & _
                sLongestAxis & "," & sCompact & "," & sConvolution & "," & sFractal & "," & _
                sPattonDiv & "," & sElliptic & "," & sBetaDegMax & "," & sDistanceHeightMax & "," & _
                sBetaDegAve & "," & sDistanceHeightAve & "," & sEnclosFull & "," & sEnclosThres & "," & _
                sEnclosMin & "," & sEnclosLoose & "," & dRatioMinMax & "," & dRatioAveMax & "," & _
                dRatioMinAve & "," & sTypology & vbNewLine
     
2071:     iCsv = Int((iPntCount - 1) / 400) + 1
2072:     sCsv = Format(iCsv, "####")
2073:           If iCsv < 10 Then
2074:             sCsv = "0" + sCsv
2075:           End If
          
2077:           If iCsv < 100 Then
2078:             sCsv = "0" + sCsv
2079:           End If
          
2081:           If iCsv < 1000 Then
2082:             sCsv = "0" + sCsv
2083:           End If
          
'2085:         If iPntCount / 200 = Int(iPntCount / 200) Then
'2086:             Print2Csv TxtCsvTitle, TxtCsv, txtCsvName.Text + sCsv + ".csv"
'2087:         End If
2088:         If (iPntCount - 1) / 400 = Int((iPntCount - 1) / 400) Then
2089:             CreateCsv txtCsvName.Text + sCsv + ".csv", TxtCsvTitle
'2090:             Print2Csv TxtCsvTitle, txtCsvName.Text + sCsv + ".csv"
2091:         End If

2093:     Print2Csv TxtCsv, txtCsvName.Text + sCsv + ".csv"
2094: End If

'Storing Value in shapefile
'  If (TypeOf m_pApp Is IMxApplication) Then 'ArcMAP 27/02/04
2097: If frmLineOfSight.chkStore = 1 Then '(1) Optional Store value 27/02/04
2098:     pFeature.Value(indexRadius) = dRadius
2099:     pFeature.Value(indexTotVol) = dTotVol
2100:     pFeature.Value(indexVSI) = dViewsphere
2101:     pFeature.Value(indexVSISTD) = dViewsphereSTD
2102:     pFeature.Value(indexMaxRad) = dObstRadMax
2103:     pFeature.Value(indexVSImax) = dVSImax
2104:     pFeature.Value(indexVSImaxSTD) = dVSImaxSTD
2105:     pFeature.Value(indexAveRad) = dObstRadAve
2106:     pFeature.Value(indexVSIave) = dVSIave
2107:     pFeature.Value(indexVSIaveSTD) = dVSIaveSTD
2108:     pFeature.Value(indexMinRad) = dObstRadMin
2109:     pFeature.Value(indexVSImin) = dVSImin
2110:     pFeature.Value(indexVSIminSTD) = dVSIminSTD
2111:     pFeature.Value(indexSVF) = dSVF
2112:     pFeature.Value(indexSVFSTD) = dSVFSTD
2113:     pFeature.Value(indexBetaMax) = dVisBetaDegMax
2114:     pFeature.Value(indexBetaAve) = dVisBetaDegAve
2115:     pFeature.Value(indexVisArea) = dTotVisArea
2116:     pFeature.Value(indexVisPeri) = dTotPerimeter
2117:     pFeature.Value(indexLongestAxis) = dLongestAxis
2118:     pFeature.Value(indexCompact) = dCompact
2119:     pFeature.Value(indexFractal) = dFractal
2120:     pFeature.Value(indexConvo) = dConvolution
2121:     pFeature.Value(indexElliptic) = dElliptic
2122:     pFeature.Value(indexRMinMax) = dRatioMinMax
2123:     pFeature.Value(indexRAveMax) = dRatioAveMax
2124:     pFeature.Value(indexRMinAve) = dRatioMinAve
2125:     pFeature.Store 'STORE RESULT
2126: End If '(1) Optional Store value 27/02/04
'  End If 'ArcMAP 27/02/04

2129: dTotVol = 0
2130: dTotInVol = 0
2131: dTotVolMax = 0
2132: dTotInVolMax = 0
2133: dTotVolAve = 0
2134: dTotInVolAve = 0
2135: dTotVolMin = 0
2136: dTotSVF = 0
2137: dViewsphereSTD = 0 'rocky
2138: dVSImaxSTD = 0 'rocky
2139: dVSIaveSTD = 0 'rocky
2140: dVSIminSTD = 0 'rocky
2141: dSVFSTD = 0 'rocky
2142: dObstRadMax = 0
2143: dObstRadTot = 0
2144: dObstRadAve = 0
2145: dVisBetaRadMax = 0
2146: dTotBetaRad = 0
2147: dTotVisArea = 0
2148: dTotPerimeter = 0
2149: dLongestAxis = 0
2150: dEnclosureFull = 0
2151: dEnclosureThreshold = 0
2152: dEnclosureMin = 0
2153: dEnclosureLoose = 0
2154: iNoObstrCount = 0
2155: iNVizEdge = 0

2157: If frmLineOfSight.chkShowMsg = 1 Then '(1) Optional Show Message Box
    '    MsgBox ("Visible vol. = " & sTarVizVol & " ; Invisible vol. = " & sTotInVol & (Chr(13)) & _
        "  Visible Surf. = " & sTotSurf & " ; Visual Edge = " & sVizEdge & "%" & (Chr(13)) & _
        "  UserRad = " & sUserRadius & " ; VSI = " & sViewsphere & (Chr(13)) & _
        "  VSI STD = " & sViewsphereSTD & (Chr(13)) & _
        "  MaxRad = " & sObstRadMax & " ; VSImax = " & sVSImax & (Chr(13)) & _
        "  VSImax STD = " & sVSImaxSTD & (Chr(13)) & _
        "  AveRad = " & sObstRadAve & " ; VSIave = " & sVSIave & (Chr(13)) & _
        "  VSIave STD = " & sVSIaveSTD & (Chr(13)) & _
        "  MinRad = " & sObstRadMin & " ; VSImin = " & sVSImin & (Chr(13)) & _
        "  VSImin STD = " & sVSIminSTD & (Chr(13)) & _
        "  inVSI = " & sInVSI & " ; inVSImax = " & sInVSImax & " ; inVSIave = " & sInVSIave & (Chr(13)) & _
        "  Max Occlusion Ratio = " & sOcclRatMax & " ; Ave Occlusion Ratio = " & sOcclRatAve & (Chr(13)) & _
        "  Sky View Factor (SVF) = " & sSVF & (Chr(13)) & _
        "  Sky View Factor (SVF) STD = " & sSVFSTD & (Chr(13)) & (Chr(13)) & _
        "  Area = " & sVisArea & " ; Perimeter = " & sPerimeter & " ; Longest Axis = " & sLongestAxis & (Chr(13)) & _
        "  Compactness = " & sCompact & " ; Convolution = " & sConvolution & " ; Fractal = " & sFractal & (Chr(13)) & _
        "  Patton Diversity = " & sPattonDiv & " ; Ellipticity = " & sElliptic & (Chr(13)) & (Chr(13)) & _
        "  MaxBeta = " & sBetaDegMax & " ; Max D/H = " & sDistanceHeightMax & (Chr(13)) & _
        "  AveBeta = " & sBetaDegAve & " ; Ave D/H = " & sDistanceHeightAve & (Chr(13)) & _
        "  FullEnclos = " & sEnclosFull & " ; ThresEnclos = " & sEnclosThres & (Chr(13)) & _
        "  MinEnclos = " & sEnclosMin & " ; LooseEnclos = " & sEnclosLoose & (Chr(13)) & (Chr(13)) & _
        "  VSImin/VSImax = " & dRatioMinMax & (Chr(13)) & _
        "  VSIave/VSImax = " & dRatioAveMax & (Chr(13)) & _
        "  VSImin/VSIave = " & dRatioMinAve & (Chr(13)) & _
        "  Typology = " & sTypology)
    
'Output as transparent text 'v0.20
2185:     ShowTranTxt TxtLogString, 0
    
2187:     MsgBoxPause
2188:     ShowTranTxt "", 1
    
2190: End If '(1)


'Next hpoints

'Input Point from selected Point Shapefile
2196: Set pFeature = pFCursor.NextFeature
2197: Set pGeomInput = Nothing

2199: Wend

'End if (bPntFrmLayer)

'End calculation (27-02-2004)

2205: ShowTranTxt "", 1

2207: CmdOK.Visible = True
2208: cmdStop.Visible = False



  Exit Sub
ErrorHandler:
  HandleError True, "CmdOK_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub cmdCancel_Click()
  On Error GoTo ErrorHandler

    
2221:     Unload Me
    


  Exit Sub
ErrorHandler:
  HandleError False, "cmdCancel_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub CmdReset_Click()
  On Error GoTo ErrorHandler

    
2234:     If Not (TxtLog.Text = Empty) Then
        'To decide the position of MsgEx
2236:         MsgEx.LeftDialogPos = MsgBoxPosLeft
2237:         MsgEx.TopDialogPos = MsgBoxPosTop
        
        'To define MsgEx
2240:         MsgEx.Icon = mbQuestion
2241:         MsgEx.CustomIconSound = mbQuestionSound
2242:         MsgEx.Title = "Confirmation"
2243:         MsgEx.Position = vbStartUpManual
2244:         MsgEx.Buttons = mbYesNo
2245:         MsgEx.Prompt = "Do you want to clear the logging TextBox?"
2246:         If MsgEx.Show = mbYes Then
2247:             TxtLog.Text = ""
2248:         Else
            Exit Sub
2250:         End If
2251:     End If
    
    'If Not (TxtLog.Text = Empty) Then
    '    Dim Msg, Style, Title, Response
    '    Msg = "Do you want to clear the logging TextBox?"   ' Define message.
    '    Style = vbYesNo + vbQuestion + vbDefaultButton2 + vbMsgBoxSetForeground   ' _
Define buttons.
    '    Title = "Confirmation"   ' Define title.
    '    Display message.
    '    Response = MsgBox(Msg, Style, Title)
    '    If Response = vbYes Then   ' User chose Yes.
    '       TxtLog.Text = ""
    '       Else   ' User chose No.
    '       Exit Sub
    '    End If
    'End If
    


  Exit Sub
ErrorHandler:
  HandleError True, "CmdReset_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub CmdSav2CSV_File_Click()
  On Error GoTo ErrorHandler

    
    ' Set CancelError is True
2280:     CommonDialog1.CancelError = False
    ' Set flags
2282:     CommonDialog1.FLAGS = cdlOFNHideReadOnly
    ' Set filters
2284:     CommonDialog1.Filter = "All Files (*.*)|*.*|CSV Files" & _
    "(*.csv)|*.csv"
    ' Specify default filter
2287:     CommonDialog1.FilterIndex = 2
    ' Display the Open dialog box
2289:     CommonDialog1.ShowOpen
    ' Display name of selected file
2291:     If Not (CommonDialog1.FileName = Empty) Then
2292:         txtCsvName.Text = CommonDialog1.FileName
2293:     End If
    


  Exit Sub
ErrorHandler:
  HandleError False, "CmdSav2CSV_File_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub cmdStop_Click()
  On Error GoTo ErrorHandler

    
2306:     cmdStop.Visible = False
2307:     CmdOK.Visible = True
2308:     If m_Stop = False Then
2309:         m_Stop = True
2310:     End If
        


  Exit Sub
ErrorHandler:
  HandleError True, "cmdStop_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub Command1_Click()
  On Error GoTo ErrorHandler




  Exit Sub
ErrorHandler:
  HandleError False, "Command1_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub Form_Load()
  On Error GoTo ErrorHandler

    
2334:     Me.Icon = Nothing
2335:     txtObsOffset.Text = Str(Round(m_obsOffset, 2))
2336:     txtTarOffset.Text = Str(Round(m_tarOffset, 2))
2337:     txtiNLoS.Text = Str(Round(m_iNLoS, 2))
2338:     txtRadius.Text = Str(Round(m_Radius, 2))
2339:     txtBetaDeg.Text = Str(Round(m_BetaDeg, 2))
2340:     txtAlphaInitDeg.Text = Str(Round(m_AlphaInitDeg, 2))
2341:     txtAlphaCoverDeg.Text = Str(Round(m_AlphaCoverDeg, 2))
2342:     chkCurv.Value = 0
2343:     chkViz.Value = 0
2344:     chkShowMsg.Value = 0
2345:     chkShowSeg.Value = 0
2346:     chkStore.Value = 0
2347:     win32Util.FloatWindow Me, True
2348:     m_Stop = False 'v0.13
2349:     m_StopfrPause = False 'v0.13
2350:     frmHeightOri = frmLineOfSight.Height 'v0.14
2351:     TransMsg1.TXT_COLOR = vbRed 'v0.20
    


  Exit Sub
ErrorHandler:
  HandleError True, "Form_Load " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo EH
2362:     m_obsOffset = CDbl(txtObsOffset.Text)
2363:     m_tarOffset = CDbl(txtTarOffset.Text)
2364:     m_iNLoS = CDbl(txtiNLoS.Text)
2365:     m_Radius = CDbl(txtRadius.Text)
2366:     m_BetaDeg = CDbl(txtBetaDeg.Text)
2367:     m_AlphaInitDeg = CDbl(txtAlphaInitDeg.Text)
2368:     m_AlphaCoverDeg = CDbl(txtAlphaCoverDeg.Text)
    Exit Sub
EH:
2371:     MsgBox Err.Description
End Sub

Private Function MsgBoxPause()
  On Error GoTo ErrorHandler

    
    'To decide the position of MsgEx
2379:     MsgEx.LeftDialogPos = MsgBoxPosLeft
2380:     MsgEx.TopDialogPos = (frmLineOfSight.Top + 9000) / 15
    
    'To define MsgEx
2383:     MsgEx.Icon = mbQuestion
2384:     MsgEx.CustomIconSound = mbQuestionSound
2385:     MsgEx.Title = "Pause"
2386:     MsgEx.Position = vbStartUpManual
2387:     MsgEx.Buttons = mbYesNo
2388:     MsgEx.Prompt = "Do you want to continue ?"
2389:     If MsgEx.Show = mbYes Then
        Exit Function
2391:     Else
2392:         TxtLog.SelStart = Len(TxtLog.Text)
2393:         TxtLog.SelText = vbNewLine & vbNewLine & "Process cancelled by user"
2394:         m_StopfrPause = True
2395:     End If
    
    'Dim Msg, Style, Title, Response
    '2194: Msg = "Do you want to continue ?"   ' Define message.
    '2195: Style = vbYesNo + vbQuestion + vbDefaultButton1 + vbMsgBoxSetForeground   ' _
Define buttons.
    '2196: Title = "Pause"   ' Define title.
    ' Display message.
    '2198: Response = MsgBox(Msg, Style, Title)
    '2199: If Response = vbYes Then   ' User chose Yes.
    '   Exit Function
    '2201: Else   ' User chose No.
    '2202:    TxtLog.SelStart = Len(TxtLog.Text)
    '2203:    TxtLog.SelText = vbNewLine & vbNewLine & "Process cancelled by user"
    '2204:    m_StopfrPause = True
    '2205: End If
    


  Exit Function
ErrorHandler:
  HandleError False, "MsgBoxPause " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Private Sub Print2Csv(ByVal pTxt As String, ByVal pPathTitle As String)
  On Error GoTo ErrorHandler
    Const ForWriting = 2, ForAppending = 8
    Const TristateUseDefault = 2, TristateTrue = 1, TristateFalse = 0
    Dim fs, a
2424:     Set fs = CreateObject("Scripting.FileSystemObject")
'2425:     Set a = fs.OpenTextFile(pPathTitle, True)
2425:   Set a = fs.OpenTextFile(pPathTitle, ForAppending, TristateFalse)
'2426:    Set ts = a.OpenAsTextStream(ForAppending, TristateUseDefault)
2427:    a.Write pTxt
2428:    a.Close

  Exit Sub
ErrorHandler:
  HandleError False, "Print2Csv " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub CreateCsv(ByVal pPathTitle As String, pFirstLine As String)
  On Error GoTo ErrorHandler

Dim fs, a
2450:     Set fs = CreateObject("Scripting.FileSystemObject")
2451:     Set a = fs.CreateTextFile(pPathTitle, True)
2452:       a.Write pFirstLine
2453:     a.Close

  Exit Sub
ErrorHandler:
  HandleError False, "Print2Csv " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Function MsgBoxPosLeft() As Long
  On Error GoTo ErrorHandler

2463:     Set pWindPos = m_pApp
2464:     If ((frmLineOfSight.Left + frmLineOfSight.Width / 2) / 15) < (pWindPos.Left + _
pWindPos.Width / 2) Then
2466:         MsgBoxPosLeft = (frmLineOfSight.Left + frmLineOfSight.Width) / 15 + 30
2467:         Else: MsgBoxPosLeft = frmLineOfSight.Left / 15 - 300
2468:     End If


  Exit Function
ErrorHandler:
  HandleError False, "MsgBoxPosLeft " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Private Function MsgBoxPosTop() As Long
  On Error GoTo ErrorHandler

2479:     MsgBoxPosTop = pWindPos.Height / 2 - 100


  Exit Function
ErrorHandler:
  HandleError False, "MsgBoxPosTop " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Private Sub ShowMsgInfo(Title As String, txtMsg As String)
  On Error GoTo ErrorHandler

    'To decide the position of MsgEx
2491:     MsgEx.LeftDialogPos = MsgBoxPosLeft
2492:     MsgEx.TopDialogPos = MsgBoxPosTop
    
    'To define MsgEx
2495:     MsgEx.Icon = mbInformation
2496:     MsgEx.CustomIconSound = mbInformationSound
2497:     MsgEx.Title = Title
2498:     MsgEx.Position = vbStartUpManual
2499:     MsgEx.Buttons = mbOKOnly
2500:     MsgEx.Prompt = txtMsg
2501:     MsgEx.Show


  Exit Sub
ErrorHandler:
  HandleError False, "ShowMsgInfo " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub ShowMsgCritical(Title As String, txtMsg As String)
  On Error GoTo ErrorHandler

    'To define and show MsgEx
2513:     MsgEx.LeftDialogPos = MsgBoxPosLeft
2514:     MsgEx.TopDialogPos = MsgBoxPosTop
    
    'To define MsgEx
2517:     MsgEx.Icon = mbCritical
2518:     MsgEx.CustomIconSound = mbCriticalSound
2519:     MsgEx.Title = Title
2520:     MsgEx.Position = vbStartUpManual
2521:     MsgEx.Buttons = mbOKOnly
2522:     MsgEx.Prompt = txtMsg
2523:     MsgEx.Show


  Exit Sub
ErrorHandler:
  HandleError False, "ShowMsgCritical " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub ShowTranTxt(Text, Index As Integer)
  On Error GoTo ErrorHandler

Select Case Index
Dim lng&
    Case Is = 0 'draw the text
        '   If optAlign(0).Value = True Then
        '       lng = DT_left
        '   ElseIf optAlign(1).Value = True Then
        '       lng = DT_CENTER
        '   ElseIf optAlign(2).Value = True Then
        '       lng = DT_right
        '   End If
        'this is the bounding rectangle
        'if the text gets cliped, just increase
        'its width or height
2547:         TransMsg1.DrawTextToScreen Text, _
                                       (MsgBoxPosLeft - 20), _
                                       (MsgBoxPosLeft + 180), _
                                       (frmLineOfSight.Top) / 15, _
                                       (frmLineOfSight.Top + 9000) / 15, _
                                       DT_left, _
                                       TransMsg1.TXT_COLOR, _
                                       CBool(0), _
                                       4

    Case Is = 1 'erase the text
2558:             TransMsg1.EraseTextDrawnToScreen
2559: End Select


  Exit Sub
ErrorHandler:
  HandleError False, "ShowTranTxt " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub optMapFC_Click()
    cmdBrInputSF.Enabled = False
End Sub

Private Sub optSFFC_Click()
    cmdBrInputSF.Enabled = True
End Sub
