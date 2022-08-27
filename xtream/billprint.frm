VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Begin VB.Form billprint 
   Caption         =   "Bill Print"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   Icon            =   "billprint.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin CRVIEWER9LibCtl.CRViewer9 CRViewer 
      Height          =   2985
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4515
      lastProp        =   500
      _cx             =   7964
      _cy             =   5265
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
   End
End
Attribute VB_Name = "billprint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RS1 As New ADODB.Recordset
Dim Rs As New ADODB.Recordset
Dim inv As Double


Private Sub Form_Load()
Dim CRApp As New CRAXDRT.Application
Dim CRRpt As New CRAXDRT.Report
Dim CRTbl As CRAXDRT.DatabaseTable
'>>> Check for bill type
inv = FrmPrintBill.invno
 If Rs.State = adStateOpen Then Rs.Close
    Rs.Open "select client_name from bill where invoice_no=" & inv & "", Cn, adOpenStatic, adLockReadOnly
    If RS1.State = adStateOpen Then RS1.Close
    RS1.Open "select i_gst from client_master where client_name ='" & Rs("client_name") & "'", Cn, adOpenStatic, adLockReadOnly
    If RS1("i_gst") <= 0 Then
Set CRRpt = CRApp.OpenReport(App.path & "/reports/bill.rpt")
Else
Set CRRpt = CRApp.OpenReport(App.path & "/reports/igst_bill.rpt")
End If
If CRRpt.HasSavedData Then CRRpt.DiscardSavedData
CRRpt.VerifyOnEveryPrint = True
CRViewer.EnableExportButton = True
CRViewer.EnableCloseButton = True
CRViewer.EnableProgressControl = True
CRViewer.EnableZoomControl = True
CRViewer.ReportSource = CRRpt
CRRpt.Database.Tables(1).ConnectionProperties("Jet OLEDB:Database Password") = "xtream@30"
CRViewer.ViewReport
CRViewer.Refresh
CRViewer.DisplayGroupTree = False
Me.WindowState = vbMaximized


    Set CRTbl = Nothing
    Set CRRpt = Nothing
    Set CRApp = Nothing
End Sub

Private Sub Form_Resize()
With CRViewer
.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
.Zoom (1)
End With
End Sub
