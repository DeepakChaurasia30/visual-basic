VERSION 5.00
Begin VB.Form FrmReports 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reports"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4920
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmReports.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   4920
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cashrptbtn 
      Caption         =   "&Cash Report"
      Height          =   480
      Left            =   180
      TabIndex        =   4
      Top             =   1440
      Width           =   4650
   End
   Begin VB.CommandButton CmdClose 
      Caption         =   "&CLOSE"
      Height          =   480
      Left            =   180
      TabIndex        =   3
      Top             =   2670
      Width           =   4650
   End
   Begin VB.CommandButton CmdBillSummary 
      Caption         =   "Bill Summary"
      Height          =   480
      Left            =   180
      TabIndex        =   2
      Top             =   2070
      Width           =   4650
   End
   Begin VB.CommandButton CmdProductSummary 
      Caption         =   "&No Bill Report"
      Height          =   480
      Left            =   180
      TabIndex        =   1
      Top             =   810
      Width           =   4650
   End
   Begin VB.CommandButton CmdPrintBill 
      Caption         =   "Find/Print &Bill"
      Height          =   480
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   4650
   End
End
Attribute VB_Name = "FrmReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cashrptbtn_Click()
Cashrpt.Show 1
End Sub

Private Sub CmdPrintBill_Click()
FrmPrintBill.Show 1
End Sub

Private Sub CmdProductSummary_Click()
nbill.Show 1
End Sub

Private Sub Form_Load()
Me.Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 6
End Sub
