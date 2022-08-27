VERSION 5.00
Begin VB.Form FrmCompanyInfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Compnay Info"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7230
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmCompanyInfo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "&Update"
      Height          =   435
      Left            =   5340
      TabIndex        =   16
      Top             =   4680
      Width           =   1785
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Edit"
      Height          =   465
      Left            =   240
      TabIndex        =   15
      Top             =   4620
      Width           =   1665
   End
   Begin VB.TextBox txtgst 
      Height          =   405
      Left            =   2310
      TabIndex        =   14
      Top             =   3930
      Width           =   3105
   End
   Begin VB.TextBox TxtVatNo 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   2280
      MaxLength       =   200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   2820
      Width           =   4440
   End
   Begin VB.TextBox TxtPin 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2280
      MaxLength       =   50
      TabIndex        =   4
      Top             =   1970
      Width           =   3405
   End
   Begin VB.TextBox TxtCity 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2280
      MaxLength       =   50
      TabIndex        =   3
      Top             =   1545
      Width           =   3405
   End
   Begin VB.TextBox TxtAddress2 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2280
      MaxLength       =   50
      TabIndex        =   2
      Top             =   1120
      Width           =   4575
   End
   Begin VB.TextBox TxtAddress1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2280
      MaxLength       =   50
      TabIndex        =   1
      Top             =   695
      Width           =   4575
   End
   Begin VB.TextBox TxtTelephone 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2280
      MaxLength       =   50
      TabIndex        =   5
      Top             =   2395
      Width           =   3405
   End
   Begin VB.TextBox TxtCompanyName 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2280
      MaxLength       =   50
      TabIndex        =   0
      Top             =   270
      Width           =   4575
   End
   Begin VB.Label Label7 
      Caption         =   "GST NO. :"
      Height          =   555
      Left            =   360
      TabIndex        =   13
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Bank info"
      Height          =   210
      Left            =   315
      TabIndex        =   12
      Top             =   2880
      Width           =   870
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "PAN NO. :"
      Height          =   210
      Left            =   315
      TabIndex        =   11
      Top             =   2025
      Width           =   915
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "City :"
      Height          =   210
      Left            =   315
      TabIndex        =   10
      Top             =   1597
      Width           =   495
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Address :"
      Height          =   210
      Left            =   315
      TabIndex        =   9
      Top             =   747
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "CIN NO. :"
      Height          =   210
      Left            =   315
      TabIndex        =   8
      Top             =   2460
      Width           =   885
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Compnay Name :"
      Height          =   210
      Left            =   315
      TabIndex        =   7
      Top             =   322
      Width           =   1620
   End
End
Attribute VB_Name = "FrmCompanyInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************
'            eBilling System
'             Version 1.0.0
'      Created by Mr. Atanu Maity
'          Date : 21-Aug-2006
'*************************************
' add/edit/delete company details
'      Used Table : company_master
'open the company_master
'display first record in form load
'add edit save delete and navigation
'*************************************

Option Explicit
Dim RS1 As New ADODB.Recordset
Dim Ctl As Control

Private Sub Command1_Click()
    For Each Ctl In Me.Controls
     If TypeName(Ctl) = "TextBox" Then Ctl.Enabled = True
Next Ctl
TxtCompanyName.Enabled = False
TxtAddress1.SetFocus
Me.TxtAddress1.SelStart = Len(Me.TxtAddress1) + 1
Command1.Enabled = False
End Sub

Private Sub Command2_Click()
    '>>> move record ponter to first record
    '>>> display first record
    On Error Resume Next
    If RS1.State = adStateClosed Then Exit Sub
        RS1("address1") = TxtAddress1.Text
        RS1("address2") = TxtAddress2.Text
        RS1("city") = TxtCity.Text
        RS1("pan_no") = TxtPin.Text
        RS1("cin_no") = TxtTelephone.Text
        RS1("bank_info") = TxtVatNo.Text
        RS1("gst_no") = txtgst.Text
        RS1.Update
        RS1.Requery
        RS1.MoveFirst
        MsgBox "Record Updated"
        Call DisplayRecord
        Command2.Enabled = False
        Command1.Enabled = True
        Command1.SetFocus
End Sub



Private Sub Form_Load()
    '>>> center the form
    Me.Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
    '>>> reset connection
    '>>> clear all text
    OpenCon
    '>>> load alreday saved clent data
    '>>> and show the first record
    
    For Each Ctl In Me.Controls
     If TypeName(Ctl) = "TextBox" Then Ctl.Enabled = False
Next Ctl


    If RS1.State = adStateOpen Then RS1.Close
    RS1.Open "select * from company_master ", Cn, adOpenDynamic, adLockOptimistic
    If RS1.RecordCount > 0 Then
        RS1.MoveFirst
        Call DisplayRecord
    End If
End Sub

Private Sub DisplayRecord()
    '>>> display current record
    On Error Resume Next
    TxtCompanyName.Text = IIf(IsNull(RS1("company_name")) = True, "", RS1("company_name"))
    TxtAddress1.Text = IIf(IsNull(RS1("address1")) = True, "", RS1("address1"))
    TxtAddress2.Text = IIf(IsNull(RS1("address2")) = True, "", RS1("address2"))
    TxtCity.Text = IIf(IsNull(RS1("city")) = True, "", RS1("city"))
    TxtPin.Text = IIf(IsNull(RS1("pan_no")) = True, "", RS1("pan_no"))
    TxtTelephone.Text = IIf(IsNull(RS1("cin_no")) = True, "", RS1("cin_no"))
    TxtVatNo.Text = IIf(IsNull(RS1("bank_info")) = True, "", RS1("bank_info"))
    txtgst.Text = IIf(IsNull(RS1("gst_no")) = True, "", RS1("gst_no"))
End Sub


