VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmPrintBill 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Bill"
   ClientHeight    =   4020
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6375
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmPrintBill.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   6375
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox LstInvoiceNo 
      Height          =   2580
      Left            =   240
      TabIndex        =   7
      Top             =   1290
      Width           =   3465
   End
   Begin VB.CommandButton CmdClose 
      Caption         =   "&Close"
      Height          =   450
      Left            =   4230
      TabIndex        =   6
      Top             =   3375
      Width           =   1875
   End
   Begin VB.CommandButton CmdShowBill 
      Caption         =   "&Show/Print Bill"
      Height          =   450
      Left            =   4230
      TabIndex        =   5
      Top             =   2700
      Width           =   1875
   End
   Begin VB.CommandButton CmdGetBill 
      Caption         =   "Get Bill"
      Height          =   450
      Left            =   4230
      TabIndex        =   4
      Top             =   705
      Width           =   1875
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      Height          =   345
      Left            =   1950
      TabIndex        =   1
      Top             =   195
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   609
      _Version        =   393216
      Format          =   116654081
      CurrentDate     =   38951
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      Height          =   345
      Left            =   1935
      TabIndex        =   3
      Top             =   735
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   609
      _Version        =   393216
      Format          =   116654083
      CurrentDate     =   38951
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Invoice Date To"
      Height          =   210
      Left            =   180
      TabIndex        =   2
      Top             =   795
      Width           =   1470
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Invoice Date from "
      Height          =   210
      Left            =   150
      TabIndex        =   0
      Top             =   240
      Width           =   1725
   End
End
Attribute VB_Name = "FrmPrintBill"
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
'     Print already saved bill
'      Used Table : bill
'                 : bill_details
'Module to allow user to select the invoice
'no from seleted date range
'and open the bill by its invoice/bill no
'*************************************

Option Explicit
Dim Rs As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
Dim path As String
Public invno As String


Private Sub CmdClose_Click()
    '>>> close the bill
    Unload Me
End Sub

Private Sub CmdGetBill_Click()
    '>>> get bill no as per selected date range
    '>>> fill the list box from record set
    LstInvoiceNo.Clear
    If Rs.State = adStateOpen Then Rs.Close
    Rs.Open "select invoice_no from bill where invoice_date>=#" & Format(DTPicker1.Value, "dd-mmm-yy") & "# and invoice_date<=#" & Format(DTPicker2.Value, "dd-mmm-yy") & "# order by invoice_no", Cn, adOpenStatic, adLockReadOnly
    While Rs.EOF = False
        LstInvoiceNo.AddItem Rs!invoice_no
        Rs.MoveNext
    Wend
    
End Sub

Private Sub CmdShowBill_Click()
    '>>> get teh bill details from bill table as per selected bill no
    If LstInvoiceNo.ListIndex < 0 Then Exit Sub
        '>>> delete old data from temp table and insert from bill, bill_deatils table
        Cn.Execute "delete from temp_bill_details"
        Cn.Execute "delete from temp_bill"
        Cn.Execute "insert into temp_bill select * from bill where invoice_no=" & Val(LstInvoiceNo.Text) & ""
        Cn.Execute "insert into temp_bill_details select * from entrytbl where invoice_no=" & Val(LstInvoiceNo.Text) & " order BY doc_date ASC"
        '>>> reset the connection
         invno = Val(LstInvoiceNo.Text)
         Call OpenCon
        billprint.Show 1
End Sub
    


Private Sub Form_Load()

    '>>> center form
    Me.Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
    
    LstInvoiceNo.Clear
    '>>> set datepicker to current date
    DTPicker1.Value = Date
    DTPicker2.Value = Date
    '>>> reset connection
    Call OpenCon
End Sub
