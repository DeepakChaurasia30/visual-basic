VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form nbill 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "No Bill Report"
   ClientHeight    =   2760
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4845
   Icon            =   "nbill.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   4845
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Show"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1560
      TabIndex        =   4
      Top             =   2160
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   315
      Left            =   1920
      TabIndex        =   1
      Top             =   1350
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Russo One"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   115867649
      CurrentDate     =   44789
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   1920
      TabIndex        =   0
      Top             =   780
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Russo One"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   115867649
      CurrentDate     =   44789
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FF8080&
      Caption         =   "No Bill Report"
      BeginProperty Font 
         Name            =   "Russo One"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1320
      TabIndex        =   5
      Top             =   270
      Width           =   2115
   End
   Begin VB.Label Label2 
      Caption         =   "Date To :"
      BeginProperty Font 
         Name            =   "Russo One"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   480
      TabIndex        =   3
      Top             =   1380
      Width           =   1155
   End
   Begin VB.Label Label1 
      Caption         =   "Date From :"
      BeginProperty Font 
         Name            =   "Russo One"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   450
      TabIndex        =   2
      Top             =   840
      Width           =   1305
   End
End
Attribute VB_Name = "nbill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim fdate As String
Dim tdate As String
Cn.Execute "delete from temp_bill_details"
Cn.Execute "insert into temp_bill_details select * from entrytbl where doc_date>=#" & Format(DTPicker1.Value, "dd-mmm-yy") & "# and doc_date<=#" & Format(DTPicker2.Value, "dd-mmm-yy") & "# and invoice_no is null order BY client_name ASC"
Call OpenCon
nbillrpt.Show 1
End Sub

Private Sub Form_Load()

    '>>> center form
    Me.Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
    
    '>>> set datepicker to current date
    DTPicker1.Value = Date
    DTPicker2.Value = Date
    '>>> reset connection
    Call OpenCon
End Sub
