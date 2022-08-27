VERSION 5.00
Begin VB.Form csttype 
   Caption         =   "Customer T ype"
   ClientHeight    =   1680
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   Icon            =   "csttype.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   1680
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&CREDIT CUSTOMER"
      Height          =   555
      Left            =   90
      TabIndex        =   1
      Top             =   930
      Width           =   4335
   End
   Begin VB.CommandButton cashbtn 
      Caption         =   "&CASH CUSTOMER"
      Height          =   555
      Left            =   60
      TabIndex        =   0
      Top             =   240
      Width           =   4335
   End
End
Attribute VB_Name = "csttype"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cashbtn_Click()
cash_ClientInfo.Show 1
End Sub

Private Sub Command1_Click()
FrmClientInfo.Show 1
End Sub

Private Sub Form_Load()
Me.Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 6
End Sub
