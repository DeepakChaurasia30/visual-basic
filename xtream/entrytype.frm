VERSION 5.00
Begin VB.Form entrytype 
   Caption         =   "Entry T ype"
   ClientHeight    =   1590
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4605
   Icon            =   "entrytype.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   1590
   ScaleWidth      =   4605
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton creditbtn 
      Caption         =   "&CREDIT ENTRY"
      Height          =   555
      Left            =   150
      TabIndex        =   1
      Top             =   930
      Width           =   4335
   End
   Begin VB.CommandButton cashbtn 
      Caption         =   "&CASH ENTRY"
      Height          =   555
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   4335
   End
End
Attribute VB_Name = "entrytype"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cashbtn_Click()
cash_entry.Show 1
End Sub

Private Sub creditbtn_Click()
entry.Show 1
End Sub

Private Sub Form_Load()
Me.Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 6
End Sub
