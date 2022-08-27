VERSION 5.00
Begin VB.Form centre 
   Caption         =   "Centre"
   ClientHeight    =   2430
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4140
   Icon            =   "centre.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   2430
   ScaleWidth      =   4140
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Add"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   750
      TabIndex        =   1
      Top             =   1560
      Width           =   2385
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   630
      TabIndex        =   0
      Top             =   660
      Width           =   2775
   End
End
Attribute VB_Name = "centre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As New ADODB.Recordset

Private Sub Command1_Click()
If Rs.State = adStateOpen Then Rs.Close
    Rs.Open "select * from dest_master", Cn, adOpenDynamic, adLockOptimistic
    Rs.AddNew
     Rs("city_name") = Text1.Text
     Rs.Update
     Rs.Requery
     MsgBox "Centre Added"
     Text1.Text = ""
End Sub

Private Sub Form_Load()
    '>>> center the form
    Me.Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
End Sub
