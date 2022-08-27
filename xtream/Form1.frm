VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   3015
   ClientLeft      =   6045
   ClientTop       =   4920
   ClientWidth     =   7125
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   7125
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Orbitron"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   5
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Login"
      BeginProperty Font 
         Name            =   "Orbitron"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   4
      Top             =   2160
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Orbitron"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   2880
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1440
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Orbitron"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   2
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "PASSWORD"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   1
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label USERNAME 
      Caption         =   "USERNAME"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   600
      Width           =   1875
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************
'             Login Module
'      Used Table : logintbl
'Module to check user login and load
'user rights as per user type.
'*************************************

Option Explicit
Dim Rs As New ADODB.Recordset

'enter key to shift focus
Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 9 Or KeyAscii = 13 Then Me.Text2.SetFocus
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 9 Or KeyAscii = 13 Then Me.Command1.SetFocus
End Sub

Private Sub Command1_Click()
    '>>> check wheather user name and password are blank
    '>>> if its is blan warn user to enter
    If Text1.Text = "" Or Text2.Text = "" Then
        MsgBox "Enter user name and password ...", vbExclamation
        Text1.SetFocus
        Exit Sub
    End If
    
    '>>> check for username and password
    '>>> query to user_master with user_name and password
    If Rs.State = adStateOpen Then Rs.Close
    Rs.Open "select * from logintbl where user ='" & Text1.Text & "' and password ='" & Text2 & "'", Cn, adOpenStatic, adLockReadOnly
    If Rs.RecordCount > 0 Then
        FrmMain.Show
        Unload Me
        Else
        MsgBox "Invalid User Name and Password ... ", vbExclamation, "Login Error "
        Text2.Text = ""
        Text1.SetFocus
        Exit Sub
    End If

End Sub

Private Sub Command2_Click()
    '>>> close the application
    End
    Set Form1 = Nothing
End Sub


Private Sub Form_Load()
    '>>> center the form
    Me.Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
   If Cn.State = 1 Then Cn.Close
    OpenCon
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    '>>> release all the object variable used by form
    Set Form1 = Nothing
End Sub

