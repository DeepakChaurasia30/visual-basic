VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "Xtream"
   ClientHeight    =   4680
   ClientLeft      =   5805
   ClientTop       =   600
   ClientWidth     =   11565
   LinkTopic       =   "MDIForm1"
   Begin VB.Menu Company 
      Caption         =   "&Company"
   End
   Begin VB.Menu Billing 
      Caption         =   "&Billing"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Billing_Click()
entry.Show
End Sub

Private Sub MDIForm_Load()
WindowState = vbMaximized
End Sub
