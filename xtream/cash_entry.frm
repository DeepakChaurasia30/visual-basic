VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form cash_entry 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cash Entry"
   ClientHeight    =   7905
   ClientLeft      =   6210
   ClientTop       =   1035
   ClientWidth     =   9975
   Icon            =   "cash_entry.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7905
   ScaleWidth      =   9975
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Calendar1 
      Height          =   30
      Left            =   5880
      ScaleHeight     =   30
      ScaleWidth      =   30
      TabIndex        =   31
      Top             =   120
      Width           =   30
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Russo One"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7200
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Russo One"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6150
      TabIndex        =   6
      Top             =   7200
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "Russo One"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4230
      TabIndex        =   5
      Top             =   7200
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Edit"
      BeginProperty Font 
         Name            =   "Russo One"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2520
      TabIndex        =   4
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000B&
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "Russo One"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      TabIndex        =   3
      Top             =   7200
      Width           =   1455
   End
   Begin VB.TextBox doctxt 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "Russo One"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4710
      TabIndex        =   1
      Top             =   120
      Width           =   3045
   End
   Begin VB.Frame Frame1 
      Caption         =   "Details"
      BeginProperty Font 
         Name            =   "Permanent Marker"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6255
      Left            =   1080
      TabIndex        =   0
      Top             =   750
      Width           =   8055
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Russo One"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3330
         TabIndex        =   34
         Top             =   600
         Width           =   3705
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Russo One"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3330
         TabIndex        =   32
         Top             =   1950
         Width           =   3105
      End
      Begin VB.CheckBox Check8 
         Caption         =   "Check8"
         Height          =   195
         Left            =   3000
         TabIndex        =   30
         Top             =   5280
         Width           =   255
      End
      Begin VB.CheckBox Check7 
         Caption         =   "Check7"
         Height          =   195
         Left            =   3000
         TabIndex        =   29
         Top             =   4680
         Width           =   255
      End
      Begin VB.CheckBox Check6 
         Caption         =   "Check6"
         Height          =   195
         Left            =   3000
         TabIndex        =   28
         Top             =   3960
         Width           =   255
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Check5"
         Height          =   195
         Left            =   3000
         TabIndex        =   27
         Top             =   3360
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Check4"
         Height          =   195
         Left            =   3000
         TabIndex        =   26
         Top             =   2760
         Width           =   255
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Check3"
         Height          =   195
         Left            =   3000
         TabIndex        =   25
         Top             =   2040
         Width           =   255
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Check2"
         Height          =   195
         Left            =   3000
         TabIndex        =   24
         Top             =   1320
         Width           =   255
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   195
         Left            =   3000
         TabIndex        =   23
         Top             =   720
         Width           =   255
      End
      Begin VB.ComboBox srvtxt 
         BeginProperty Font 
            Name            =   "Orbitron"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3360
         TabIndex        =   21
         Top             =   3840
         Width           =   1455
      End
      Begin VB.ComboBox modetxt 
         BeginProperty Font 
            Name            =   "Orbitron"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3360
         TabIndex        =   20
         Top             =   3240
         Width           =   1455
      End
      Begin VB.ComboBox doctytxt 
         BeginProperty Font 
            Name            =   "Orbitron"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3360
         TabIndex        =   19
         Top             =   2640
         Width           =   1455
      End
      Begin VB.TextBox freighttxt 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Russo One"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3360
         TabIndex        =   18
         Top             =   5760
         Width           =   1335
      End
      Begin VB.TextBox ratetxt 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Russo One"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3360
         TabIndex        =   17
         Top             =   5160
         Width           =   1335
      End
      Begin VB.TextBox weighttxt 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Russo One"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3360
         TabIndex        =   2
         Top             =   4560
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker datetxt 
         Bindings        =   "cash_entry.frx":10CA
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         DataSource      =   "entrydb"
         Height          =   375
         Left            =   3360
         TabIndex        =   33
         Top             =   1290
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Russo One"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   116260865
         CurrentDate     =   44772
      End
      Begin VB.Label Label10 
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Orbitron"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   22
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "Freight"
         BeginProperty Font 
            Name            =   "Orbitron"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         TabIndex        =   16
         Top             =   5760
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "Rate"
         BeginProperty Font 
            Name            =   "Orbitron"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   15
         Top             =   5160
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Weight"
         BeginProperty Font 
            Name            =   "Orbitron"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         TabIndex        =   14
         Top             =   4560
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Service"
         BeginProperty Font 
            Name            =   "Orbitron"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         TabIndex        =   13
         Top             =   3840
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Mode"
         BeginProperty Font 
            Name            =   "Orbitron"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   12
         Top             =   3240
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Dox/ND"
         BeginProperty Font 
            Name            =   "Orbitron"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         TabIndex        =   11
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Destination"
         BeginProperty Font 
            Name            =   "Orbitron"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   10
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Customer"
         BeginProperty Font 
            Name            =   "Orbitron"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   9
         Top             =   600
         Width           =   1575
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Consignment"
      BeginProperty Font 
         Name            =   "Orbitron"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   8
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "cash_entry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
Dim Rs2 As New ADODB.Recordset
Dim Rs3 As New ADODB.Recordset
Dim blnAuto As Boolean 'Keeps the autocomplete functions from
                        'triggering the Change event


Private Sub Command1_Click()
    '>>> close con
    If Rs2.State = adStateOpen Then Rs2.Close
If Command5.Caption = "Update" Then Command5.Caption = "Save"
Command5.Enabled = True
Command4.Enabled = True
Command2.Enabled = False
Command3.Enabled = False
Frame1.Enabled = True
Combo1.SetFocus
Command1.Enabled = False
End Sub

Private Sub Command2_Click()
If Rs2.State = adStateOpen Then Rs2.Close
Command5.Caption = "Update"
If Command5.Caption = "Update" Then
Command4.Enabled = 1
Command3.Enabled = True
Command5.Enabled = True
Command1.Enabled = False
Frame1.Enabled = True
Combo1.SetFocus
End If
End Sub

Private Sub Command3_Click()
If Rs2.State = adStateOpen Then Rs2.Close
    Rs2.Open "select * from cash_entrytbl where doc_no='" & doctxt.Text & "'", Cn, adOpenDynamic, adLockOptimistic
        If MsgBox("Delete the Record ? ", vbCritical + vbYesNo) = vbYes Then
            Rs2.Delete
            Call ClearText
            Frame1.Enabled = False
            Command1.Enabled = False
            Command2.Enabled = False
            Command3.Enabled = False
            Command4.Enabled = False
            Command5.Enabled = False
            doctxt.SetFocus
            End If
End Sub

Private Sub Command4_Click()
Call ClearText
            Frame1.Enabled = False
            Command1.Enabled = False
            Command2.Enabled = False
            Command3.Enabled = False
            Command4.Enabled = False
            Command5.Enabled = False
            doctxt.SetFocus
End Sub

Private Sub Command5_Click()
'>>> verify
    If Rs2.State = adStateOpen Then Rs2.Close
    Rs2.Open "select * from cash_entrytbl where doc_no='" & doctxt.Text & "'", Cn, adOpenDynamic, adLockOptimistic
'>>> viladation for isAdd or isEdit
If Command5.Caption = "Save" Then
     Rs2.AddNew
     Rs2("doc_no") = doctxt.Text
     Rs2("client_name") = Combo1.Text
     Rs2("doc_date") = Format(datetxt.Value, "dd-mmm-yy")
     Rs2("doc_dest") = Combo2.Text
     Rs2("doc_type") = doctytxt.Text
     Rs2("doc_mode") = modetxt.Text
     Rs2("srv_type") = srvtxt.Text
     Rs2("doc_weight") = weighttxt.Text
     Rs2("srv_amt") = freighttxt.Text
     Rs2("entry_date") = Date
     Rs2.Update
            Frame1.Enabled = False
            Command1.Enabled = False
            Command2.Enabled = False
            Command3.Enabled = False
            Command4.Enabled = False
            Command5.Enabled = False
            doctxt.SetFocus
            If Check1.Value = 0 Then Combo1.Text = ""
            If Check2.Value = 0 Then datetxt.Value = Date
            If Check3.Value = 0 Then Combo2.Text = ""
            If Check4.Value = 0 Then doctytxt.Text = ""
            If Check5.Value = 0 Then modetxt.Text = ""
            If Check6.Value = 0 Then srvtxt.Text = ""
            If Check7.Value = 0 Then weighttxt.Text = ""
            If Check8.Value = 0 Then freighttxt.Text = ""
            If IsNumeric(doctxt.Text) Then
            doctxt.Text = doctxt.Text + 1
            Me.doctxt.SelStart = Len(Me.doctxt) + 1
            Else
            '>>> auto increament text + no. last 3 digit
            Dim strtxt As String
            Dim stringLength As Integer
            Dim lendiff As Integer
            stringLength = Len(doctxt.Text)
            lendiff = stringLength - 3
            strtxt = Right(doctxt.Text, 3)
            strtxt = strtxt + 1
            doctxt.Text = Left(doctxt.Text, lendiff) + Right(strtxt, 3)
            Me.doctxt.SelStart = Len(Me.doctxt) + 1
     End If
     Me.doctxt.SetFocus
     ElseIf Command5.Caption = "Update" Then
     Rs2("client_name") = Combo1.Text
     Rs2("doc_date") = Format(datetxt.Value, "dd-mmm-yy")
     Rs2("doc_dest") = Combo2.Text
     Rs2("doc_type") = doctytxt.Text
     Rs2("doc_mode") = modetxt.Text
     Rs2("srv_type") = srvtxt.Text
     Rs2("doc_weight") = weighttxt.Text
     Rs2("srv_amt") = freighttxt.Text
     Rs2.Update
     Rs2.Requery
            Frame1.Enabled = False
            Command1.Enabled = False
            Command2.Enabled = False
            Command3.Enabled = False
            Command4.Enabled = False
            Command5.Enabled = False
            Me.doctxt.SelStart = Len(Me.doctxt) + 1
            doctxt.SetFocus
            Call ClearText
    End If
End Sub

Private Sub Form_Load()
    '>>> show current dates
    datetxt.Value = Date
    
    '>>> reset con
    Call OpenCon

'>>> Add iteams in combo box
If RS1.State = adStateOpen Then RS1.Close
 RS1.Open "select city_name from dest_master where city_name is not null order by city_name", Cn, adOpenStatic, adLockReadOnly
    Do While Not RS1.EOF
    Combo2.AddItem RS1!city_name
    RS1.MoveNext
Loop

    
    If Rs.State = adStateOpen Then Rs.Close
    Rs.Open "select client_name from cash_client_master where client_name is not null order by client_name", Cn, adOpenStatic, adLockReadOnly
    Do While Not Rs.EOF
    Combo1.AddItem Rs!client_name
    Rs.MoveNext
Loop

    modetxt.AddItem "Surface"
    modetxt.AddItem "Air"
    srvtxt.AddItem "Standard"
    srvtxt.AddItem "F.Track"
    doctytxt.AddItem "Dox"
    doctytxt.AddItem "N-Dox"
    Command1.Enabled = False
    Command2.Enabled = False
    Command3.Enabled = False
    Command4.Enabled = False
    Command5.Enabled = False
    
    '>>> disable frame
    Frame1.Enabled = False
End Sub
Private Sub Combo1_Change()
    Dim strPart As String, iLoop As Integer, iStart As Integer, strItem As String
    'don't do if no text or if change was made by autocomplete coding
    If Not blnAuto And Combo1.Text <> "" Then
        'save the selection start point (cursor position)
        iStart = Combo1.SelStart
        'get the part the user has typed (not selected)
        strPart = Left$(Combo1.Text, iStart)
        For iLoop = 0 To Combo1.ListCount - 1
            'compare each item to the part the user has typed,
            '"complete" with the first good match
            strItem = UCase$(Combo1.List(iLoop))
            If strItem Like UCase$(strPart & "*") And _
                    strItem <> UCase$(Combo1.Text) Then
                'partial match but not the whole thing.
                '(if whole thing, nothing to complete!)
                blnAuto = True
                Combo1.SelText = Mid$(Combo1.List(iLoop), iStart + 1) 'add on the new ending
                Combo1.SelStart = iStart   'reset the selection
                Combo1.SelLength = Len(Combo1.Text) - iStart
                blnAuto = False
                Exit For
            End If
        Next iLoop
    
        'Add statement here like FilterRecord or whatever :)
    End If
End Sub

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
    'Unless we watch out for it, backspace or delete will just delete
    'the selected text (the autocomplete part), so we delete it here
    'first so it doesn't interfere with what the user expects
    If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
        blnAuto = True
        Combo1.SelText = ""
        blnAuto = False
    ElseIf KeyCode = vbKeyReturn Then 'Accept autocomplete on 'Enter' keypress
        Combo1_LostFocus
        'the following causes the item to be selected and
        'the cursor placed at the end:
        Combo1.SelStart = Len(Combo1.Text)
        
        'This would select the whole thing instead:
        'cboName.SelLength = Len(cboName.Text)
    
        'alternatively, you could move the focus to the next control here
    End If
End Sub
Private Sub Combo1_LostFocus()

Dim iLoop As Integer
'Match capitalization if item entered is one on the list
    If Combo1.Text <> "" Then
        For iLoop = 0 To Combo1.ListCount - 1
            If UCase$(Combo1.List(iLoop)) = UCase$(Combo1.Text) Then
                blnAuto = True
                Combo1.Text = Combo1.List(iLoop)
                blnAuto = False
                Exit For
            End If
        Next iLoop
    End If
   
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 9 Or KeyAscii = 13 Then
If Combo1.Text = "" Then
MsgBox "Can't be empty", vbOKOnly
Combo1.SetFocus
Else
Combo2.SetFocus
End If
End If
End Sub

Private Sub Combo2_Change()
    Dim strPart As String, iLoop As Integer, iStart As Integer, strItem As String
    'don't do if no text or if change was made by autocomplete coding
    If Not blnAuto And Combo2.Text <> "" Then
        'save the selection start point (cursor position)
        iStart = Combo2.SelStart
        'get the part the user has typed (not selected)
        strPart = Left$(Combo2.Text, iStart)
        For iLoop = 0 To Combo2.ListCount - 1
            'compare each item to the part the user has typed,
            '"complete" with the first good match
            strItem = UCase$(Combo2.List(iLoop))
            If strItem Like UCase$(strPart & "*") And _
                    strItem <> UCase$(Combo2.Text) Then
                'partial match but not the whole thing.
                '(if whole thing, nothing to complete!)
                blnAuto = True
                Combo2.SelText = Mid$(Combo2.List(iLoop), iStart + 1) 'add on the new ending
                Combo2.SelStart = iStart   'reset the selection
                Combo2.SelLength = Len(Combo2.Text) - iStart
                blnAuto = False
                Exit For
            End If
        Next iLoop
    
        'Add statement here like FilterRecord or whatever :)
    End If
End Sub

Private Sub Combo2_KeyDown(KeyCode As Integer, Shift As Integer)
    'Unless we watch out for it, backspace or delete will just delete
    'the selected text (the autocomplete part), so we delete it here
    'first so it doesn't interfere with what the user expects
    If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
        blnAuto = True
        Combo2.SelText = ""
        blnAuto = False
    ElseIf KeyCode = vbKeyReturn Then 'Accept autocomplete on 'Enter' keypress
        Combo2_LostFocus
        'the following causes the item to be selected and
        'the cursor placed at the end:
        Combo2.SelStart = Len(Combo2.Text)
        
        'This would select the whole thing instead:
        'cboName.SelLength = Len(cboName.Text)
    
        'alternatively, you could move the focus to the next control here
    End If
End Sub
Private Sub Combo2_LostFocus()
Dim iLoop As Integer
'Match capitalization if item entered is one on the list
    If Combo2.Text <> "" Then
        For iLoop = 0 To Combo2.ListCount - 1
            If UCase$(Combo2.List(iLoop)) = UCase$(Combo2.Text) Then
                blnAuto = True
                Combo2.Text = Combo2.List(iLoop)
                blnAuto = False
                Exit For
            End If
        Next iLoop

    End If
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 9 Or KeyAscii = 13 Then
If Combo1.Text = "" Then
MsgBox "Can't be empty", vbOKOnly
Combo1.SetFocus
Else
doctytxt.SetFocus
End If
End If
End Sub


'>>> main entry codes start
Private Sub doctxt_KeyPress(KeyAscii As Integer)
If KeyAscii = 9 Or KeyAscii = 13 Then
    If Rs2.State = adStateOpen Then Rs2.Close
    Rs2.Open "select * from cash_entrytbl where doc_no='" & doctxt.Text & "'", Cn, adOpenStatic, adLockReadOnly
    If Rs2.RecordCount > 0 Then
    Call ClearText
    blnAuto = True
    Combo1.Text = Rs2("client_name")
    datetxt.Value = Rs2("doc_date")
    Combo2.Text = Rs2("doc_dest")
    doctytxt.Text = Rs2("doc_type")
    modetxt.Text = Rs2("doc_mode")
    srvtxt.Text = Rs2("srv_type")
    weighttxt.Text = Rs2("doc_weight")
    freighttxt.Text = Rs2("srv_amt")
    Command2.Enabled = True
    Command2.SetFocus
    Command1.Enabled = False
    ElseIf Rs2.RecordCount <= 0 Then
    Command2.Enabled = False
    Command3.Enabled = False
    Command4.Enabled = False
    Command5.Enabled = False
    Command1.Enabled = True
    Command1.SetFocus
    If Command5.Caption = "Update" Then Command5.Caption = "Save"
    End If
    End If
    Exit Sub
    End Sub

   Private Sub ClearText()
    '>>> clear all text box in the form
    Dim Ctl As Control
    For Each Ctl In Me.Controls
            Select Case TypeName(Ctl)
                Case Is = "TextBox"
                    If Ctl.Name <> "doctxt" Then Ctl.Text = ""
                    Case Is = "ComboBox"
                    Ctl.Text = ""
                    Combo1.ListIndex = -1
            End Select
        Next Ctl
End Sub


Private Sub ratetxt_KeyPress(KeyAscii As Integer)
Dim rate As String
If KeyAscii = 9 Or KeyAscii = 13 Then
freighttxt.Text = ""
rate = Val(weighttxt.Text) * Val(ratetxt.Text)
freighttxt.Text = rate
Me.freighttxt.SetFocus
Me.freighttxt.SelStart = Len(Me.freighttxt) + 1
End If
End Sub

'>>> lost focus function define
Private Sub doctxt_LostFocus()
If doctxt.Text = "" Then
MsgBox "Doc can't be empty", vbOKOnly
doctxt.SetFocus
Frame1.Enabled = False
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
Command5.Enabled = False
End If
End Sub

Private Sub doctytxt_KeyPress(KeyAscii As Integer)
If KeyAscii = 9 Or KeyAscii = 13 Then
If doctytxt.Text = "" Then
MsgBox "Can't be empty", vbOKOnly
doctytxt.SetFocus
Else
modetxt.SetFocus
End If
End If
End Sub

Private Sub modetxt_KeyPress(KeyAscii As Integer)
If KeyAscii = 9 Or KeyAscii = 13 Then
If modetxt.Text = "" Then
MsgBox "Can't be empty", vbOKOnly
modetxt.SetFocus
Else
srvtxt.SetFocus
End If
End If
End Sub

Private Sub srvtxt_KeyPress(KeyAscii As Integer)
If KeyAscii = 9 Or KeyAscii = 13 Then
If srvtxt.Text = "" Then
MsgBox "Can't be empty", vbOKOnly
srvtxt.SetFocus
Else
weighttxt.SetFocus
Me.weighttxt.SelStart = Len(Me.weighttxt) + 1
End If
End If
End Sub

Private Sub weighttxt_KeyPress(KeyAscii As Integer)
If KeyAscii = 9 Or KeyAscii = 13 Then
If weighttxt.Text = "" Then
MsgBox "Can't be empty", vbOKOnly
weighttxt.SetFocus
Else
ratetxt.SetFocus
Me.ratetxt.SelStart = Len(Me.ratetxt) + 1
End If
End If
End Sub

Private Sub freighttxt_KeyPress(KeyAscii As Integer)
If KeyAscii = 9 Or KeyAscii = 13 Then
If freighttxt.Text = "" Then
MsgBox "Can't be empty", vbOKOnly
freighttxt.SetFocus
Else
Command5.SetFocus
End If
End If
End Sub



    

