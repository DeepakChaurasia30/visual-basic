VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form invoice 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Invoice Gen.."
   ClientHeight    =   4005
   ClientLeft      =   7770
   ClientTop       =   1350
   ClientWidth     =   6960
   Icon            =   "invoice.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   6960
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Russo One"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2550
      TabIndex        =   13
      Top             =   2280
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   330
      Left            =   2550
      TabIndex        =   12
      Top             =   2760
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Russo One"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   116195329
      CurrentDate     =   44783
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   345
      Left            =   2580
      TabIndex        =   11
      Top             =   1650
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Russo One"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   116195329
      CurrentDate     =   44783
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   330
      Left            =   2580
      TabIndex        =   10
      Top             =   1140
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Russo One"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   116195329
      CurrentDate     =   44783
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Russo One"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2610
      TabIndex        =   9
      Top             =   510
      Width           =   3375
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Clear"
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
      Left            =   3720
      TabIndex        =   8
      Top             =   3450
      Width           =   1245
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Close"
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
      Left            =   5280
      TabIndex        =   7
      Top             =   3450
      Width           =   1425
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Delete"
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
      Left            =   1860
      TabIndex        =   6
      Top             =   3450
      Width           =   1515
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Generate"
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
      Left            =   330
      TabIndex        =   5
      Top             =   3420
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Invoice Date"
      BeginProperty Font 
         Name            =   "Russo One"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   540
      TabIndex        =   4
      Top             =   2760
      Width           =   1635
   End
   Begin VB.Label Label4 
      Caption         =   "Invoice no."
      BeginProperty Font 
         Name            =   "Russo One"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   570
      TabIndex        =   3
      Top             =   2220
      Width           =   1725
   End
   Begin VB.Label Label3 
      Caption         =   "Date To"
      BeginProperty Font 
         Name            =   "Russo One"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   570
      TabIndex        =   2
      Top             =   1650
      Width           =   1245
   End
   Begin VB.Label Label2 
      Caption         =   "Date From"
      BeginProperty Font 
         Name            =   "Russo One"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   570
      TabIndex        =   1
      Top             =   1110
      Width           =   1305
   End
   Begin VB.Label Label1 
      Caption         =   "Customer"
      BeginProperty Font 
         Name            =   "Russo One"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   570
      TabIndex        =   0
      Top             =   540
      Width           =   1305
   End
End
Attribute VB_Name = "invoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ninv As String
Dim Rs As New ADODB.Recordset
Dim Rs2 As New ADODB.Recordset
Dim Rs3 As New ADODB.Recordset
Dim Rs4 As New ADODB.Recordset
Dim Rs5 As New ADODB.Recordset
Dim blnAuto As Boolean 'Keeps the autocomplete functions from
                        'triggering the Change event
Dim tamt As Double
Dim samt As Double
Dim gst As Double
'Dim cgst As String
Dim famt As Double



Private Sub Command1_Click()
Command1.Enabled = False
If Combo1.Text = "" Then
Combo1.SetFocus
Else
'>>> entry invoice no in entrytbl
     If Rs3.State = adStateOpen Then Rs3.Close
    Rs3.Open "update [entrytbl] set [invoice_no]='" & Text1.Text & "'  where doc_date>=#" & Format(DTPicker1.Value, "dd-mmm-yy") & "# and doc_date<=#" & Format(DTPicker2.Value, "dd-mmm-yy") & "# and invoice_no is null and client_name='" & Combo1.Text & "'", Cn, adOpenDynamic, adLockOptimistic

'>>> sum up the records founds
If Rs2.State = adStateOpen Then Rs2.Close
Rs2.Open "select SUM(srv_amt) As s_amt from entrytbl where invoice_no=" & Val(Text1.Text) & "", Cn, adOpenDynamic, adLockOptimistic
If IsNull(Rs2![s_amt]) = False Then

'>>> get gst and fuel information
If Rs4.State = adStateOpen Then Rs4.Close
Rs4.Open "select * from client_master where client_name='" & Combo1.Text & "'", Cn, adOpenDynamic, adLockOptimistic
samt = Rs2![s_amt]
famt = Rs4![f_perc] / 100 * samt
gst = (9 / 100) * (samt + famt)
tamt = samt + famt + (gst * 2)
RsWord (Val(tamt))

'>>> Add them in bill table to save as records
If Rs5.State = adStateOpen Then Rs5.Close
Rs5.Open "select * from bill ", Cn, adOpenDynamic, adLockOptimistic
     Rs5.AddNew
     Rs5("invoice_no") = Val(Text1.Text)
     Rs5("client_name") = Combo1.Text
     Rs5("invoice_date") = Format(DTPicker3.Value, "dd-mmm-yy")
     Rs5("bill_fdate") = Format(DTPicker1.Value, "dd-mmm-yy")
     Rs5("bill_tdate") = Format(DTPicker2.Value, "dd-mmm-yy")
     Rs5("total_amt") = Val(tamt)
     Rs5("amt_word") = RsWord(Val(tamt))
     Rs5("fuel_amt") = famt
     Rs5("sgst_amt") = gst
     Rs5("cgst_amt") = gst
     Rs5("sub_amt") = samt
     '>>> temprary fix
     Rs5("igst_amt") = (gst * 2)
     Rs5("dev") = "xox"
     Rs5.Update
     Rs5.Requery
     MsgBox "Bill Generated", vbOKOnly
ElseIf IsNull(Rs2![s_amt]) = True Then
MsgBox "No Record found in given date Range", vbOKOnly
End If
End If
End Sub

Private Sub Command2_Click()
'>>> delting record
If MsgBox("Delete the Invoice ? ", vbCritical + vbYesNo) = vbYes Then
Cn.Execute "delete from bill where invoice_no=" & Val(Text1.Text) & ""
Cn.Execute "update [entrytbl] set invoice_no = NULL where invoice_no=" & Val(Text1.Text) & ""
MsgBox "Invoice Deleted ", vbOKOnly
'>>> reset connection
Call OpenCon
Command2.Enabled = False
End If
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command4_Click()
Call ClearText
End Sub

Private Sub Form_Load()

'>>> cnter the form
Me.Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 6

DTPicker1.Value = Date
DTPicker2.Value = Date
DTPicker3.Value = Date

Command1.Enabled = False
Command2.Enabled = False

'>>> call opencon
Call OpenCon

'>>> new inv no
Text1.Text = newinv(ninv)

'>>> Add iteams to combobox
    If Rs.State = adStateOpen Then Rs.Close
    Rs.Open "select client_name from client_master where client_name is not null order by client_name", Cn, adOpenStatic, adLockReadOnly
    Do While Not Rs.EOF
    Combo1.AddItem Rs!client_name
    Rs.MoveNext
Loop
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

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 9 Or KeyAscii = 13 Then
    If Rs2.State = adStateOpen Then Rs2.Close
    Rs2.Open "select * from bill where invoice_no=" & Val(Text1.Text) & "", Cn, adOpenStatic, adLockReadOnly
    If Rs2.RecordCount > 0 Then
    Command1.Enabled = False
    MsgBox "A bill already associated with this invoice no. to " & Rs2("client_name"), vbInformation + vbOKOnly
    Command2.Enabled = True
    Command2.SetFocus
    ElseIf Rs2.RecordCount <= 0 Then
    If Text1.Text = "" Then
    Command1.Enabled = False
    Else
    Command1.Enabled = True
End If
End If
End If
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
            End Select
        Next Ctl
End Sub

