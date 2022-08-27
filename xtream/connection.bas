Attribute VB_Name = "connection"

'*************************************
'            Billing System
'             Version 1.0.0
'      Created by Mr. Deepak Chaurasia
'          Date : 01-Aug-2022
'*************************************
'
'declare global variable and procedure
'
'*************************************
Option Explicit

Public Cn As New ADODB.connection


Public Sub OpenCon()
    '>>> open connction
    If Cn.State = 1 Then Cn.Close
    Cn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;User ID=admin;Data Source=" & App.path & "\donttouch.mdb;Jet OLEDB:Database Password=xtream@30"
    Cn.CursorLocation = adUseClient
    Cn.Open
End Sub

Public Function newinv(ByVal table As String) As Integer
    '>>> find new inv no.
    Dim Rs As New ADODB.Recordset
    If Rs.State = 1 Then Rs.Close
    Rs.Open "select max(invoice_no) from bill", Cn, adOpenStatic, adLockReadOnly
    If IsNull(Rs(0)) = True Then
        newinv = 1
    Else
        newinv = Val(Rs(0)) + 1
    End If

End Function

