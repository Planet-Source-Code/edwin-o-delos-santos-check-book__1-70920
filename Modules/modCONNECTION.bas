Attribute VB_Name = "modCONNECTION"

Option Explicit
'Public CN    As  Connection 'user by INVENTORY.MDB
'Public Con   As  Connection 'used by USERS.MDB
'Public CnPay As  Connection 'used by payroll
Public cnBank As Connection  'used by Cash desbursement
Public cnRef As Connection 'use by reference

Public Sub OpenDB(ByRef MDB As String, newConn As Connection, Optional ByVal needPASS As Boolean, Optional ByVal mdbPASS As String)
'// each MDB requires new connection
 Set newConn = New ADODB.Connection  '//microsoft activeX Object 2.1 Library
 newConn.CursorLocation = adUseClient
 If needPASS = False Then
   newConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source= " & App.Path & "\DB\" & MDB
 Else
   newConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=false;Data Source= " & App.Path & "\DB\" & MDB & ";Jet OLEDB:Database Password=" & mdbPASS
 End If
End Sub

Public Sub CloseDB()
 cnBank.Close
 cnRef.Close
 Set cnBank = Nothing
 Set cnRef = Nothing
End Sub

Sub Main()
'[===============================]
'<OPEN OTHER CONNECTION          >
'[===============================]
  Call OpenDB("BANK.MDB", cnBank)
  Call OpenDB("REFERENCE.MDB", cnRef)
  Load frmSPLASH
  frmSPLASH.Show
End Sub

