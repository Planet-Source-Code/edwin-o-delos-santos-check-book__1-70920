VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmReference 
   Caption         =   "Reference Entry Form"
   ClientHeight    =   6510
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6330
   Icon            =   "frmReference.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   6330
   StartUpPosition =   2  'CenterScreen
   Begin CHECKBOOK.Hline Hline2 
      Height          =   30
      Left            =   240
      TabIndex        =   27
      Top             =   600
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   53
   End
   Begin VB.TextBox txtEntry 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   325
      Index           =   1
      Left            =   1920
      TabIndex        =   20
      Top             =   840
      Width           =   3315
   End
   Begin VB.TextBox txtEntry 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   325
      Index           =   3
      Left            =   1920
      TabIndex        =   19
      Top             =   1560
      Width           =   3915
   End
   Begin VB.TextBox txtEntry 
      Appearance      =   0  'Flat
      Height          =   325
      Index           =   2
      Left            =   1920
      TabIndex        =   18
      Top             =   1200
      Width           =   3915
   End
   Begin VB.TextBox txtEntry 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   325
      Index           =   4
      Left            =   1920
      TabIndex        =   17
      Top             =   1920
      Width           =   3915
   End
   Begin VB.ComboBox CboTable 
      Height          =   315
      ItemData        =   "frmReference.frx":1CCA
      Left            =   120
      List            =   "frmReference.frx":1CD4
      TabIndex        =   9
      Text            =   "Select Reference Table..."
      Top             =   120
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CheckBox ChkSelect 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Select Reference Table"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   2895
   End
   Begin VB.PictureBox Hline1 
      Height          =   30
      Left            =   120
      ScaleHeight     =   30
      ScaleWidth      =   6015
      TabIndex        =   8
      Top             =   600
      Width           =   6015
   End
   Begin VB.PictureBox Picture1 
      Height          =   525
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   6195
      TabIndex        =   0
      Top             =   2400
      Width           =   6255
      Begin VB.CommandButton cmdButton 
         BackColor       =   &H00E0E0E0&
         Height          =   495
         Index           =   1
         Left            =   1080
         Picture         =   "frmReference.frx":1CE7
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Save New"
         Top             =   0
         Width           =   495
      End
      Begin VB.PictureBox PicNav 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2040
         ScaleHeight     =   270
         ScaleWidth      =   2385
         TabIndex        =   12
         Top             =   120
         Width           =   2415
         Begin VB.TextBox txtEntry 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E8FBF9&
            BorderStyle     =   0  'None
            ForeColor       =   &H00008080&
            Height          =   285
            Index           =   0
            Left            =   720
            Locked          =   -1  'True
            TabIndex        =   28
            Top             =   0
            Width           =   915
         End
         Begin VB.CommandButton CmdLast 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   275
            Left            =   2040
            MousePointer    =   99  'Custom
            Picture         =   "frmReference.frx":2451
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Last"
            Top             =   0
            Width           =   375
         End
         Begin VB.CommandButton CmdNext 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   275
            Left            =   1680
            MousePointer    =   99  'Custom
            Picture         =   "frmReference.frx":2706
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Next"
            Top             =   0
            Width           =   375
         End
         Begin VB.CommandButton CmdPrev 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   275
            Left            =   360
            MousePointer    =   99  'Custom
            Picture         =   "frmReference.frx":29BB
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Previous"
            Top             =   0
            Width           =   375
         End
         Begin VB.CommandButton CmdFirst 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   0
            MaskColor       =   &H00404040&
            MousePointer    =   99  'Custom
            Picture         =   "frmReference.frx":2C70
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "First"
            Top             =   0
            Width           =   375
         End
      End
      Begin VB.CommandButton cmdButton 
         BackColor       =   &H00E0E0E0&
         Height          =   495
         Index           =   6
         Left            =   5760
         Picture         =   "frmReference.frx":2F25
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Refresh"
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton cmdButton 
         BackColor       =   &H00E0E0E0&
         Height          =   495
         Index           =   5
         Left            =   5280
         Picture         =   "frmReference.frx":368F
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Delete"
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton cmdButton 
         BackColor       =   &H00E0E0E0&
         Height          =   495
         Index           =   3
         Left            =   1080
         Picture         =   "frmReference.frx":3DF9
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Update "
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton cmdButton 
         BackColor       =   &H00E0E0E0&
         Height          =   495
         Index           =   2
         Left            =   600
         Picture         =   "frmReference.frx":4563
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Edit"
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton cmdButton 
         BackColor       =   &H00E0E0E0&
         Height          =   495
         Index           =   4
         Left            =   4800
         Picture         =   "frmReference.frx":4CCD
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Cancel"
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton cmdButton 
         BackColor       =   &H00E0E0E0&
         Height          =   495
         Index           =   0
         Left            =   120
         Picture         =   "frmReference.frx":5437
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Add"
         Top             =   0
         Width           =   495
      End
   End
   Begin MSComctlLib.ImageList i16x16 
      Left            =   3360
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   22
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReference.frx":5BA1
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReference.frx":65B3
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReference.frx":670D
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReference.frx":6867
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReference.frx":69C1
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReference.frx":6CBB
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReference.frx":7055
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReference.frx":73EF
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReference.frx":7E01
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReference.frx":7E55
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReference.frx":81EF
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReference.frx":8589
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReference.frx":8923
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReference.frx":8CBD
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReference.frx":96CF
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReference.frx":A0E1
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReference.frx":AAF3
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReference.frx":B505
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReference.frx":BF17
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReference.frx":C929
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReference.frx":D33B
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReference.frx":D8D7
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvList 
      Height          =   3255
      Left            =   0
      TabIndex        =   26
      Top             =   3120
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   5741
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "i16x16"
      SmallIcons      =   "i16x16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Image ImgItemHelp 
      Height          =   360
      Left            =   5400
      MouseIcon       =   "frmReference.frx":DE73
      MousePointer    =   99  'Custom
      Picture         =   "frmReference.frx":E73D
      Top             =   840
      Width           =   360
   End
   Begin VB.Label lblFLDi 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   3
      Left            =   360
      TabIndex        =   24
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label lblFLDi 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   2
      Left            =   360
      TabIndex        =   23
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label lblFLDi 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   1
      Left            =   360
      TabIndex        =   22
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label lblFLDi 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   4
      Left            =   360
      TabIndex        =   21
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label lblTABLE 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reference"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   4920
      TabIndex        =   11
      Top             =   240
      Width           =   900
   End
   Begin VB.Label lblFLDi 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Do not delete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   0
      Left            =   480
      TabIndex        =   25
      Top             =   840
      Visible         =   0   'False
      Width           =   1335
   End
End
Attribute VB_Name = "frmReference"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim rcdSet As Recordset
Dim rsTemp As Recordset


Private Sub CboTable_Click()
 lvList.ListItems.Clear
 lblTABLE.Caption = CboTable.Text
 ChkSelect.Value = 0
 '//
  On Error GoTo ERRORHANDLE
    Dim sqlStr As String
    Dim lblStr As String

    Set rcdSet = New ADODB.Recordset
    rcdSet.CursorLocation = adUseClient
    sqlStr = "SELECT * FROM [" & lblTABLE.Caption & "]"
    rcdSet.Open sqlStr, cnRef, adOpenStatic, adLockOptimistic
    Load_DATA
    Call ShowFldsLabel(Me, rcdSet)

    cmdButtonShow ("1012011"), Me
    TextBox_Visible Me, rcdSet

ERRORHANDLE:
    errorMsg Err, Me.Name, "CboTables_Click()"
End Sub

Private Sub ChkSelect_Click()
 CboTable.Visible = (ChkSelect.Value = 1)
End Sub

Private Sub nextNumber()
      Dim nextNo As Long
      Set rsTemp = New ADODB.Recordset
      Set rsTemp = rcdSet
      If rsTemp.State = adStateOpen Then
        rsTemp.Close
      End If
      rsTemp.Open "SELECT * From [" & lblTABLE.Caption & "]"
      nextNo = Last_Recc(rsTemp)
      If nextNo > 0 Then
       txtEntry(0).Text = nextNo
       txtEntry(1).SetFocus
      Else
       nextNo = nextNo = 1
       txtEntry(0).Text = nextNo
       txtEntry(1).SetFocus
      End If
      Set rsTemp = Nothing
End Sub


Private Sub cmdButton_Click(Index As Integer)
'//                  A S E U C D R
'On Error GoTo ERRORHANDLE
Dim sqlStr As String
'Dim nextNo As Long
Select Case Index
   Case BtnAdd                       '<------ add new record ------->'
      clearText
      addRec = True
      cmdButtonShow ("0102100"), Me
      nextNumber
   Case BtnSave                       '<------ save new record ------>'
        cmdButtonShow ("1012011"), Me
        If rcdSet Is Nothing Then Exit Sub
        Call WriteData(Me, rcdSet, True)
        Call PopulateData(lvList, rcdSet, 2, txtEntry(0).Text)
        addRec = False
   Case BtnEdit                         '<------ edit record ---------->'
          '//VerifyEdit
          Dim src As String  'source
          Dim trg As String  'targer match
          src = txtEntry(0).Text
          trg = lvList.SelectedItem.Text
         If Trim$(src) <> Trim$(trg) Then
            MsgBox "Please Verify Selected Item# !", vbCritical, "Warning ! [Edit Mode]"
             cmdButtonShow ("1012011"), Me
           editRec = False
           Exit Sub
         End If
     '//
        editRec = True
        cmdButtonShow ("0201100"), Me
        txtEntry(1).SetFocus
   Case BtnUpdate                     '<------ update record -------->'
        cmdButtonShow ("1012011"), Me
        If rcdSet Is Nothing Then Exit Sub
        Call WriteData(Me, rcdSet, False)
        LvwReplaceData Me, rcdSet, lvList
        editRec = False
   Case BtnCancel                     '<------ cancel update -------->'
        cmdButtonShow ("1012011"), Me
        addRec = False
        editRec = False
   Case BtnDelete                     '<------ delete record -------->'
        Call Delete_Record(rcdSet, lvList)
   Case BtnRefresh                    '<------ Refresh record ------->'
        addRec = False
        editRec = False
       If rcdSet Is Nothing Then Exit Sub
       If rcdSet.State = adStateOpen Then
          rcdSet.Close
        End If
         sqlStr = "SELECT * FROM [" & lblTABLE.Caption & "]"
         rcdSet.Open sqlStr & " order by SN", cnRef, adOpenStatic, adLockOptimistic
        Load_DATA
        lvList.SetFocus
End Select
'ERRORHANDLE:
' errorMsg Err, Me.Name, "Command Button"

End Sub

Private Sub CmdFirst_Click()
If rcdSet Is Nothing Then Exit Sub
   rcdSet.MoveFirst
 Call BindDatasource(Me, rcdSet, lvList, False)
Call ListView_Search(lvList, txtEntry(0).Text, 0)
End Sub

Private Sub CmdLast_Click()
If rcdSet Is Nothing Then Exit Sub
 rcdSet.MoveLast
 Call BindDatasource(Me, rcdSet, lvList, False)
Call ListView_Search(lvList, txtEntry(0).Text, 0)
End Sub

Private Sub CmdNext_Click()
 If rcdSet Is Nothing Then Exit Sub
 If rcdSet.EOF = True Then rcdSet.MoveLast
 rcdSet.MoveNext
 Call BindDatasource(Me, rcdSet, lvList, False)
Call ListView_Search(lvList, txtEntry(0).Text, 0)
End Sub

Private Sub CmdPrev_Click()
If rcdSet Is Nothing Then Exit Sub
If rcdSet.BOF = True Then rcdSet.MoveFirst
rcdSet.MovePrevious
Call BindDatasource(Me, rcdSet, lvList, False)
Call ListView_Search(lvList, txtEntry(0).Text, 0)
End Sub




Private Sub Form_Load()
  CboTable.Clear
  cmdButtonShow ("0000000"), Me
  LoadTables
  Call OpenDB("REFERENCE.MDB", cnRef)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If MsgBox("Exit this system?", vbYesNo + vbQuestion, "References") = vbNo Then
  Cancel = 1 'true
Else
  Unload Me
End If

End Sub

Private Sub Form_Resize()
With Me
  If .WindowState = 0 Then
   .Height = 7020
   .Width = 6390
  End If
End With
 SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE

End Sub

Private Sub LoadTables()

    Dim db As Database
    Dim qdef As QueryDef
    Dim td As TableDef
    Dim dbname As String

    ' Open the database. replace "c:\DBfile.mdb" with your
    ' database file name
    
    Set db = OpenDatabase(App.Path & "\DB\REFERENCE.mdb")
    ' List the table names.
    For Each td In db.TableDefs
    ' if you want to display also the system tables, replace the line
    ' below with:  List1.AddItem td.Name
       If td.Attributes = 0 Then CboTable.AddItem td.Name
    Next td
    db.Close
End Sub


Private Sub Load_DATA()
On Error GoTo ERRORHANDLE
If rcdSet.RecordCount < 1 Then Exit Sub
'// set columnheaders
 Call InsertColumn(lvList, rcdSet)
'//set details
 Call FillListView(lvList, rcdSet, 3)
ERRORHANDLE:
      errorMsg Err, Me.Name, "load data"

End Sub

Private Sub Form_Unload(Cancel As Integer)
 Unload frmReference
 Set rcdSet = Nothing
 Set rsTemp = Nothing
End Sub

Private Sub ImgItemHelp_Click()
  myMsg "If the data does not show " _
  & "on textbox, click [Refresh] button!" & vbCrLf & vbCrLf _
  & "", "Items Help", 2, True

End Sub

Private Sub LvList_Click()
On Error GoTo ERRORHANDLE
If addRec = True Or editRec = True Then Exit Sub
Call BindDatasource(Me, rcdSet, lvList, True)
ERRORHANDLE:
'    If Err.Number = 91 Then
'       Exit Sub
'    Else
      errorMsg Err, Me.Name
'    End If
End Sub

Private Sub lvList_KeyUp(KeyCode As Integer, Shift As Integer)
 If KeyCode = 38 Or KeyCode = 40 Or KeyCode = 33 Or KeyCode = 34 Then LvList_Click
End Sub


Private Sub clearText()
Dim i As Integer
For i = txtEntry.LBound To txtEntry.UBound
        txtEntry(i).Text = ""
        Next i
End Sub

Private Sub txtEntry_GotFocus(Index As Integer)
nxTab = Index
txtEntry(nxTab).SelStart = 0
txtEntry(nxTab).SelLength = Len(txtEntry(nxTab).Text)
End Sub

Private Sub TxtEntry_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim lastTab As Integer
On Error GoTo ERRORHANDLE
lastTab = (rcdSet.Fields.Count - 1)
If KeyCode = 13 Then
     If nxTab = lastTab Then Exit Sub
     nxTab = nxTab + 1
ElseIf KeyCode = 38 Then  'up arrow key
     If nxTab = 0 Or nxTab = 1 Then Exit Sub
     nxTab = nxTab - 1
End If
txtEntry(nxTab).SetFocus
ERRORHANDLE:
 errorMsg Err, Me.Name
End Sub

Public Sub PopulateData(ByRef sListView As ListView, _
                           ByRef sRecSource As Recordset, _
                           ByVal sIcoNdx As Byte, _
                           Optional ByRef SN As Long = 0)
'//set details
    Dim X As ListItem
    Dim i As Byte
    Dim sFieldsNum As Integer
    On Error Resume Next
    '//initialize
    sFieldsNum = (sRecSource.Fields.Count - 1)  '(sRecSource.Fields.Count - 1)
With sRecSource
    .Requery    '//Use this method to make sure that a Recordset contains the most recent data
                  'first record becomes the current record
    .MoveLast
End With
    If sRecSource.RecordCount < 1 Then Exit Sub
         Set X = sListView.ListItems.Add(, , SN, sIcoNdx, sIcoNdx)
         For i = 1 To sFieldsNum  'sub items must start from 1 not zero / cause of error Invalid Property Value err.number 380
           With X.ListSubItems.Add(, , txtEntry(i).Text)
           End With
          Next i
End Sub


