VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmPrint 
   BackColor       =   &H007ADBE9&
   Caption         =   "Print Summary"
   ClientHeight    =   4575
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7680
   HelpContextID   =   60
   Icon            =   "frmPrint.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   7680
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab SSTab1 
      Height          =   3975
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   7011
      _Version        =   393216
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   8051689
      TabCaption(0)   =   "Data Option"
      TabPicture(0)   =   "frmPrint.frx":0FA2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FmePrint"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "CmdOk"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "indexList"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Footer / Header"
      TabPicture(1)   =   "frmPrint.frx":0FBE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Picture3"
      Tab(1).Control(1)=   "Picture4"
      Tab(1).Control(2)=   "chkLongDate"
      Tab(1).Control(3)=   "Hline1"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Print Option"
      TabPicture(2)   =   "frmPrint.frx":0FDA
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame2"
      Tab(2).Control(1)=   "cmdPrint"
      Tab(2).Control(2)=   "Frame1"
      Tab(2).ControlCount=   3
      Begin VB.Frame Frame2 
         Caption         =   "Paper "
         ForeColor       =   &H00C00000&
         Height          =   2895
         Left            =   -72120
         TabIndex        =   22
         Top             =   720
         Width           =   2295
         Begin VB.OptionButton Option3 
            Caption         =   "Letter 8.5x11 in"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   24
            Top             =   840
            Value           =   -1  'True
            Width           =   1935
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Size"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   23
            Top             =   480
            Width           =   375
         End
      End
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H00E0E0E0&
         Height          =   975
         Left            =   -69360
         MouseIcon       =   "frmPrint.frx":0FF6
         MousePointer    =   99  'Custom
         Picture         =   "frmPrint.frx":18C0
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Print"
         Top             =   960
         Width           =   1095
      End
      Begin VB.Frame Frame1 
         Caption         =   "Orientation"
         ForeColor       =   &H00C00000&
         Height          =   2895
         Left            =   -74640
         TabIndex        =   18
         Top             =   720
         Width           =   2175
         Begin VB.OptionButton Option1 
            Caption         =   "Portrait"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   20
            Top             =   480
            Width           =   1695
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Land Scape"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   19
            Top             =   960
            Value           =   -1  'True
            Width           =   1695
         End
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   -74640
         ScaleHeight     =   1155
         ScaleWidth      =   4875
         TabIndex        =   13
         Top             =   720
         Width           =   4935
         Begin VB.TextBox txtHeader 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   2
            Left            =   0
            TabIndex        =   16
            Text            =   "REPORT DATE"
            Top             =   840
            Width           =   4935
         End
         Begin VB.TextBox txtHeader 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   1
            Left            =   0
            TabIndex        =   15
            Text            =   "TYPE OF REPORT"
            Top             =   480
            Width           =   4935
         End
         Begin VB.TextBox txtHeader 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   0
            Left            =   0
            TabIndex        =   14
            Text            =   "YOUR COMPANY"
            Top             =   120
            Width           =   4935
         End
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00FFFFFF&
         Height          =   855
         Left            =   -74640
         ScaleHeight     =   795
         ScaleWidth      =   4875
         TabIndex        =   10
         Top             =   2400
         Width           =   4935
         Begin VB.TextBox txtFooter 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   1
            Left            =   0
            TabIndex        =   12
            Text            =   "YOUR JOB TITLE"
            Top             =   480
            Width           =   4935
         End
         Begin VB.TextBox txtFooter 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   0
            Left            =   0
            TabIndex        =   11
            Text            =   "YOUR NAME / Signatory"
            Top             =   120
            Width           =   4935
         End
      End
      Begin VB.CheckBox chkLongDate 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Long Date"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -69360
         TabIndex        =   9
         Top             =   1680
         Width           =   1095
      End
      Begin VB.ListBox indexList 
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         ForeColor       =   &H00FFFFFF&
         Height          =   1980
         ItemData        =   "frmPrint.frx":240E
         Left            =   0
         List            =   "frmPrint.frx":2430
         TabIndex        =   8
         Top             =   1920
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton CmdOk 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5640
         Picture         =   "frmPrint.frx":2455
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   3480
         Width           =   1215
      End
      Begin VB.Frame FmePrint 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Select ( Print only what you  need ! )"
         ForeColor       =   &H00C00000&
         Height          =   2895
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   6975
         Begin VB.CommandButton CmdMoveBack 
            Caption         =   "<"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3120
            TabIndex        =   5
            Top             =   1560
            Width           =   615
         End
         Begin VB.CommandButton CmdMove 
            Caption         =   ">"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3120
            TabIndex        =   4
            Top             =   1080
            Width           =   615
         End
         Begin VB.ListBox List2Print 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   2370
            ItemData        =   "frmPrint.frx":2FA3
            Left            =   120
            List            =   "frmPrint.frx":2FAA
            Sorted          =   -1  'True
            TabIndex        =   3
            Top             =   360
            Width           =   2895
         End
         Begin VB.ListBox ListPrint 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   2280
            ItemData        =   "frmPrint.frx":2FBD
            Left            =   3840
            List            =   "frmPrint.frx":2FBF
            Style           =   1  'Checkbox
            TabIndex        =   2
            Top             =   360
            Width           =   2895
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Validate list,  Click check button. "
            ForeColor       =   &H000000C0&
            Height          =   195
            Left            =   3840
            TabIndex        =   6
            Top             =   0
            Width           =   2355
         End
      End
      Begin VB.PictureBox Hline1 
         Height          =   30
         Left            =   -74640
         ScaleHeight     =   30
         ScaleWidth      =   6495
         TabIndex        =   17
         Top             =   2160
         Width           =   6495
      End
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'[---------------------------------]
'< Print Report                    >
'<(Print only what you need to...  >
'< designed for Global purpose     >
'< ... can be called by any table  >
'< coded by:edwin delos santos     >
'[---------------------------------]
                
Option Explicit

Public pPrintForm As Form  'source Form
Public pPrintTABLE As String
Public pPrintCon As New ADODB.Connection
Public pPrintRecset As New ADODB.Recordset

Private prevDate As String 'store prev date format (use by txtheader(2) - see tab)


Private Sub chkLongDate_Click()
Dim chk As Integer
chk = chkLongDate.Value
If chk = 1 Then
   If IsDate(txtHeader(2).Text) Then
      prevDate = txtHeader(2).Text
      Dim xd As String
       xd = Format(txtHeader(2).Text, "Long Date")
        txtHeader(2).Text = xd
   Else
      chkLongDate.Value = 0
   End If
ElseIf chk = 0 Then
      txtHeader(2).Text = prevDate
End If
End Sub

Private Sub CmdMove_Click()
' Move one item from left to right.
    If List2Print.ListIndex >= 0 Then
        ListPrint.AddItem List2Print.Text
        List2Print.RemoveItem List2Print.ListIndex
        initPrint = False
    End If
End Sub

Private Sub CmdMoveBack_Click()
' Move one item from left to right.
Dim idx As Integer
idx = ListPrint.ListIndex
    If ListPrint.ListIndex >= 0 Then
        If ListPrint.Selected(idx) = True Then Exit Sub
        List2Print.AddItem ListPrint.Text
        ListPrint.RemoveItem ListPrint.ListIndex
        initPrint = False
    End If
End Sub

Private Sub cmdOK_Click()
  '//validate
  initPrint = print_Init(ListPrint)
End Sub

Private Sub CmdPrint_Click()
     If ListPrint.ListCount = 0 Then
       myMsg "No data to print !", "Select data ...", 2, True
       Exit Sub
     End If
     If initPrint = False Then
       myMsg "Please Validate First !", "Cash Transaction (Print)", 2, True
       Exit Sub
     End If
     If Option1.Value = True Then
        PrintReport pPrintRecset, 1
     ElseIf Option2.Value = True Then
        PrintReport pPrintRecset, 2
     End If
End Sub

'// PRINTER REPORT PROCEDURE
'//CODED BY EDWIN DELOS SANTOS
Private Sub Headers()
 Dim dat As String
 
 dat = Format(Now(), "long date")
 Printer.Print Tab(6); dat
If Option1.Value = True Then
 Call prnCenterText(txtHeader(0).Text, 140) 'company
 Call prnCenterText(txtHeader(1).Text, 140) 'type of report
 Call prnCenterText(txtHeader(2).Text, 140) ' date
End If

If Option2.Value = True Then
 Call prnCenterText(txtHeader(0).Text, 180)
 Call prnCenterText(txtHeader(1).Text, 180)
 Call prnCenterText(txtHeader(2).Text, 180)
End If
End Sub
Private Sub PrintReport(ByRef srcRS As Recordset, Optional ByVal mOrient As Long = 2)
'// coded by edwin delos santos
'// fixed settings has been made to this report
'// you may adjust settings using variables ...
'// like default:  tab, page orientation, number of pages to print, quality and so on...
'// no modify except for the above said defaul settings ...
On Error GoTo printErr
Dim curr_Rec As Long  'current record looping becomes 0 if max line per page is reaced
Dim rec_Counter As Long 'record  counter
Dim recPerPage As Long
Dim ans 'answer
Dim strFont As String, sngSize As Single
'//initialize
rec_Counter = 0
curr_Rec = 0

If mOrient = 2 Then
   recPerPage = 25
End If
If mOrient = 1 Then
   recPerPage = 35
End If
'//
ans = MsgBox("Proceed?", vbYesNo + vbQuestion, "Print Summary")
  If ans = vbYes Then
'//save current printer settings
         strFont = Printer.Font
         sngSize = Printer.FontSize
         Printer.Orientation = mOrient    '2   'Landscape
         Printer.Font = "ms sans serif"
         Printer.FontSize = 9
         Printer.Print
'// headers
         Headers             '< --------------------  headers ----------------->
         Printer.Print
         Printer.Print
       With srcRS
            .MoveFirst
              Printer.Font.Underline = True
              Call Print_Headings(srcRS, 6, ListPrint, 12)
              Printer.Font.Underline = False
         While Not .EOF = True
'//print details
               Call Print_Details(srcRS, 6, ListPrint)  'see procedure
               Printer.Print                            'line space
'//if not eof = true
                .MoveNext
                rec_Counter = rec_Counter + 1          'counter
                curr_Rec = curr_Rec + 1                 'store record printed
             If curr_Rec = recPerPage Then    '25 Then                      'proceed to next page
                Printer.Print Tab(6); "<< next page >>"
                curr_Rec = 0                           'reset current record to 0
'//next page
               If rec_Counter < .RecordCount Then
                   Printer.NewPage
                   Printer.Print
                   Printer.Print
                   Headers   'next page headers
                   Printer.Print
                   Printer.Font.Underline = True
                   Call Print_Headings(srcRS, 6, ListPrint, 12) '2nd page headings
                   Printer.Font.Underline = False
                   Printer.Print
                End If 'rec_count
            End If   'curr_Rec
        Wend  'while not EOF
'//print line
        If mOrient = 1 Then
          Printer.Print Tab(6); "___________________________________________________________________________________________________"
        End If
        If mOrient = 2 Then
          Printer.Print Tab(6); "___________________________________________________________________________________________________________________________________________________"
        End If
'// print Total
          Printer.Print Tab(6); "T O T A L >>";
          Call Print_Total(srcRS, 6, ListPrint)
          Printer.Print
'//Footer
          Footers             '< --------------------  footers ----------------->
          Printer.Print
      End With
'//send information to the printer
         Printer.EndDoc
'//reset printer setting
         Printer.Font = strFont
         Printer.FontSize = sngSize
         If srcRS.EOF = True Then   'end of file
              MsgBox "D O N E !!", vbInformation, "Printing..."
         End If
   Else
       Exit Sub
   End If  'answer
printErr:
'   errorMsg err, Me.name

End Sub

Private Sub Footers()
     Printer.Print Tab(6); txtFooter(0).Text
     Printer.Print Tab(6); txtFooter(1).Text
     Printer.FontSize = 5
     Printer.Print Tab(10); "eS(c)2008"
     Printer.Print
End Sub



Private Sub Form_Activate()
'Gradient Me, 100, 150, 255, 1
End Sub

Private Sub Form_Load()
   ' FormRndCorner Me, 640, 260
   'Gradient Me, 100, 150, 255, 1
    '//initialize
    SSTab1.Tab = 0
    txtHeader(2).Text = Format(Now(), "mm/dd/yyyy")
    Show
    List2Print.SetFocus
    ListPrint.ListIndex = 0
    Call Insert_Fields(List2Print, pPrintRecset)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If MsgBox("Exit this system?", vbYesNo + vbQuestion, "Print Summary") = vbNo Then
  Cancel = 1 'true
Else
  Unload Me
End If

End Sub

Private Sub Form_Resize()
With Me
  If .WindowState = 0 Then
   .Height = 5085
   .Width = 7800
  End If
End With
'Gradient Me, 100, 150, 255, 1
   SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set pPrintForm = Nothing
Set pPrintCon = Nothing
Set pPrintRecset = Nothing
Set frmPrint = Nothing
Unload frmPrint
End Sub

Private Sub List2Print_DblClick()
   CmdMove_Click
End Sub

Private Sub Option3_Click()
  Exit Sub
End Sub

