VERSION 5.00
Begin VB.Form frmSPLASH 
   BorderStyle     =   0  'None
   Caption         =   "ReadMe.First"
   ClientHeight    =   4830
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   6765
   Icon            =   "frmSPLASH.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSPLASH.frx":B01A
   ScaleHeight     =   4830
   ScaleWidth      =   6765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdClose 
      BackColor       =   &H0024C5DB&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6240
      MouseIcon       =   "frmSPLASH.frx":DF28
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Close"
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton CmdAbout 
      BackColor       =   &H0024C5DB&
      Caption         =   "<:- About -:>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4680
      MouseIcon       =   "frmSPLASH.frx":E7F2
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3360
      Width           =   1815
   End
   Begin VB.CommandButton CmdGO 
      BackColor       =   &H0024C5DB&
      Caption         =   "<:- Continue -:>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4680
      MouseIcon       =   "frmSPLASH.frx":F0BC
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3960
      Width           =   1815
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(c)2008 all rights reserved"
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "edwinSoftware"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   345
      Left            =   1575
      TabIndex        =   9
      Top             =   3700
      Width           =   2040
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   630
      Left            =   3720
      TabIndex        =   8
      Top             =   2040
      Width           =   780
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H007ADBE9&
      BackStyle       =   0  'Transparent
      Caption         =   "<:-  Freeware  -:>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Left            =   4920
      TabIndex        =   5
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "version"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   510
      Left            =   2040
      TabIndex        =   4
      Top             =   2160
      Width           =   1545
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CHECK BOOK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   600
      Left            =   1560
      TabIndex        =   3
      Top             =   360
      Width           =   3525
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0920-6747545"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2420
      TabIndex        =   2
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "cyber_edu2005@yahoo.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1050
      TabIndex        =   1
      Top             =   3960
      Width           =   2595
   End
End
Attribute VB_Name = "frmSPLASH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdAbout_Click()
 FrmAbout.Show
End Sub

Private Sub CmdClose_Click()
      On Error Resume Next
        Dim Cancel As Boolean
        Dim msg As String
        msg = "This will close the application."
        msg = msg & vbCrLf & "Do you want to proceed?"
        If MsgBox(msg, vbYesNo + vbQuestion, "Check Book System") = vbNo Then
           Cancel = True
        Else
        UnloadAllForms
        myMsg "See yah later!!! " _
        & "Edwin delos Santos" & vbCrLf & vbCrLf _
        & "", "Bye Bye !", 2, True
        Unload Me
        End If
End Sub

Private Sub CmdGO_Click()
 frmCashDisburse.Show
End Sub



Private Sub Form_Load()
On Error GoTo errMsg
FormRndCorner Me, 450, 320
App.HelpFile = App.Path & "\checkbook.chm"
'[===============================================================]
'< Not the best way to check                                     >
'< Better to use the FindWindow API, when running an exe file    >
'[===============================================================]
If App.PrevInstance = True Then
  MsgBox "This application is already running.", vbInformation, "Warning!"
  End
End If
errMsg:
  errorMsg Err, Me.Name, "Form_Load"
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   down = True
    w = X
    t = Y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If down Then
        Top = Top + Y - t
        Left = Left + X - w
   End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 down = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Set frmSPLASH = Nothing
 UnloadAllForms
End Sub
