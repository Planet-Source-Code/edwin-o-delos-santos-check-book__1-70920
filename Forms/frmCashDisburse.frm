VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmCashDisburse 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Bank Account"
   ClientHeight    =   9975
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13575
   HelpContextID   =   90
   Icon            =   "frmCashDisburse.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmCashDisburse.frx":1E26
   ScaleHeight     =   9975
   ScaleWidth      =   13575
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImgMenu 
      Left            =   120
      Top             =   8520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   49
      ImageHeight     =   49
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashDisburse.frx":5904
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashDisburse.frx":773A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashDisburse.frx":9570
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashDisburse.frx":B3A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashDisburse.frx":D1DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCashDisburse.frx":F012
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox PicNav 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   11040
      ScaleHeight     =   255
      ScaleWidth      =   2415
      TabIndex        =   62
      Top             =   8760
      Width           =   2415
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
         Picture         =   "frmCashDisburse.frx":10E48
         Style           =   1  'Graphical
         TabIndex        =   65
         ToolTipText     =   "Next"
         Top             =   0
         Width           =   375
      End
      Begin VB.TextBox TxtEntry 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E4F0B7&
         BorderStyle     =   0  'None
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   0
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   80
         Top             =   0
         Width           =   975
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
         Picture         =   "frmCashDisburse.frx":110FD
         Style           =   1  'Graphical
         TabIndex        =   66
         ToolTipText     =   "Last"
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
         Picture         =   "frmCashDisburse.frx":113B2
         Style           =   1  'Graphical
         TabIndex        =   64
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
         Picture         =   "frmCashDisburse.frx":11667
         Style           =   1  'Graphical
         TabIndex        =   63
         ToolTipText     =   "First"
         Top             =   0
         Width           =   375
      End
   End
   Begin MSComctlLib.ListView LvMenu 
      Height          =   1095
      Left            =   120
      TabIndex        =   61
      Top             =   8505
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   1931
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      OLEDragMode     =   1
      PictureAlignment=   5
      _Version        =   393217
      Icons           =   "ImgMenu"
      SmallIcons      =   "ImgMenu"
      ForeColor       =   12582912
      BackColor       =   16777215
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDragMode     =   1
      NumItems        =   0
   End
   Begin VB.PictureBox PicAccount 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   7920
      ScaleHeight     =   645
      ScaleWidth      =   2985
      TabIndex        =   47
      Top             =   0
      Visible         =   0   'False
      Width           =   3015
      Begin VB.CommandButton CmdRef 
         BackColor       =   &H00DCCD78&
         Caption         =   "Refresh"
         Height          =   275
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   0
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00DCCD78&
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   275
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   0
         Width           =   255
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   3135
         Left            =   0
         TabIndex        =   50
         Top             =   240
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   5530
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "i16x16"
         SmallIcons      =   "i16x16"
         ForeColor       =   12582912
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "<:- Enter to Select !  -:>"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   51
         Top             =   0
         Width           =   1635
      End
   End
   Begin VB.PictureBox PicVendor 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   4800
      ScaleHeight     =   645
      ScaleWidth      =   2745
      TabIndex        =   42
      Top             =   0
      Visible         =   0   'False
      Width           =   2775
      Begin MSComctlLib.ListView ListView1 
         Height          =   3015
         Left            =   0
         TabIndex        =   44
         Top             =   240
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   5318
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "i16x16"
         SmallIcons      =   "i16x16"
         ForeColor       =   12582912
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin VB.CommandButton CmdRefresh 
         BackColor       =   &H00DCCD78&
         Caption         =   "Refresh"
         Height          =   275
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   0
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00DCCD78&
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   275
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "<:- Enter to Select !  -:>"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   45
         Top             =   0
         Width           =   1635
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   9600
      Width           =   13575
      _ExtentX        =   23945
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   18283
            Text            =   "<:- system designed by:  edwinSoftware  (c)-2008  all rights reserved -:>"
            TextSave        =   "<:- system designed by:  edwinSoftware  (c)-2008  all rights reserved -:>"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "8/3/2008"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "12:49 AM"
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7935
      HelpContextID   =   90
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   13575
      _ExtentX        =   23945
      _ExtentY        =   13996
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   6
      TabHeight       =   794
      BackColor       =   16777215
      ForeColor       =   12582912
      TabCaption(0)   =   "Expenses"
      TabPicture(0)   =   "frmCashDisburse.frx":1191C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "PicDetails"
      Tab(0).Control(1)=   "cmdButton(1)"
      Tab(0).Control(2)=   "List1"
      Tab(0).Control(3)=   "i16x16"
      Tab(0).Control(4)=   "Picture1"
      Tab(0).Control(5)=   "cmdButton(6)"
      Tab(0).Control(6)=   "cmdButton(5)"
      Tab(0).Control(7)=   "cmdButton(4)"
      Tab(0).Control(8)=   "cmdButton(3)"
      Tab(0).Control(9)=   "cmdButton(2)"
      Tab(0).Control(10)=   "cmdButton(0)"
      Tab(0).Control(11)=   "ChkPrint"
      Tab(0).Control(12)=   "lvList"
      Tab(0).Control(13)=   "Shape1"
      Tab(0).Control(14)=   "lblFLDi(9)"
      Tab(0).ControlCount=   15
      TabCaption(1)   =   "Items"
      TabPicture(1)   =   "frmCashDisburse.frx":11938
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "lblSelectedAmt"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lblSelected"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "ImgItemHelp"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "LblFLD(0)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "LblFLD(2)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "LblFLD(1)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "LblFLD(3)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "LblFLD(4)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "LblFLD(5)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "LblFLD(6)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "LblFLD(7)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "LblFLD(8)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "LblFLD(10)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "TbMenu"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "lvItems"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "ImgToolbar"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "List2"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "CmdFilter"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "TextEntry(5)"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "TextEntry(4)"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "TextEntry(3)"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "TextEntry(2)"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "TextEntry(1)"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "TextEntry(0)"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "Hline1"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "TextEntry(6)"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "TextEntry(7)"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "TextEntry(8)"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "TextEntry(9)"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "List3"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).ControlCount=   30
      Begin VB.ListBox List3 
         Height          =   1230
         Left            =   7440
         TabIndex        =   93
         Top             =   2400
         Width           =   2055
      End
      Begin VB.TextBox TextEntry 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1440
         Index           =   9
         Left            =   9720
         TabIndex        =   89
         Top             =   1080
         Width           =   3375
      End
      Begin VB.TextBox TextEntry 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   8
         Left            =   7080
         TabIndex        =   86
         Top             =   1800
         Width           =   2175
      End
      Begin VB.TextBox TextEntry 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   7
         Left            =   7080
         TabIndex        =   84
         Top             =   1440
         Width           =   2175
      End
      Begin VB.TextBox TextEntry 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   6
         Left            =   7080
         TabIndex        =   82
         Top             =   1080
         Width           =   2175
      End
      Begin CHECKBOOK.Hline Hline1 
         Height          =   30
         Left            =   240
         TabIndex        =   79
         Top             =   2760
         Width           =   13095
         _ExtentX        =   23098
         _ExtentY        =   53
      End
      Begin VB.TextBox TextEntry 
         Appearance      =   0  'Flat
         BackColor       =   &H00E8FBF9&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   72
         Top             =   720
         Width           =   3375
      End
      Begin VB.TextBox TextEntry 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   1920
         TabIndex        =   71
         Top             =   1080
         Width           =   3375
      End
      Begin VB.TextBox TextEntry 
         Appearance      =   0  'Flat
         BackColor       =   &H00E8FBF9&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   70
         Top             =   1440
         Width           =   3375
      End
      Begin VB.TextBox TextEntry 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   1920
         TabIndex        =   69
         Top             =   1800
         Width           =   3375
      End
      Begin VB.TextBox TextEntry 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   1920
         TabIndex        =   68
         Top             =   2160
         Width           =   3375
      End
      Begin VB.TextBox TextEntry 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   5
         Left            =   7080
         TabIndex        =   67
         Top             =   720
         Width           =   2175
      End
      Begin VB.PictureBox PicDetails 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3735
         Left            =   -64680
         MousePointer    =   99  'Custom
         Picture         =   "frmCashDisburse.frx":11954
         ScaleHeight     =   3735
         ScaleWidth      =   3135
         TabIndex        =   52
         Top             =   600
         Width           =   3135
         Begin VB.TextBox TxtEntry 
            Appearance      =   0  'Flat
            BackColor       =   &H00E4F0B7&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   5
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   55
            Top             =   360
            Width           =   1575
         End
         Begin VB.TextBox TxtEntry 
            Appearance      =   0  'Flat
            BackColor       =   &H00E4F0B7&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   3
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   54
            Top             =   1200
            Width           =   3015
         End
         Begin VB.TextBox TxtEntry 
            Appearance      =   0  'Flat
            BackColor       =   &H00E4F0B7&
            BorderStyle     =   0  'None
            Height          =   1755
            Index           =   8
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   53
            Top             =   1800
            Width           =   3015
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "(Payment)"
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
            Left            =   120
            TabIndex        =   81
            Top             =   720
            Width           =   915
         End
         Begin VB.Label BtnDown 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "[F2]"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   210
            Index           =   1
            Left            =   2760
            TabIndex        =   60
            Top             =   960
            Visible         =   0   'False
            Width           =   300
         End
         Begin VB.Label lblFLDi 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "lblFLDi"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   240
            Index           =   0
            Left            =   1920
            TabIndex        =   59
            Top             =   0
            Visible         =   0   'False
            Width           =   630
         End
         Begin VB.Label lblFLDi 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "lblFLDi"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   240
            Index           =   8
            Left            =   120
            TabIndex        =   58
            Top             =   1560
            Width           =   630
         End
         Begin VB.Label lblFLDi 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "lblFLDi"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   240
            Index           =   5
            Left            =   120
            TabIndex        =   57
            Top             =   120
            Width           =   630
         End
         Begin VB.Label lblFLDi 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "lblFLDi"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   240
            Index           =   3
            Left            =   120
            TabIndex        =   56
            Top             =   960
            Width           =   630
         End
      End
      Begin VB.CommandButton CmdFilter 
         BackColor       =   &H007ADBE9&
         Caption         =   "<< Filter >>"
         Height          =   315
         Left            =   12360
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   3480
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdButton 
         BackColor       =   &H00DCCD78&
         Caption         =   "&Save"
         Height          =   325
         Index           =   1
         Left            =   -72360
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   4440
         Width           =   900
      End
      Begin VB.ListBox List2 
         BackColor       =   &H0000FFFF&
         Height          =   1815
         ItemData        =   "frmCashDisburse.frx":154E8
         Left            =   0
         List            =   "frmCashDisburse.frx":154F2
         TabIndex        =   35
         Top             =   4800
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.ListBox List1 
         BackColor       =   &H0000FFFF&
         Height          =   1815
         ItemData        =   "frmCashDisburse.frx":154FC
         Left            =   -75000
         List            =   "frmCashDisburse.frx":1550C
         TabIndex        =   34
         Top             =   4920
         Visible         =   0   'False
         Width           =   255
      End
      Begin MSComctlLib.ImageList ImgToolbar 
         Left            =   6360
         Top             =   3120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   24
         ImageHeight     =   24
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   8
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashDisburse.frx":1551C
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashDisburse.frx":15C96
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashDisburse.frx":16410
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashDisburse.frx":16B8A
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashDisburse.frx":17304
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashDisburse.frx":17A7E
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashDisburse.frx":181F8
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashDisburse.frx":18972
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList i16x16 
         Left            =   -74400
         Top             =   5040
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   23
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashDisburse.frx":190EC
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashDisburse.frx":19686
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashDisburse.frx":19A20
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashDisburse.frx":19DBA
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashDisburse.frx":1A154
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashDisburse.frx":1A4EE
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashDisburse.frx":1A888
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashDisburse.frx":1AC22
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashDisburse.frx":1B634
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashDisburse.frx":1B688
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashDisburse.frx":1BA22
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashDisburse.frx":1BDBC
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashDisburse.frx":1C156
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashDisburse.frx":1C4F0
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashDisburse.frx":1CF02
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashDisburse.frx":1D914
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashDisburse.frx":1E326
               Key             =   ""
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashDisburse.frx":1ED38
               Key             =   ""
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashDisburse.frx":1F74A
               Key             =   ""
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashDisburse.frx":2015C
               Key             =   ""
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashDisburse.frx":20B6E
               Key             =   ""
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashDisburse.frx":2110A
               Key             =   ""
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashDisburse.frx":216A6
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3735
         Left            =   -75000
         Picture         =   "frmCashDisburse.frx":21C40
         ScaleHeight     =   3735
         ScaleWidth      =   10335
         TabIndex        =   13
         Top             =   600
         Width           =   10335
         Begin VB.TextBox TxtEntry 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0FF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   9
            Left            =   9840
            TabIndex        =   91
            Text            =   "Y"
            Top             =   3240
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox TxtEntry 
            BackColor       =   &H00E4F0B7&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   7920
            MaxLength       =   20
            TabIndex        =   19
            Top             =   240
            Width           =   2175
         End
         Begin VB.TextBox TxtEntry 
            Appearance      =   0  'Flat
            BackColor       =   &H00E4F0B7&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   7920
            TabIndex        =   18
            ToolTipText     =   "Format: 01/01/2008"
            Top             =   720
            Width           =   1815
         End
         Begin VB.TextBox TxtEntry 
            Appearance      =   0  'Flat
            BackColor       =   &H00E4F0B7&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   4
            Left            =   7920
            TabIndex        =   17
            Top             =   1320
            Width           =   2175
         End
         Begin VB.TextBox TxtEntry 
            Appearance      =   0  'Flat
            BackColor       =   &H00E4F0B7&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   6
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   16
            Top             =   1320
            Width           =   4335
         End
         Begin VB.TextBox TxtEntry 
            Appearance      =   0  'Flat
            BackColor       =   &H00E4F0B7&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1005
            Index           =   7
            Left            =   1440
            MultiLine       =   -1  'True
            TabIndex        =   15
            Top             =   2160
            Width           =   2655
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   285
            Left            =   9840
            TabIndex        =   14
            Top             =   720
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            _Version        =   393216
            Format          =   20774913
            CurrentDate     =   39650
         End
         Begin VB.Label BtnDown 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "[F2]"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   210
            Index           =   0
            Left            =   480
            TabIndex        =   39
            Top             =   360
            Visible         =   0   'False
            Width           =   300
         End
         Begin VB.Line Line1 
            X1              =   7920
            X2              =   10080
            Y1              =   600
            Y2              =   600
         End
         Begin VB.Line Line2 
            X1              =   7920
            X2              =   10080
            Y1              =   1080
            Y2              =   1080
         End
         Begin VB.Line Line3 
            X1              =   7920
            X2              =   10080
            Y1              =   1680
            Y2              =   1680
         End
         Begin VB.Line Line4 
            X1              =   2040
            X2              =   6360
            Y1              =   1680
            Y2              =   1680
         End
         Begin VB.Label lblAmtInWord 
            AutoSize        =   -1  'True
            BackColor       =   &H000000FF&
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
            Left            =   360
            TabIndex        =   32
            Top             =   1800
            Width           =   2445
         End
         Begin VB.Line Line5 
            X1              =   240
            X2              =   9240
            Y1              =   2040
            Y2              =   2040
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pesos"
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
            Left            =   9480
            TabIndex        =   31
            Top             =   1800
            Width           =   585
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Php"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   7320
            TabIndex        =   30
            Top             =   1320
            Width           =   360
         End
         Begin VB.Label lblFLDi 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "lblFLDi"
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
            Index           =   1
            Left            =   7080
            TabIndex        =   29
            Top             =   360
            Width           =   630
         End
         Begin VB.Label lblFLDi 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "lblFLDi"
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
            Index           =   2
            Left            =   7080
            TabIndex        =   28
            Top             =   720
            Width           =   630
         End
         Begin VB.Label lblFLDi 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "lblFLDi"
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
            Index           =   4
            Left            =   7200
            TabIndex        =   27
            Top             =   1080
            Visible         =   0   'False
            Width           =   630
         End
         Begin VB.Label lblFLDi 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "lblFLDi"
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
            Index           =   6
            Left            =   240
            TabIndex        =   26
            Top             =   1560
            Visible         =   0   'False
            Width           =   630
         End
         Begin VB.Label lblFLDi 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "lblFLDi"
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
            Index           =   7
            Left            =   240
            TabIndex        =   25
            Top             =   2520
            Width           =   630
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0920-6747-545"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   300
            Left            =   1440
            TabIndex        =   24
            Top             =   3360
            Width           =   1920
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "edwinSoftware"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   3375
            TabIndex        =   23
            Top             =   120
            Width           =   1440
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Makati City, Philippines"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   3330
            TabIndex        =   22
            Top             =   360
            Width           =   1635
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Pay to the &Order of"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   240
            TabIndex        =   21
            Top             =   1320
            Width           =   1650
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "cyber_edu2005@yahoo.com"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   3090
            TabIndex        =   20
            Top             =   600
            Width           =   2115
         End
      End
      Begin VB.CommandButton cmdButton 
         BackColor       =   &H00DCCD78&
         Caption         =   "&Refresh"
         Height          =   325
         Index           =   6
         Left            =   -69480
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   4440
         Width           =   900
      End
      Begin VB.CommandButton cmdButton 
         BackColor       =   &H00DCCD78&
         Caption         =   "&Delete"
         Height          =   325
         Index           =   5
         Left            =   -70440
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   4440
         Width           =   900
      End
      Begin VB.CommandButton cmdButton 
         BackColor       =   &H00DCCD78&
         Caption         =   "&Cancel"
         Height          =   325
         Index           =   4
         Left            =   -71400
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   4440
         Width           =   900
      End
      Begin VB.CommandButton cmdButton 
         BackColor       =   &H00DCCD78&
         Caption         =   "&Update"
         Height          =   325
         Index           =   3
         Left            =   -72360
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   4440
         Width           =   900
      End
      Begin VB.CommandButton cmdButton 
         BackColor       =   &H00DCCD78&
         Caption         =   "&Edit"
         Height          =   325
         Index           =   2
         Left            =   -73320
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   4440
         Width           =   900
      End
      Begin VB.CommandButton cmdButton 
         BackColor       =   &H00DCCD78&
         Caption         =   "&Add"
         Height          =   325
         Index           =   0
         Left            =   -74280
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   4440
         Width           =   900
      End
      Begin VB.CheckBox ChkPrint 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   -66240
         TabIndex        =   1
         Top             =   4560
         Value           =   1  'Checked
         Width           =   225
      End
      Begin MSComctlLib.ListView lvItems 
         Height          =   4035
         Left            =   0
         TabIndex        =   4
         Top             =   3840
         Width           =   13545
         _ExtentX        =   23892
         _ExtentY        =   7117
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         PictureAlignment=   5
         _Version        =   393217
         Icons           =   "i16x16"
         SmallIcons      =   "i16x16"
         ForeColor       =   -2147483640
         BackColor       =   15135229
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
         Picture         =   "frmCashDisburse.frx":11DE80
      End
      Begin MSComctlLib.ListView lvList 
         Height          =   2955
         Left            =   -75000
         TabIndex        =   5
         Top             =   4920
         Width           =   13545
         _ExtentX        =   23892
         _ExtentY        =   5212
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         PictureAlignment=   5
         _Version        =   393217
         Icons           =   "i16x16"
         SmallIcons      =   "i16x16"
         ForeColor       =   -2147483640
         BackColor       =   16579829
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
         Picture         =   "frmCashDisburse.frx":1241D8
      End
      Begin MSComctlLib.Toolbar TbMenu 
         Height          =   660
         Left            =   240
         TabIndex        =   33
         Top             =   3000
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   1164
         ButtonWidth     =   1191
         ButtonHeight    =   1164
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "ImgToolbar"
         DisabledImageList=   "ImgToolbar"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   9
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Add"
               Key             =   "new"
               Object.ToolTipText     =   "Add"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Save"
               Key             =   "save"
               Object.ToolTipText     =   "Save"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Edit"
               Key             =   "edit"
               Object.ToolTipText     =   "Edit"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Update"
               Key             =   "update"
               Object.ToolTipText     =   "Update"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Cancel"
               Key             =   "cancel"
               Object.ToolTipText     =   "Cancel"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Delete"
               Key             =   "delete"
               Object.ToolTipText     =   "Delete"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Refresh"
               Key             =   "refresh"
               Object.ToolTipText     =   "Refresh"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Print"
               Key             =   "print"
               Object.ToolTipText     =   "Print Check Voucher"
               ImageIndex      =   8
            EndProperty
         EndProperty
         Begin VB.ListBox ListPrint 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   705
            ItemData        =   "frmCashDisburse.frx":127D7C
            Left            =   5400
            List            =   "frmCashDisburse.frx":127D7E
            Style           =   1  'Checkbox
            TabIndex        =   41
            Top             =   0
            Visible         =   0   'False
            Width           =   1215
         End
      End
      Begin VB.Label LblFLD 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000C&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Memo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Index           =   10
         Left            =   9720
         TabIndex        =   90
         Top             =   720
         Width           =   3375
      End
      Begin VB.Label LblFLD 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000C&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Amount:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Index           =   8
         Left            =   5640
         TabIndex        =   87
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label LblFLD 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000C&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Quantity:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Index           =   7
         Left            =   5640
         TabIndex        =   85
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label LblFLD 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000C&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Price:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Index           =   6
         Left            =   5640
         TabIndex        =   83
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label LblFLD 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000C&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Unit:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Index           =   5
         Left            =   5640
         TabIndex        =   75
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label LblFLD 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000C&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Description:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Index           =   4
         Left            =   480
         TabIndex        =   78
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label LblFLD 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000C&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Item:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Index           =   3
         Left            =   480
         TabIndex        =   77
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label LblFLD 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000C&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Date:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Index           =   1
         Left            =   480
         TabIndex        =   76
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label LblFLD 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000C&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Check Number:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Index           =   2
         Left            =   480
         TabIndex        =   74
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label LblFLD 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000C&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SN:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Index           =   0
         Left            =   480
         TabIndex        =   73
         Top             =   720
         Width           =   1455
      End
      Begin VB.Image ImgItemHelp 
         Height          =   360
         Left            =   13080
         MouseIcon       =   "frmCashDisburse.frx":127D80
         MousePointer    =   99  'Custom
         Picture         =   "frmCashDisburse.frx":12864A
         Top             =   0
         Width           =   360
      End
      Begin VB.Label lblSelected 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Selected Check Number:"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   11520
         TabIndex        =   38
         Top             =   2880
         Width           =   1785
      End
      Begin VB.Label lblSelectedAmt 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount:"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   12720
         TabIndex        =   37
         Top             =   3120
         Width           =   585
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         Height          =   450
         Left            =   -74760
         Shape           =   4  'Rounded Rectangle
         Top             =   4395
         Width           =   6735
      End
      Begin VB.Label lblFLDi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "lblFLDi"
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
         Index           =   9
         Left            =   -66000
         TabIndex        =   2
         Top             =   4560
         Width           =   630
      End
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "<:-Expenses Navigation Button-:>"
      Height          =   195
      Left            =   11040
      TabIndex        =   92
      Top             =   8520
      Width           =   2370
   End
   Begin VB.Label LblFLD 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000C&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Amount:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Index           =   9
      Left            =   9240
      TabIndex        =   88
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Check Transaction"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   240
      TabIndex        =   40
      Top             =   0
      Width           =   1680
   End
End
Attribute VB_Name = "frmCashDisburse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsCash As Recordset
Private rsVend As Recordset
Private rsItem As Recordset
Private rsAcct As Recordset
Dim totalBy As clsTotalBy
Dim rsTemp As Recordset
Dim convert As NumToWord
Dim MyMenu As String  'lvmenu temporary storage


Private Sub CmdRef_Click()
Dim strSQL As String
If rsAcct.State = adStateOpen Then
    rsAcct.Close
End If
strSQL = "SELECT Account,Description "
strSQL = strSQL & "From Expenses order by Description"
rsAcct.Open strSQL, cnRef, adOpenStatic, adLockOptimistic
Load_Expenses
autoAlignCol ListView2
End Sub

Private Sub cmdRefresh_Click()
Dim SQL As String
If rsVend.State = adStateOpen Then
    rsVend.Close
End If
SQL = "SELECT Vendor_Name,Address,Vendor_ID "
SQL = SQL & "From Vendor order by Vendor_Name"
rsVend.Open SQL, cnRef, adOpenStatic, adLockOptimistic
Load_Vendor
autoAlignCol ListView1
End Sub




Private Sub ChkPrint_Click()
Dim ckPrn As Integer
ckPrn = ChkPrint.Value
If ckPrn = 1 Then
 TxtEntry(9).Text = "Y"
Else
 TxtEntry(9).Text = "n"
End If
End Sub



Private Sub nextNumber()
      Dim nextNo As Long
      Set rsTemp = New ADODB.Recordset
      Set rsTemp = rsCash
       If rsTemp.State = adStateOpen Then
          rsTemp.Close
       End If
      rsTemp.Open "SELECT * From CheckTrans order by SN", cnBank
      nextNo = Last_Recc(rsTemp)
      If nextNo > 0 Then
       TxtEntry(0).Text = nextNo
       TxtEntry(1).SetFocus
      Else
       nextNo = nextNo = 1
       TxtEntry(0).Text = nextNo
       TxtEntry(1).SetFocus
      End If
      Set rsTemp = Nothing
End Sub

Private Sub nextItemNumber()
      Dim nextNo As Long
      Set rsTemp = New ADODB.Recordset
       Set rsTemp = rsItem
        If rsTemp.State = adStateOpen Then
           rsTemp.Close
        End If
      rsTemp.Open "SELECT * From CheckItems order by SN", cnBank
      nextNo = Last_Recc(rsTemp)
      If nextNo > 0 Then
       TextEntry(0).Text = nextNo
       TextEntry(1).SetFocus
      Else
       nextNo = nextNo = 1
       TextEntry(0).Text = nextNo
       TextEntry(1).SetFocus
      End If
      Set rsTemp = Nothing
End Sub

Private Sub cmdButton_Click(Index As Integer)
'//                  A S E U C D R
'On Error GoTo ERRORHANDLE

Dim nextNo As Long
Select Case Index
   Case BtnAdd                       '<------ add new record ------->'
     clearTxt
     addRec = True
     TxtEntry(9).Text = "Y"
     cmdButtonShow ("0102100"), Me
     nextNumber
   Case BtnSave                       '<------ save new record ------>'
        cmdButtonShow ("1012011"), Me
        Call WriteData(Me, rsCash, True)
        Dim lngSN As Long
        lngSN = CLng(TxtEntry(0).Text)
        Call PopulateData(lvList, rsCash, 2, lngSN)
        addRec = False
        lvList.SetFocus
   Case BtnEdit                       '<------ edit record ---------->'
       '//VerifyEdit
     Dim src As String  'source
     Dim trg As String  'targer match
     src = TxtEntry(0).Text
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
        TxtEntry(1).SetFocus
   Case BtnUpdate                     '<------ update record -------->'
        cmdButtonShow ("1012011"), Me
        Call WriteData(Me, rsCash, False)
        LvwReplaceData Me, rsCash, lvList
        editRec = False
        lvList.SetFocus
   Case BtnCancel                     '<------ cancel update -------->'
        cmdButtonShow ("1012011"), Me
        addRec = False
        editRec = False
        lvList.SetFocus
   Case BtnDelete                     '<------ delete record -------->'
        '// no delete here please !
        Call Delete_Record(rsCash, lvList)
        lvList.SetFocus
   Case BtnRefresh                    '<------ Refresh record ------->'
        addRec = False
        editRec = False
       If rsCash.State = adStateOpen Then
          rsCash.Close
        End If
        rsCash.Open "SELECT * From CheckTrans order by check_number", cnBank, adOpenStatic, adLockOptimistic
        Load_DATA
        lvList.SetFocus
End Select
'ERRORHANDLE:
 'errorMsg Err, Me.Name, "Command Button"

End Sub








Private Sub CmdFilter_Click()
      Dirty "Items"
      
  Dim chkNumber As String
  chkNumber = FilterNumber(lblSelected.Caption)

      If rsItem.State = adStateOpen Then
          rsItem.Close
       End If
       rsItem.Open "SELECT * From [CheckItems] WHERE [Check_Number]like'" & chkNumber & "'"
       If rsItem.RecordCount > 0 Then
         Load_ITEMS
         lvItems.SetFocus
       Else
         lvItems.ListItems.Clear
       End If
exitSub:
End Sub



Private Sub Command1_Click()
  PicAccount.Visible = False
End Sub

Private Sub Command2_Click()
 PicVendor.Visible = False
End Sub


Private Sub DTPicker1_CloseUp()
   TxtEntry(2).Text = Format(DTPicker1.Value, "mmm-dd-yyyy")
   TxtEntry(2).SetFocus
End Sub

Private Sub Form_Load()
'// initialized
Set totalBy = New clsTotalBy
With PicVendor
   .Left = 1215
   .Top = 2775
   .Width = 5295
   .Height = 3315
End With
With PicAccount
  .Left = 8055
  .Top = 2655
  .Height = 3435
  .Width = 5175
End With
'// listprint initialize
        Dim i As Integer
        With ListPrint
            .AddItem "Item"
            .AddItem "Description"
            .AddItem "Amount"
        End With
       initPrint = print_Init(ListPrint)
'//

lblAmtInWord.BackStyle = 0
Set convert = New NumToWord
BtnDown(0).Move (TxtEntry(6).Left + TxtEntry(6).Width), (TxtEntry(6).Top), 315, 315
'BtnDown(1).Move (TxtEntry(3).Left + TxtEntry(3).Width), (TxtEntry(3).Top), 315, 315
cmdButtonShow ("1012011"), Me
TbButtonShow "101001111"
SSTab1.Tab = 0

'//
    With LvMenu
       
        .ListItems.Add , "frmPrint", "Print Summary", 1, 1
        .ListItems.Add , "frmSearch", "Search / Filter", 2, 2
        .ListItems.Add , "frmReference", "Reference", 3, 3
        .ListItems.Add , "frmCalcu", "Calculator", 4, 4
        .ListItems.Add , "frmCheckBook", "Check Book", 5, 5
        .ListItems.Add , "help", "Help", 6, 6
        
     End With
'//
Set rsCash = New ADODB.Recordset
rsCash.Open "SELECT * From CheckTrans order by date", cnBank, adOpenStatic, adLockOptimistic
Load_DATA
FlatHeader lvList
Call ShowFldsLabel(Me, rsCash)
'Call Add_Item(rsCash, "Check_Number", CboNumber)

Set rsItem = New ADODB.Recordset
rsItem.Open "SELECT * From CheckItems order by date", cnBank, adOpenStatic, adLockOptimistic
Load_ITEMS
FlatHeader lvItems
Set rsVend = New ADODB.Recordset
Dim SQL As String
SQL = "SELECT Vendor_Name,Address,Vendor_ID "
SQL = SQL & "From Vendor order by Vendor_Name"
rsVend.Open SQL, cnRef, adOpenStatic, adLockOptimistic
Load_Vendor
FlatHeader ListView1
autoAlignCol ListView1
Set rsAcct = New ADODB.Recordset
Dim strSQL As String
strSQL = "SELECT Account,Description "
strSQL = strSQL & "From Expenses order by Description"
rsAcct.Open strSQL, cnRef, adOpenStatic, adLockOptimistic
Load_Expenses
FlatHeader ListView2
autoAlignCol ListView2

errMsg:
  errorMsg Err, Me.Name, "Form Load"
End Sub

Private Sub Load_DATA()
On Error GoTo ERRORHANDLE
If rsCash.RecordCount < 1 Then Exit Sub
'// set columnheaders
Call InsertColumn(lvList, rsCash, 8)
'//set details
Call FillListView(lvList, rsCash, 2, 3)
'//get total
 Call Listview_Total(lvList, rsCash)
ERRORHANDLE:
    errorMsg Err, Me.Name, "Load_Data"
End Sub

Private Sub Load_ITEMS()
On Error GoTo ERRORHANDLE
If rsItem.RecordCount < 1 Then Exit Sub
'// set columnheaders
Call InsertColumn(lvItems, rsItem)
'//set details
Call FillListView(lvItems, rsItem, 2)
'//get total
 Call LvItems_Total(lvItems, rsItem)
ERRORHANDLE:
    errorMsg Err, Me.Name, "Load_Data"
End Sub

Private Sub Load_Vendor()
'// set columnheaders
'Insert_ExtraCol lvList, rsDed
If rsVend.RecordCount = 0 Then Exit Sub
Call InsertColumn(ListView1, rsVend)
'//set details
Call FillListView(ListView1, rsVend, 6)
End Sub

Private Sub Load_Expenses()
'// set columnheaders
'Insert_ExtraCol lvList, rsDed
If rsAcct.RecordCount = 0 Then Exit Sub
Call InsertColumn(ListView2, rsAcct)
'//set details
Call FillListView(ListView2, rsAcct, 4)
End Sub




Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If MsgBox("Exit this system?", vbYesNo + vbQuestion, "Check Transaction") = vbNo Then
  Cancel = 1  'true
Else
  Unload Me
End If
End Sub

Private Sub Form_Resize()
With Me
  If .WindowState = 0 Then
   .Height = 10485
   .Width = 13695
  End If
End With

' SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Set rsCash = Nothing
 Set rsVend = Nothing
 Set rsItem = Nothing
 Set rsAcct = Nothing
 Set frmCashDisburse = Nothing
Unload frmCashDisburse
End Sub

Private Sub ImgItemHelp_Click()
  myMsg "Items entry must be used only " _
  & "for itemized payment!" & vbCrLf & vbCrLf _
  & "", "Items Help", 2, True
End Sub





Private Sub lblSelected_Change()
 CmdFilter_Click
End Sub

Private Sub ListView1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 27 Then
    TxtEntry(6).SetFocus
    PicVendor.Visible = False
  ElseIf KeyCode = 13 Then
   TxtEntry(6).Text = ListView1.SelectedItem.Text
   TxtEntry(7).Text = ListView1.SelectedItem.ListSubItems(1).Text
   TxtEntry(5).Text = ListView1.SelectedItem.ListSubItems(2).Text
   TxtEntry(6).SetFocus
   PicVendor.Visible = False
  End If
End Sub




Private Sub ListView2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 27 Then
    PicAccount.Visible = False
    TxtEntry(3).SetFocus
  ElseIf KeyCode = 13 Then
   TxtEntry(3).Text = ListView2.SelectedItem.Text & ":" & ListView2.SelectedItem.ListSubItems(1).Text
   TxtEntry(3).SetFocus
   PicAccount.Visible = False
  End If
End Sub

Private Sub lvItems_Click()
  If addRec = True Or editRec = True Then Exit Sub
  Call BindDataItems(rsItem, lvItems, True)
End Sub

Private Sub lvItems_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = 38 Or KeyCode = 40 Or KeyCode = 33 Or KeyCode = 34 Then lvItems_Click
End Sub


Private Sub lvList_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = 38 Or KeyCode = 40 Or KeyCode = 33 Or KeyCode = 34 Then LvList_Click
  'Call BindDatasource(Me, rsCash, lvList, True)
  If Val(TxtEntry(4).Text) = 0 Then
     Exit Sub
  Else
    lblAmtInWord = "***" & convert.TOword(TxtEntry(4)) & "***"
  End If
End Sub
Private Sub LvList_Click()
  If addRec = True Or editRec = True Then Exit Sub
  Call BindDatasource(Me, rsCash, lvList, True)
  If Val(TxtEntry(4).Text) = 0 Then
     Exit Sub
  Else
    lblAmtInWord = "***" & convert.TOword(TxtEntry(4)) & "***"
  End If
End Sub


Private Sub lvMenu_DblClick()
MyMenu = LvMenu.SelectedItem.Text
    Select Case LvMenu.SelectedItem.Key
        Case "frmSearch" '//: loadForm frmSearch
          If FormLoaded("frmSearch") = True Then
             MsgBox "Already running !", vbInformation, "Search / Filter"
             Exit Sub
          End If
           SSTab1.Tab = 0
            With frmSearch
               Set .pFindForm = Me
               Set .pFindRecset = rsCash
               Set .pFindCon = cnBank
                   .pFindTABLE = "CheckTrans"
                   .Caption = .Caption & " <:- Check Transaction -:>"
                   .Show
            End With
        Case "frmPrint"  '//: loadForm frmPrint
          If FormLoaded("frmPrint") = True Then
             MsgBox "Already running !", vbInformation, "Print Summary"
             Exit Sub
          End If
          With frmPrint
               Set .pPrintForm = Me
               Set .pPrintRecset = rsCash
               Set .pPrintCon = cnBank
                   .pPrintTABLE = "CheckTrans"
                   .Caption = .Caption & " <:- Check Transaction -:>"
                  .Show
            End With
        Case "frmReference"  ': loadForm frmReference
          If FormLoaded("frmReference") = True Then
             MsgBox "Already running !", vbInformation, "Reference"
             Exit Sub
          Else
             loadForm frmReference
          End If
        Case "frmCalcu" ': loadForm FrmCalcu
          If FormLoaded("FrmCalcu") = True Then
             MsgBox "Already running !", vbInformation, "Calculator"
             Exit Sub
          Else
             loadForm FrmCalcu
          End If
        Case "frmCheckBook" ': loadForm frmCheckBook
          If FormLoaded("frmcheckbook") = True Then
             MsgBox "Already running !", vbInformation, "Check Book"
             Exit Sub
          Else
             loadForm frmCheckBook
          End If
        Case "help"
        HHShowContents Me.hWnd
   End Select
End Sub


Private Sub LvMenu_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = vbLeftButton Then
        Call Drag_It(LvMenu.hWnd)
  End If
End Sub

Private Sub PicAccount_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   PicAccount.Top = 2655
  If Button = vbLeftButton Then
        Call Drag_It(PicAccount.hWnd)
  End If
End Sub









Private Sub PicMenu_Resize()
 LvMenu.Width = PicMenu.Width - LvMenu.Left
 LvMenu.Height = PicMenu.Height
 LvMenu.View = lvwIcon
End Sub



Private Sub PicVendor_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  PicVendor.Top = 2775
  If Button = vbLeftButton Then
        Call Drag_It(PicVendor.hWnd)
  End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
  If PreviousTab = 1 Then
     Dirty "[Items]"
  Else
     Dirty "[Expenses]"
  End If
  If SSTab1.Tab = 1 Then
      If PicVendor.Visible = True Then PicVendor.Visible = False
      If PicAccount.Visible = True Then PicAccount.Visible = False
      'If ListDate.Visible = False Then ListDate.Visible = False
      SelectCheckNum
  End If

  
End Sub

Private Sub SelectCheckNum()
    lblSelected.Caption = ""
    lblSelectedAmt.Caption = ""
    lblSelected.Caption = "Selected Check Number: " & "( " & TxtEntry(1).Text & " )"
    lblSelectedAmt.Caption = "Amount: " & "( " & TxtEntry(4).Text & " )"
End Sub



Private Sub TbMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
  Case "new"
     clearText
     addRec = True
     TbButtonShow "010012212"
     TextEntry(1).Text = TxtEntry(2).Text 'DATE
     TextEntry(2).Text = TxtEntry(1).Text
     TextEntry(1).SetFocus
     nextItemNumber
  Case "save"
     TbButtonShow "101001111"
     Call WriteItems(rsItem, True)
     Dim lngSN As Long
      lngSN = Val(TextEntry(0).Text)
      Call PopulateItems(lvItems, rsItem, 2, lngSN)
     lvItems.SetFocus
     addRec = False
  Case "edit"
       '//VerifyEdit
     Dim src As String  'source
     Dim trg As String  'targer match
     src = TextEntry(0).Text
     trg = lvItems.SelectedItem.Text
        If Trim$(src) <> Trim$(trg) Then
           MsgBox "Please Verify Selected Item# !", vbCritical, "Warning ! [Edit Mode]"
           TbButtonShow "1010011111"
           editRec = False
           Exit Sub
         End If
     '//
     editRec = True
     TbButtonShow "000112212"
     TextEntry(1).SetFocus
  Case "update"
     TbButtonShow "101001111"
     Call WriteItems(rsItem, False)
     LvwReplaceItems rsItem, lvItems
     editRec = False
     lvItems.SetFocus
  Case "cancel"
     TbButtonShow "101001111"
     addRec = False
     editRec = False
     lvItems.SetFocus
  Case "delete"
     Call Delete_Record(rsItem, lvItems)
     lvItems.SetFocus
  Case "refresh"
       If rsItem.State = adStateOpen Then
          rsItem.Close
       End If
       rsItem.Open "SELECT * From CheckItems order by SN", cnBank, adOpenStatic, adLockOptimistic
       Load_ITEMS
       lvItems.SetFocus
  Case "print"
   Dim chkNumber As String
   chkNumber = FilterNumber(lblSelected.Caption)
   If Len(chkNumber) < 1 Then
      MsgBox "Please Select Check Number to print!", vbInformation, "Print Voucher!"
      Exit Sub
   End If
   Call PrintVoucher(rsItem, 1)
End Select
End Sub





Private Sub TextEntry_Change(Index As Integer)
 Select Case Index
   Case Is = 2
      Exit Sub
   Case Is = 6, 7
      TextEntry(8).Text = totalBy.times(TextEntry(6), TextEntry(7))
   End Select
End Sub

Private Sub TextEntry_GotFocus(Index As Integer)
nxTab = Index
Select Case nxTab
Case Is = 9
 TextEntry(nxTab).SelStart = Len(TextEntry(nxTab).Text)
Case Else
 TextEntry(nxTab).SelStart = 0
 TextEntry(nxTab).SelLength = Len(TextEntry(nxTab).Text)
End Select
If addRec = False Or editRec = False Then
  TextEntry(nxTab).Locked = True
End If
If addRec = True Or editRec = True Then
   TextEntry(nxTab).Locked = False
   TextLocked Me, List2
End If
End Sub

Private Sub TextEntry_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim lastTab As Integer
On Error GoTo ERRORHANDLE
lastTab = 8
If KeyCode = 13 Then
     If nxTab = lastTab Then Exit Sub
     nxTab = nxTab + 1
     If nxTab = 2 Then nxTab = 3
ElseIf KeyCode = 38 Then  'up arrow key

     If nxTab = 0 Or nxTab = 1 Then Exit Sub
     nxTab = nxTab - 1
     If nxTab = 2 Then nxTab = 1
End If
TextEntry(nxTab).SetFocus
ERRORHANDLE:
 errorMsg Err, Me.Name
End Sub

Private Sub TextEntry_LostFocus(Index As Integer)
If Not IsDate(TextEntry(1).Text) Then
    TextEntry(1).SetFocus
End If
End Sub

Private Sub txtEntry_Change(Index As Integer)
Select Case Index
Case Is = 9
  If TxtEntry(9).Text = "Y" Then
   ChkPrint.Value = 1
  Else
   ChkPrint.Value = 0
  End If
End Select
End Sub

Private Sub TbButtonShow(ByRef buttonString As String)
'< syntax:  TbButtonShow ("000111112")
''--------------------------------------------------
''-- This routine handles setting the enabled --
''-- to true / false on the buttons.                --
''-------------------------------------------------
''-- A string of 0101 passed. If 0, disabled   --
''-------------------------------------------------
Dim indx As Integer
buttonString = Trim$(buttonString)
For indx = 1 To Len(buttonString)
  If (Mid$(buttonString, indx, 1) = "1") Then
     TbMenu.Buttons(indx).Enabled = True
     TbMenu.Buttons(indx).Visible = True    '(index-1) use only if index start from 0
  ElseIf (Mid$(buttonString, indx, 1) = "0") Then
     TbMenu.Buttons(indx).Visible = False
  ElseIf (Mid$(buttonString, indx, 1) = "2") Then
     TbMenu.Buttons(indx).Enabled = False
  End If
  Next indx
End Sub


Private Sub WriteItems(ByRef srcRS As Recordset, _
                      ByVal newRec As Boolean, _
                      Optional ByVal srcNumFlds As Integer = 0)
'//addnew = true for new record else false > forced
'//srcnumflds = number of fields loaded in textbox  > optional
                'if not all fields are loaded, srcnumflds is equal to text upperbound indeces
                'based on the numbers of textbox showed in the form (see enabled textbox procedures)
If srcRS Is Nothing Then Exit Sub
If srcRS.RecordCount > 0 Then
If srcRS.EOF = True Or srcRS.BOF = True Then
   'MsgBox "Either EOF or BOF reached.", vbInformation, "Write Data!"
   'Exit Sub
   srcRS.MoveLast
End If
End If
Dim i As Integer
Dim NOF As Integer 'Number Of Feilds
If srcNumFlds > 0 Then
   NOF = srcNumFlds
Else
   NOF = (srcRS.Fields.Count - 1)  'remember that indeces are zero based
End If
ReDim entries(NOF) As TextBox
For i = 0 To NOF
    Set entries(i) = TextEntry(i)  'm tired of using frm, set number of elements allowed
    Next i
i = 0
With srcRS
  If newRec = True Then
      .AddNew
  End If
      For i = 0 To NOF
      Select Case srcRS.Fields.Item(i).Type
       Case Is = 3   'integer
           If IsNumeric(entries(i).Text) Then
              srcRS.Fields(i) = toNumber(entries(i).Text)
           End If
      Case Is = 5, 6  'currency or double
           If IsNumeric(entries(i).Text) Then
             srcRS.Fields(i) = toMoney(entries(i).Text)
           End If
       Case Is = 7   'date
           If IsDate(entries(i).Text) Then
               srcRS.Fields(i) = CDate(entries(i).Text)
           Else '//save empty entry
               srcRS.Fields(i) = Null
           End If
       Case Is = 202, 203    'text, memo
             srcRS.Fields(i) = CStr(entries(i).Text)
      End Select
      Next i
      .Update
End With
End Sub
Private Sub BindDataItems(ByRef srcRS As Recordset, _
                          ByRef lv As ListView, _
                          Optional ByVal findFirst As Boolean = True, _
                          Optional ByVal numOfFlds As Integer = 0)
'//findFIRST - optional/false when use for next,previous,last,first
   If srcRS Is Nothing Then Exit Sub
With srcRS
  If .RecordCount = 0 Then
      Exit Sub
   End If
End With
If findFirst = False Then
 If srcRS.EOF = True Then
    MsgBox "EOF reached.", vbInformation, "Bind Data!"
    Exit Sub
 ElseIf srcRS.BOF = True Then
   MsgBox "BOF reached.", vbInformation, "Bind Data!"
   Exit Sub
 End If
End If
Dim abPos As Boolean   'absolutePosition
Dim i As Integer
Dim strFind As String
Dim strMatch As String
Dim NOF As Integer 'Number Of Feilds
'//
If srcRS.RecordCount = 0 Then Exit Sub
'// initialized
If numOfFlds > 0 Then
   NOF = numOfFlds
Else
   NOF = (srcRS.Fields.Count - 1)  'remember that indeces are zero based
End If
For i = 0 To NOF
   TextEntry(i) = Empty
   Next i
If IsNumeric(TrimSpaces(CStr(lv.SelectedItem.Text))) Then
    strFind = TrimSpaces(CStr(lv.SelectedItem.Text))
    abPos = False
Else
    strFind = lv.SelectedItem.Index
    abPos = True
End If
If findFirst = True Then
 With srcRS
 .MoveFirst
   Do Until srcRS.EOF
   If abPos = False Then
        lv.MousePointer = vbHourglass
       strMatch = TrimSpaces(CStr(toNumber(srcRS.Fields(0))))
     Else
       lv.MousePointer = vbHourglass
       'slower//i use only on alpha type// so you can show the value one
       'row even if there is duplicate reference for viewing record
       'remember that reference must be a unique key
       strMatch = srcRS.Bookmark '// .AbsolutePosition
   End If
   If strMatch = strFind Then
         lv.MousePointer = vbDefault

         GoTo iFound
   Else
     .MoveNext
   End If
   Loop
 End With
 lv.MousePointer = vbDefault
End If 'findFirst
iFound:
With srcRS
         If srcRS.EOF = True Or srcRS.BOF = True Then Exit Sub
         For i = 0 To NOF
          If Not IsNull(srcRS.Fields(i)) Then
             TextEntry(i) = FormatRS(srcRS.Fields(i))
              If srcRS.Fields(i).Type = 6 Or srcRS.Fields(i).Type = 5 Then
                TextEntry(i).Alignment = 1
                 If Val(TextEntry(i)) = 0 Then
                      TxtEntry(i).ForeColor = &HD38545
                 ElseIf Val(TextEntry(i)) < 0 Then
                      TextEntry(i).ForeColor = vbRed      ' if the value is negative
                 Else
                      TextEntry(i).ForeColor = vbBlack
                End If
             Else                                          'string value and non-zero value
                   TextEntry(i).ForeColor = vbBlack
            End If
          Else
              TextEntry(i) = Empty
          End If
         Next i
    '//end of Search
End With

End Sub

Private Sub Dirty(ByVal sWhere As String)
Dim Msg As String
      Msg = "You have pending records to save..."
      Msg = Msg & vbCrLf & "in " & sWhere & " !"
  If addRec = True Then
      MsgBox Msg, vbCritical, "Add"
    Exit Sub
  ElseIf editRec = True Then
     MsgBox Msg, vbCritical, "Edit"
    Exit Sub
  End If
End Sub

Private Sub txtEntry_GotFocus(Index As Integer)
nxTab = Index

Select Case nxTab
Case Is = 3
 BtnDown(0).Visible = False
 'BtnDown(1).Move (TxtEntry(3).Left + TxtEntry(3).Width), (TxtEntry(3).Top), 315, 315
 If addRec = True Or editRec = True Then
    BtnDown(1).Visible = True
 End If
Case Is = 6
 BtnDown(1).Visible = False
 BtnDown(0).Move (TxtEntry(6).Left + TxtEntry(6).Width), (TxtEntry(6).Top), 315, 315
 If addRec = True Or editRec = True Then
     BtnDown(0).Visible = True
 End If
Case Is = 7, 8
 TxtEntry(nxTab).SelStart = Len(TxtEntry(nxTab).Text)
Case Else
 TxtEntry(nxTab).SelStart = 0
 TxtEntry(nxTab).SelLength = Len(TxtEntry(nxTab).Text)
End Select
 TxtEntry(nxTab).BackColor = &HFFFFFF
If addRec = False Or editRec = False Then
  TxtEntry(nxTab).Locked = True
End If
If addRec = True Or editRec = True Then
   TxtEntry(nxTab).Locked = False
   TxtLocked Me, List1
End If
End Sub

Private Sub TxtEntry_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim lastTab As Integer
On Error GoTo ERRORHANDLE
lastTab = 8  ' rsCash.Fields.Count - 1 'or txtEntry Upper Bound if kung limitado lang ang textbox
If KeyCode = 13 Then
     If nxTab = lastTab Then Exit Sub
'     If nxTab = 3 Then
'        nxTab = Index                              'stay foot ka lang
'        Exit Sub
'     End If
     nxTab = nxTab + 1
     If nxTab = 4 Then nxTab = 1                   '//current tab is 3 *passed 4 punta ka ng 1
     If nxTab = 3 Then nxTab = 4                   '//current tab is 2 *passed 3 punta ka ng 4
     If nxTab = 5 Then nxTab = 6                   '//current tab is 4 *Passed 5 punta ka ng 6
     If nxTab = 7 Then nxTab = 3                   '//current tab is 6 *Passed 3 punta ka ng 3
ElseIf KeyCode = 38 Then  'up arrow key
     If nxTab = 0 Or nxTab = 1 Then Exit Sub
     nxTab = nxTab - 1
     If nxTab = 2 Then nxTab = 6                   '//current tab is 3 *passed 2 balik ka 6
     If nxTab = 5 Then nxTab = 4                   '//current tab is 6 *passed 5 balik ka sa 4
     If nxTab = 3 Then nxTab = 2

End If
TxtEntry(nxTab).SetFocus
ERRORHANDLE:
 errorMsg Err, Me.Name
End Sub

Private Sub txtEntry_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case Index
  Case Is = 6
    If KeyCode = 113 Then
       PicVendor.Visible = True
       ListView1.SetFocus
    End If
 Case Is = 3
    If KeyCode = 113 Then
       PicAccount.Visible = True
       ListView2.SetFocus
    End If
End Select
End Sub

Private Sub txtEntry_LostFocus(Index As Integer)
nxTab = Index
TxtEntry(nxTab).BackColor = &HE4F0B7
If Not IsDate(TxtEntry(2).Text) Then
    TxtEntry(2).SetFocus
End If
End Sub

'//procedure to show in listview what you have edited
'//need not refresh every time you edit record
'//<< syntax >> Call LvwReplaceData(Me, rs, lvname)
Private Sub LvwReplaceItems(ByRef rs As Recordset, _
                      ByRef lv As ListView, _
                      Optional ByVal numOfFlds As Integer = 0)
Dim i As Integer
Dim NOF As Integer  'number of fields
If numOfFlds > 0 Then
   NOF = numOfFlds
Else
   NOF = (rs.Fields.Count - 1)  'remember that indeces are zero based
End If
For i = 1 To NOF
  lv.SelectedItem.ListSubItems(i).Text = TextEntry(i).Text
  Next i
End Sub

Private Sub clearTxt()
Dim i As Integer
For i = TxtEntry.LBound To TxtEntry.UBound
        TxtEntry(i).Text = ""
        Next i
End Sub

Private Sub clearText()
Dim i As Integer
For i = TextEntry.LBound To TextEntry.UBound
        TextEntry(i).Text = ""
        Next i
End Sub

'Procedure used to show in listview what you have already added
'need not refresh every time you add new record
'coded by edwin delos santos
'<< syntax >> PopulateLvw lvname, rs, 2
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
           With X.ListSubItems.Add(, , TxtEntry(i).Text)
           End With
          Next i
End Sub


Public Sub PopulateItems(ByRef sListView As ListView, _
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
           With X.ListSubItems.Add(, , TextEntry(i).Text)
           End With
          Next i
End Sub



Private Sub CmdFirst_Click()
If rsCash Is Nothing Then Exit Sub
   rsCash.MoveFirst
 Call BindDatasource(Me, rsCash, lvList, False)
 SelectCheckNum
 Call ListView_Search(lvList, TxtEntry(0).Text, 0)
End Sub

Private Sub CmdLast_Click()
If rsCash Is Nothing Then Exit Sub
 rsCash.MoveLast
 Call BindDatasource(Me, rsCash, lvList, False)
 SelectCheckNum
 Call ListView_Search(lvList, TxtEntry(0).Text, 0)
End Sub

Private Sub CmdNext_Click()
 If rsCash Is Nothing Then Exit Sub
 If rsCash.EOF = True Then rsCash.MoveLast
 rsCash.MoveNext
 Call BindDatasource(Me, rsCash, lvList, False)
SelectCheckNum
Call ListView_Search(lvList, TxtEntry(0).Text, 0)
End Sub

Private Sub CmdPrev_Click()
If rsCash Is Nothing Then Exit Sub
If rsCash.BOF = True Then rcdSet.MoveFirst
rsCash.MovePrevious
Call BindDatasource(Me, rsCash, lvList, False)
SelectCheckNum
Call ListView_Search(lvList, TxtEntry(0).Text, 0)
End Sub

Private Sub Headers()
 Dim Dayt As Variant
 Dayt = Format(Now(), "SHORT DATE")
         Printer.Print Tab(110); "Date:" & Dayt
         Printer.Print Tab(6); "Payee   : " & TxtEntry(6).Text;
         Printer.Print Tab(6); "___________________________________________________________________________________________________"
         Printer.Print Tab(6); "Check#/Date"; Tab(30); "Description"; Tab(110); "Amount"
         Printer.Print Tab(6); TxtEntry(1).Text; Tab(30); TxtEntry(3).Text
         Printer.Print Tab(6); TxtEntry(2).Text; Tab(110); TxtEntry(4).Text
         Printer.Print Tab(6); "-----------------------------------------------------------------------------------------------------------------------------------------------------------------------"
End Sub

Private Sub PrintVoucher(ByRef srcRS As Recordset, Optional ByVal mOrient As Long = 2)
'// coded by edwin delos santos
'// fixed settings has been made to this report
'// you may adjust settings using variables ...
'// like default:  tab, page orientation, number of pages to print, quality and so on...
'// no modify except for the above said defaul settings ...
On Error GoTo printErr
Dim chkNumber As String
Dim Msg As String
chkNumber = FilterNumber(lblSelected.Caption)
Dim curr_Rec As Long  'current record looping becomes 0 if max line per page is reaced
Dim rec_Counter As Long 'record  counter
Dim recPerPage As Long
Dim ans 'answer
Dim strFont As String, sngSize As Single
'//initialize
rec_Counter = 0
curr_Rec = 0
'If srcRS.RecordCount < 1 Then Exit Sub
If mOrient = 2 Then
   recPerPage = 25
End If
If mOrient = 1 Then
   recPerPage = 35
End If
'//
Msg = Msg & vbCrLf & "Check#:" & chkNumber
Msg = Msg & vbCrLf & "Date  :" & TxtEntry(2).Text
Msg = Msg & vbCrLf & "Vendor:" & TxtEntry(6).Text
Msg = Msg & vbCrLf & vbCrLf & "Proceed?"
ans = MsgBox(Msg, vbYesNo + vbQuestion, "Print Voucher")
  If ans = vbYes Then
'//save current printer settings
         strFont = Printer.Font
         sngSize = Printer.FontSize
         Printer.Orientation = mOrient    '2   'Landscape
         Printer.Font = "ms sans serif"
         Printer.FontSize = 9
         Printer.Print
         Printer.Print
         Printer.Print
         Call prnCenterText("CHECK VOUCHER", 140) 'company
'         Printer.FontBold = False
'         Printer.FontSize = 9
         Printer.Print
         Printer.Print
         Printer.Print
 
'// headers
         Headers             '< --------------------  headers ----------------->
        
'** if no items to print
        If srcRS.RecordCount < 1 Then
           Printer.Print
           Call prnCenterText("------------ no item ------------", 120)
           Printer.Print
           Printer.Print Tab(6); "___________________________________________________________________________________________________"
           Printer.Print Tab(90); "T O T A L >>"; Tab(110); TxtEntry(4).Text
           Printer.Print Tab(6); "==============================================================================================="
           Printer.Print
           Printer.Print Tab(6); "Received by : _______________________________"; Tab(70); "Approved by: _________________________"
           Printer.Print
           Printer.EndDoc
           MsgBox "D O N E !!", vbInformation, "Printing..."
           Exit Sub
        End If
'**
       With srcRS
            .MoveFirst
              Printer.Font.Underline = True
              Call Print_Headings(srcRS, 15, ListPrint, 12)
              Printer.Font.Underline = False
         While Not .EOF = True
'//print details
                Call Print_Details(srcRS, 15, ListPrint)  'see procedure
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
'                   Headers   'next page headers
                   Printer.Print
                   Printer.Font.Underline = True
                   Call Print_Headings(srcRS, 15, ListPrint, 12) '2nd page headings
                   Printer.Font.Underline = False
                   Printer.Print
                End If 'rec_count
            End If   'curr_Rec
        Wend  'while not EOF
'//print line
        If mOrient = 1 Then
          Call prnCenterText("------------ nothing follows ------------", 120)
'          Call Print_Total(srcRS, 15, ListPrint)
          Printer.Print Tab(6); "___________________________________________________________________________________________________"
        End If
'// print Total
         Printer.Print Tab(90); "T O T A L >>"; Tab(110); TxtEntry(4).Text
         Printer.Print Tab(6); "=============================================================================================="
          Printer.Print
          Printer.Print Tab(6); "Received by : _______________________________"; Tab(70); "Approved by: _________________________"
'          Call Print_Total(srcRS, 20, ListPrint)
          Printer.Print
'//Footer
          'Footers             '< --------------------  footers ----------------->
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


Public Sub LvItems_Total(ByRef Lvw As ListView, ByRef srcRS As Recordset)
   With srcRS
     If .RecordCount < 1 Then
          Exit Sub
     End If
   End With
Dim rec_Count As Long
Dim isCurr As Boolean    'flag for currency or double
Dim X As ListItem
Dim strvalue As String
Dim iCount As Long  'to determine the last record
Dim i As Integer
Dim NumOfFields As Integer
On Error Resume Next
'//initialize value
isCurr = False
NumOfFields = (srcRS.Fields.Count - 1)
ReDim dblTotal(NumOfFields) As Double    'number of elements
For i = 0 To NumOfFields
   dblTotal(i) = 0
   Next i
   i = i + 1
iCount = 0
With srcRS
   rec_Count = CStr(srcRS.RecordCount)
   Set X = Lvw.ListItems.Add(, , "(" & rec_Count & ")" & "Record")
       X.Bold = True
       X.ForeColor = vbBlue
  .MoveFirst
While Not .EOF = True
iCount = iCount + 1
  For i = 1 To NumOfFields
       strvalue = Empty
      If Not IsNull(srcRS.Fields(i)) Then
             If srcRS.Fields(i).Type = 6 Then  'currency
                 strvalue = toMoney(srcRS.Fields(i))
                 isCurr = True

             Else
                 strvalue = ""
             End If
             '// dblTotal(0),dblTotal(1),dblTotal(2) and so on ...
             If isCurr = True Then
                 dblTotal(i) = dblTotal(i) + Val(strvalue)
             End If
            If iCount = .RecordCount Then
                   If dblTotal(i) > 0 Then
                      If isCurr = True Then
                        With X.ListSubItems.Add(, , Format(dblTotal(i), "standard"))
                           X.ListSubItems(i).Bold = True
                           X.ListSubItems(i).ForeColor = vbRed
                        End With
                      Else
                        With X.ListSubItems.Add(, , toNumber(dblTotal(i)))
                           X.ListSubItems(i).Bold = True
                           X.ListSubItems(i).ForeColor = vbRed
                        End With
                      End If
                    Else  '//dbltotal() = 0
                      With X.ListSubItems.Add(, , " - ")
                      End With
                    End If
           End If   '//icount
       Else
        If iCount = .RecordCount Then
          With X.ListSubItems.Add(, , " - ") 'if null string
          End With
        End If
       End If 'not isnull
    Next i
.MoveNext
Wend
End With
Set X = Nothing
End Sub

