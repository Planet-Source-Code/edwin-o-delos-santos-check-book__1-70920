VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCheckBook 
   Caption         =   "Bank Account"
   ClientHeight    =   9210
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12225
   HelpContextID   =   160
   Icon            =   "frmCheckBook.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmCheckBook.frx":1782
   ScaleHeight     =   9210
   ScaleWidth      =   12225
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox PicLv 
      Appearance      =   0  'Flat
      BackColor       =   &H00FDFDEE&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   120
      ScaleHeight     =   195
      ScaleWidth      =   165
      TabIndex        =   32
      Top             =   5160
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox PicTrans 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   5040
      ScaleHeight     =   645
      ScaleWidth      =   3225
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   3255
      Begin MSComctlLib.ListView ListView1 
         Height          =   3135
         Left            =   0
         TabIndex        =   8
         Top             =   240
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   5530
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "i16x16"
         SmallIcons      =   "i16x16"
         ForeColor       =   12582912
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
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
         Left            =   7080
         Style           =   1  'Graphical
         TabIndex        =   10
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
         TabIndex        =   9
         Top             =   0
         Width           =   1635
      End
   End
   Begin CHECKBOOK.Hline Hline1 
      Height          =   30
      Left            =   240
      TabIndex        =   31
      Top             =   2760
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   53
   End
   Begin VB.TextBox txtEntry 
      Alignment       =   1  'Right Justify
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
      Index           =   4
      Left            =   6240
      TabIndex        =   19
      Top             =   1200
      Width           =   1575
   End
   Begin VB.TextBox txtEntry 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   325
      Index           =   2
      Left            =   2040
      TabIndex        =   18
      Top             =   1440
      Width           =   3555
   End
   Begin VB.TextBox txtEntry 
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
      Height          =   325
      Index           =   3
      Left            =   2040
      TabIndex        =   17
      Top             =   1800
      Width           =   3555
   End
   Begin VB.TextBox txtEntry 
      Alignment       =   1  'Right Justify
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
      Left            =   7920
      TabIndex        =   16
      Top             =   1200
      Width           =   1575
   End
   Begin VB.TextBox txtEntry 
      Alignment       =   1  'Right Justify
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
      Index           =   6
      Left            =   9600
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   1200
      Width           =   1575
   End
   Begin VB.TextBox txtEntry 
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
      Height          =   300
      Index           =   8
      Left            =   6240
      TabIndex        =   14
      Text            =   "Y"
      Top             =   1920
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox txtEntry 
      BackColor       =   &H00E8FBFB&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   325
      Index           =   0
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   720
      Width           =   2055
   End
   Begin VB.TextBox txtEntry 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   325
      Index           =   1
      Left            =   2040
      TabIndex        =   12
      Top             =   1080
      Width           =   2055
   End
   Begin VB.TextBox txtEntry 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   7
      Left            =   2040
      TabIndex        =   11
      Top             =   2160
      Width           =   3555
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   200
      Left            =   10320
      TabIndex        =   5
      Top             =   3480
      Value           =   1  'Checked
      Width           =   200
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   0
      Top             =   8865
      Width           =   12225
      _ExtentX        =   21564
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   15901
            Text            =   "<:- system designed by:  edwinSoftware  (c)-2008  all rights reserved -:>"
            TextSave        =   "<:- system designed by:  edwinSoftware  (c)-2008  all rights reserved -:>"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "8/1/2008"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "10:04 PM"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox PicTopBar 
      BackColor       =   &H0000FFFF&
      Height          =   375
      Left            =   240
      ScaleHeight     =   315
      ScaleWidth      =   915
      TabIndex        =   2
      Top             =   3720
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSComctlLib.ImageList ImgToolbar 
      Left            =   7080
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckBook.frx":5260
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckBook.frx":59DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckBook.frx":6154
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckBook.frx":68CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckBook.frx":7048
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckBook.frx":77C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckBook.frx":7F3C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckBook.frx":86B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckBook.frx":8E30
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvList 
      Height          =   3435
      Left            =   0
      TabIndex        =   3
      Top             =   4800
      Width           =   2865
      _ExtentX        =   5054
      _ExtentY        =   6059
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
      BackColor       =   16777215
      BorderStyle     =   1
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
   End
   Begin MSComctlLib.Toolbar TbMenu 
      Height          =   660
      Left            =   480
      TabIndex        =   4
      Top             =   3000
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   1164
      ButtonWidth     =   1191
      ButtonHeight    =   1164
      Style           =   1
      ImageList       =   "ImgToolbar"
      DisabledImageList=   "ImgToolbar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
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
            Object.ToolTipText     =   "Print Summary"
            ImageIndex      =   8
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Search"
            Key             =   "search"
            Object.ToolTipText     =   "Search / Filter"
            ImageIndex      =   9
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList i16x16 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   21
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckBook.frx":95AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckBook.frx":9944
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckBook.frx":9CDE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckBook.frx":A078
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckBook.frx":A412
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckBook.frx":A7AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckBook.frx":B1BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckBook.frx":B212
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckBook.frx":B5AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckBook.frx":B946
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckBook.frx":BCE0
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckBook.frx":C07A
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckBook.frx":CA8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckBook.frx":D49E
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckBook.frx":DEB0
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckBook.frx":E8C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckBook.frx":F2D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckBook.frx":FCE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckBook.frx":106F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckBook.frx":10C94
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckBook.frx":11230
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblFLDi 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000C&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   330
      Index           =   0
      Left            =   480
      TabIndex        =   30
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label lblFLDi 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000C&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   330
      Index           =   1
      Left            =   480
      TabIndex        =   29
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label lblFLDi 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000C&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   330
      Index           =   2
      Left            =   480
      TabIndex        =   28
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label lblFLDi 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000C&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   330
      Index           =   3
      Left            =   480
      TabIndex        =   27
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label lblFLDi 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000C&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   330
      Index           =   4
      Left            =   6240
      TabIndex        =   26
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label lblFLDi 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000C&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   330
      Index           =   5
      Left            =   7920
      TabIndex        =   25
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label lblFLDi 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000C&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   330
      Index           =   6
      Left            =   9600
      TabIndex        =   24
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label lblFLDi 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000C&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   330
      Index           =   7
      Left            =   480
      TabIndex        =   23
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000009&
      Height          =   1815
      Index           =   1
      Left            =   6000
      Top             =   720
      Width           =   5415
   End
   Begin VB.Label lblBalance 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   8400
      TabIndex        =   22
      Top             =   2040
      Width           =   2775
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Balance after expenses:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   6240
      TabIndex        =   21
      Top             =   2040
      Width           =   2070
   End
   Begin VB.Label lblF2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "[ F2 ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   4200
      TabIndex        =   20
      Top             =   1080
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Label lblFLDi 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Cleared"
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
      Index           =   8
      Left            =   10680
      TabIndex        =   6
      Top             =   3480
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Check Book"
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
      Left            =   240
      TabIndex        =   1
      Top             =   0
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   1815
      Index           =   0
      Left            =   6000
      Top             =   720
      Width           =   5415
   End
End
Attribute VB_Name = "frmCheckBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsChkBook As Recordset
Dim rsCash As Recordset

Dim rsTemp As Recordset


Private Sub Check1_Click()
Dim chk As Integer
chk = Check1.Value
  If chk = 1 Then
    txtEntry(8).Text = "Y"
  Else
    txtEntry(8).Text = "n"
  End If
End Sub




Private Sub ShowBalance()
      If rsChkBook.RecordCount < 1 Then Exit Sub
      
      Set rsTemp = New ADODB.Recordset
       Set rsTemp = rsChkBook
        If rsTemp.State = adStateOpen Then
           rsTemp.Close
        End If
      rsTemp.Open "SELECT * From CheckBook order by SN", cnBank
      lblBalance.Caption = Format(CheckBalance(rsTemp, 4, 5), "standard")
      Set rsTemp = Nothing
End Sub

Private Sub Command2_Click()
PicTrans.Visible = False
End Sub

Private Sub Form_Load()
With PicTrans
 .Top = 1455
 .Left = 2175
 .Height = 3435
 .Width = 7335
End With

Show
lvList.SetFocus

'// List BackColour Formatting
Call SetListViewColor(lvList, PicLv, vbWhite, &HFDFDEE, 0.1)

TbButtonShow "101001111"
lblBalance.BackStyle = 0
Set rsChkBook = New ADODB.Recordset
rsChkBook.Open "SELECT * From CheckBook order by SN", cnBank, adOpenStatic, adLockOptimistic
Call ShowFldsLabel(Me, rsChkBook)
Load_DATA

Set rsCash = New ADODB.Recordset
Dim SQL As String
SQL = "SELECT DATE,CHECK_NUMBER,DESCRIPTION,VENDOR_NAME,AMOUNT "
SQL = SQL & "From CHECKTRANS order by CHECK_NUMBER"
rsCash.Open SQL, cnBank, adOpenStatic, adLockOptimistic
Load_CASHTRANS

End Sub

Private Sub Load_DATA()
'On Error GoTo ERRORHANDLE
If rsChkBook.RecordCount < 1 Then Exit Sub
'// set columnheaders
Call InsertColumn(lvList, rsChkBook, 7)
'//set details
Call FillListView(lvList, rsChkBook, 2)
'//get total
Call Balance_Total(lvList, rsChkBook, 4, 5)
'ERRORHANDLE:
'    errorMsg Err, Me.Name, "Load_Data"
End Sub
Private Sub Load_CASHTRANS()
On Error GoTo ERRORHANDLE
If rsCash.RecordCount < 1 Then Exit Sub
'// set columnheaders
Call InsertColumn(ListView1, rsCash)
'//set details
Call FillListView(ListView1, rsCash, 3)
ERRORHANDLE:
    errorMsg Err, Me.Name, "Load_Data"
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If MsgBox("Exit this system?", vbYesNo + vbQuestion, "Check Book") = vbNo Then
  Cancel = 1 'true
Else
  Unload Me
End If

End Sub

Private Sub Form_Resize()
On Error Resume Next
  If WindowState <> vbMinimized Then
       If Me.Width < 112345 Then Me.Width = 12345
       If Me.Height < 9525 Then Me.Height = 9525
  
          CoolBar1.Width = ScaleWidth
          lvList.Width = Me.ScaleWidth
          lvList.Top = PicTopBar.Top
          lvList.Height = (Me.ScaleHeight - 4095)

  End If
SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
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

Private Sub Form_Unload(Cancel As Integer)
Set rsChkBook = Nothing
Set rsCash = Nothing
Set rsTemp = Nothing
Set frmCheckBook = Nothing
Unload frmCheckBook
End Sub

Private Sub ListView1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 27 Then
    PicTrans.Visible = False
    txtEntry(1).SetFocus
  ElseIf KeyCode = 13 Then
   txtEntry(1).Text = ListView1.SelectedItem.Text
   txtEntry(2).Text = ListView1.SelectedItem.ListSubItems(1).Text
   txtEntry(3).Text = ListView1.SelectedItem.ListSubItems(2).Text
   txtEntry(7).Text = ListView1.SelectedItem.ListSubItems(3).Text
   txtEntry(4).Text = ListView1.SelectedItem.ListSubItems(4).Text
   txtEntry(5).SetFocus
   PicTrans.Visible = False
  End If
End Sub

Private Sub LvList_Click()
 If addRec = True Or editRec = True Then Exit Sub
 Call BindDatasource(Me, rsChkBook, lvList, True)
End Sub

Private Sub lvList_KeyUp(KeyCode As Integer, Shift As Integer)
 If KeyCode = 38 Or KeyCode = 40 Or KeyCode = 33 Or KeyCode = 34 Then LvList_Click
End Sub

Private Sub Picture2_Click()

End Sub



Private Sub PicEntry_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = vbLeftButton Then
        Call Drag_It(PicEntry.hWnd)
  End If
End Sub

Private Sub nextNumber()
      Dim nextNo As Long
      Set rsTemp = New ADODB.Recordset
       Set rsTemp = rsChkBook
        If rsTemp.State = adStateOpen Then
           rsTemp.Close
        End If
      rsTemp.Open "SELECT * From CheckBook order by SN", cnBank
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


Private Sub PicTrans_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   PicTrans.Top = 1455
  If Button = vbLeftButton Then
        Call Drag_It(PicTrans.hWnd)
  End If
End Sub

Private Sub TbMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
  Case "new"
     clearText
     addRec = True
     txtEntry(8).Text = "Y"
     TbButtonShow "0100122122"
     nextNumber
     ShowBalance
     txtEntry(1).SetFocus
  Case "save"
     TbButtonShow "1010011111"
     Call WriteData(Me, rsChkBook, True)
      Dim lngSN As Long
      lngSN = Val(txtEntry(0).Text)
      Call PopulateData(lvList, rsChkBook, 2, lngSN)
     lvList.SetFocus
     addRec = False
  Case "edit"
     '//VerifyEdit
     Dim src As String  'source
     Dim trg As String  'targer match
     src = txtEntry(0).Text
     trg = lvList.SelectedItem.Text
        If Trim$(src) <> Trim$(trg) Then
           MsgBox "Please Verify Selected Item# !", vbCritical, "Warning !!!"
           TbButtonShow "1010011111"
           editRec = False
           Exit Sub
         End If
     '//
     editRec = True
     TbButtonShow "0001122122"
     txtEntry(1).SetFocus
  Case "update"
     TbButtonShow "1010011111"
     Call WriteData(Me, rsChkBook, False)
     LvwReplaceData Me, rsChkBook, lvList
     editRec = False
     lvList.SetFocus
  Case "cancel"
     TbButtonShow "1010011111"
     addRec = False
     editRec = False
     lvList.SetFocus
  Case "delete"
     Call Delete_Record(rsChkBook, lvList)
  Case "refresh"
       If rsChkBook.State = adStateOpen Then
          rsChkBook.Close
       End If
       rsChkBook.Open "SELECT * From CheckBook order by SN", cnBank, adOpenStatic, adLockOptimistic
       Load_DATA
       lvList.SetFocus
 Case "search"
           With frmSearch
               Set .pFindForm = Me
               Set .pFindRecset = rsChkBook
               Set .pFindCon = cnBank
                   .pFindTABLE = "CheckBook"
                   .Caption = .Caption & " <:- Check Book -:>"
                   .Show
            End With
 Case "print"
         With frmPrint
               Set .pPrintForm = Me
               Set .pPrintRecset = rsChkBook
               Set .pPrintCon = cnBank
                   .pPrintTABLE = "CheckBook"
                   .Caption = .Caption & " <:- Check Book -:>"
                  .Show
            End With
End Select
End Sub

Private Sub txtEntry_Change(Index As Integer)
Select Case Index
Case Is = 8
  If txtEntry(8).Text = "Y" Then
      Check1.Value = 1
  Else
      Check1.Value = 0
  End If
Case Is = 4, 5
   If addRec = True Or editRec = True Then
      If Val(txtEntry(5).Text) = 0 Then
        txtEntry(6) = Val(toMoney(lblBalance.Caption)) - Val(txtEntry(4).Text)
      End If
      If Val(txtEntry(4).Text) = 0 Then
        txtEntry(6) = Val(toMoney(lblBalance.Caption)) + Val(txtEntry(5).Text)
      End If
    End If
End Select

End Sub

Private Sub txtEntry_GotFocus(Index As Integer)
nxTab = Index
Select Case nxTab
Case Is = 1
 If addRec = True Or editRec = True Then
    lblF2.Visible = True
 End If
Case Is = 7
 txtEntry(nxTab).SelStart = Len(txtEntry(nxTab).Text)
Case Else
 txtEntry(nxTab).SelStart = 0
 txtEntry(nxTab).SelLength = Len(txtEntry(nxTab).Text)
End Select
End Sub

Private Sub TxtEntry_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim lastTab As Integer
On Error GoTo ERRORHANDLE
lastTab = 5
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
Private Function CheckBalance(ByRef srcRS As Recordset, dr_fld As Integer, cr_fld As Integer) As String
'// which index value to compute Total
'<< syntax >>
'Label2.Caption = Format(CheckBalance(rsChkBook, 4, 5), "standard")
Dim strvalue As String
Dim iCount As Integer
Dim i As Integer
Dim drAmt As Double
Dim crAmt As Double
Dim balAmt As Double
Dim NumOfFields As Integer
'//initialize value
NumOfFields = (srcRS.Fields.Count - 1)
ReDim dblTotal(NumOfFields) As Double
For i = 0 To NumOfFields
   dblTotal(i) = 0
   Next i
iCount = 0
With srcRS
  .MoveFirst
If srcRS.EOF = True Or srcRS.BOF = True Then Exit Function
While Not .EOF = True
iCount = iCount + 1
  For i = 0 To NumOfFields
       strvalue = Empty
      If Not IsNull(srcRS.Fields(i)) Then
        strvalue = srcRS.Fields(i)
        If IsNumeric(strvalue) Then
           dblTotal(i) = dblTotal(i) + Val(strvalue)   'array/declare public
            If iCount = .RecordCount Then
              If i = dr_fld Then
                  drAmt = dblTotal(i)
              ElseIf i = cr_fld Then
                  crAmt = dblTotal(i)
              End If
                balAmt = (crAmt - drAmt)
                CheckBalance = balAmt
          End If  '*icount
        End If 'isnumeric
       End If 'not isnull
    Next i
.MoveNext
Wend
End With
End Function



Private Sub Balance_Total(ByRef Lvw As ListView, ByRef srcRS As Recordset, dr_fld As Integer, cr_fld As Integer) 'As String
'// which index value to compute Total
Dim strvalue As String
Dim iCount As Integer
Dim i As Integer
Dim drAmt As Double
Dim crAmt As Double
Dim balAmt As Double
Dim NumOfFields As Integer
Dim X As ListItem
'//initialize value
NumOfFields = (srcRS.Fields.Count - 1)
ReDim dblTotal(NumOfFields) As Double
For i = 0 To NumOfFields
   dblTotal(i) = 0
   Next i
iCount = 0
With srcRS
   Set X = Lvw.ListItems.Add(, , "Balance-:> ")
       X.ForeColor = vbRed
  .MoveFirst
  .MoveFirst
If srcRS.EOF = True Or srcRS.BOF = True Then Exit Sub
While Not .EOF = True
iCount = iCount + 1
  For i = 0 To NumOfFields
       strvalue = Empty
      If Not IsNull(srcRS.Fields(i)) Then
        strvalue = srcRS.Fields(i)
        If IsNumeric(strvalue) Then
           dblTotal(i) = dblTotal(i) + Val(strvalue)   'array/declare public
            If iCount = .RecordCount Then
              If i = dr_fld Then
                  drAmt = dblTotal(i)
                  X.SubItems(4) = Format(drAmt, "standard")
                  X.ListSubItems(4).ForeColor = vbBlue
              ElseIf i = cr_fld Then
                  crAmt = dblTotal(i)
                  X.SubItems(5) = Format(crAmt, "standard")
                  X.ListSubItems(5).ForeColor = vbBlue
              End If
              balAmt = (crAmt - drAmt)
              lblBalance.Caption = Format(balAmt, "standard")
               X.SubItems(6) = Format(balAmt, "standard")
               X.ListSubItems(6).ForeColor = vbRed
          End If  '*icount
        End If 'isnumeric
       End If 'not isnull
    Next i
.MoveNext
Wend
End With
End Sub

Private Sub clearText()
Dim i As Integer
For i = txtEntry.LBound To txtEntry.UBound
        txtEntry(i).Text = ""
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
           With X.ListSubItems.Add(, , txtEntry(i).Text)
           End With
          Next i
End Sub


Private Sub txtEntry_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case Index
  Case Is = 1
    If KeyCode = 113 Then
       PicTrans.Visible = True
       ListView1.SetFocus
    End If
End Select
End Sub

Private Sub txtEntry_LostFocus(Index As Integer)
 If Index = 1 Then lblF2.Visible = False
End Sub
