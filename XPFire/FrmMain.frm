VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmMain 
   BackColor       =   &H00D6D1D0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "VB Personal Firewall"
   ClientHeight    =   6795
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   9285
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   9285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Firewall.UserControl6 UserControl61 
      Align           =   1  'Align Top
      Height          =   555
      Left            =   0
      TabIndex        =   43
      Top             =   0
      Width           =   9285
      _ExtentX        =   16378
      _ExtentY        =   979
   End
   Begin Firewall.UserControl2 UserControl21 
      Height          =   4995
      Index           =   2
      Left            =   2760
      TabIndex        =   12
      Top             =   570
      Visible         =   0   'False
      Width           =   6585
      _ExtentX        =   11615
      _ExtentY        =   8811
      Begin Firewall.UserControl7 UserControl74 
         Height          =   360
         Left            =   4575
         TabIndex        =   22
         Top             =   3975
         Width           =   1920
         _ExtentX        =   3387
         _ExtentY        =   635
         Hold_Caption    =   "Block"
      End
      Begin Firewall.UserControl7 UserControl73 
         Height          =   360
         Left            =   4575
         TabIndex        =   21
         Top             =   4425
         Width           =   1920
         _ExtentX        =   3387
         _ExtentY        =   635
         Hold_Caption    =   "Allow"
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   4365
         Left            =   255
         TabIndex        =   20
         Top             =   435
         Width           =   4245
         _ExtentX        =   7488
         _ExtentY        =   7699
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Application Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Location"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Status"
            Object.Width           =   2364
         EndProperty
      End
      Begin Firewall.UserControl7 UserControl75 
         Height          =   360
         Left            =   4575
         TabIndex        =   50
         Top             =   450
         Width           =   1920
         _ExtentX        =   3387
         _ExtentY        =   635
         Hold_Caption    =   "Delete"
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         X1              =   255
         X2              =   5805
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Personal Firewall Program List"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   360
         Left            =   255
         TabIndex        =   13
         Top             =   30
         Width           =   5310
      End
   End
   Begin Firewall.UserControl2 UserControl21 
      Height          =   4995
      Index           =   1
      Left            =   2775
      TabIndex        =   10
      Top             =   660
      Visible         =   0   'False
      Width           =   6585
      _ExtentX        =   11615
      _ExtentY        =   8811
      Begin Firewall.Status Status2 
         Height          =   300
         Left            =   270
         TabIndex        =   44
         Top             =   705
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   529
      End
      Begin Firewall.Status Status1 
         Height          =   300
         Left            =   270
         TabIndex        =   42
         Top             =   1020
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   529
      End
      Begin VB.Timer Timer4 
         Interval        =   1
         Left            =   2760
         Top             =   3120
      End
      Begin Firewall.ConStatus ConStatus1 
         Height          =   735
         Left            =   180
         TabIndex        =   37
         Top             =   1665
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   1296
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Connection: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1005
         TabIndex        =   47
         Top             =   2115
         Width           =   1725
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Label6"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   630
         TabIndex        =   45
         Top             =   750
         Width           =   2490
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Internet Connection"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   41
         Top             =   1425
         Width           =   1965
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "................................................................."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   2130
         TabIndex        =   40
         Top             =   1455
         Width           =   4215
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Received: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1005
         TabIndex        =   39
         Top             =   1905
         Width           =   1725
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Sent: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1005
         TabIndex        =   38
         Top             =   1695
         Width           =   1725
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Label6"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   630
         TabIndex        =   19
         Top             =   1065
         Width           =   2490
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "........................................................................................."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   1770
         TabIndex        =   18
         Top             =   510
         Width           =   4215
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Personal Firewall"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Personal Firewall Status And Settings"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   360
         Left            =   255
         TabIndex        =   11
         Top             =   30
         Width           =   5310
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         X1              =   255
         X2              =   5805
         Y1              =   360
         Y2              =   360
      End
   End
   Begin Firewall.UserControl4 UserControl42 
      Height          =   1320
      Left            =   45
      TabIndex        =   6
      Top             =   2685
      Visible         =   0   'False
      Width           =   2625
      _ExtentX        =   4630
      _ExtentY        =   2328
   End
   Begin Firewall.UserControl4 UserControl41 
      Height          =   945
      Left            =   45
      TabIndex        =   5
      Top             =   1125
      Visible         =   0   'False
      Width           =   2610
      _ExtentX        =   4604
      _ExtentY        =   1667
   End
   Begin Firewall.UserControl1 UserControl12 
      Height          =   375
      Left            =   15
      TabIndex        =   4
      Top             =   2295
      Width           =   2685
      _ExtentX        =   4736
      _ExtentY        =   661
      strCaption      =   "Help"
   End
   Begin Firewall.UserControl1 UserControl11 
      Height          =   375
      Left            =   15
      TabIndex        =   3
      Top             =   735
      Width           =   2685
      _ExtentX        =   4736
      _ExtentY        =   661
      strCaption      =   "Firewall"
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2340
      Left            =   0
      ScaleHeight     =   2340
      ScaleWidth      =   9300
      TabIndex        =   2
      Top             =   5550
      Width           =   9300
      Begin VB.PictureBox PicQuestion 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   7125
         Picture         =   "FrmMain.frx":6852
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   46
         Top             =   1800
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   435
         Left            =   7860
         ScaleHeight     =   435
         ScaleWidth      =   1410
         TabIndex        =   33
         Top             =   90
         Visible         =   0   'False
         Width           =   1410
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Monitoring"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   345
            TabIndex        =   34
            Top             =   90
            Width           =   1575
         End
         Begin VB.Image Image1 
            Height          =   345
            Left            =   0
            Picture         =   "FrmMain.frx":6B94
            Top             =   0
            Width           =   330
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Height          =   1230
         Left            =   8805
         ScaleHeight     =   1230
         ScaleWidth      =   9315
         TabIndex        =   31
         Top             =   285
         Visible         =   0   'False
         Width           =   9315
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Please Wait..."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2070
            TabIndex        =   32
            Top             =   150
            Width           =   5475
         End
      End
      Begin VB.PictureBox Pic32 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   75
         ScaleHeight     =   615
         ScaleWidth      =   675
         TabIndex        =   9
         Top             =   75
         Width           =   675
      End
      Begin VB.Timer Timer1 
         Interval        =   200
         Left            =   5310
         Top             =   1785
      End
      Begin VB.Timer Timer2 
         Interval        =   300
         Left            =   6585
         Top             =   1785
      End
      Begin VB.PictureBox Pic16 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   6390
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   8
         Top             =   1395
         Visible         =   0   'False
         Width           =   240
      End
      Begin MSComctlLib.ImageList TreeViewImgList 
         Left            =   5835
         Top             =   1695
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   17
         ImageHeight     =   18
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   13
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmMain.frx":71F2
               Key             =   "Cancel"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmMain.frx":7250
               Key             =   "OpenFolder"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmMain.frx":72AE
               Key             =   "Ping"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmMain.frx":730C
               Key             =   "FileCross"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmMain.frx":736A
               Key             =   "CancelCon"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmMain.frx":73C8
               Key             =   "Cross"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmMain.frx":7426
               Key             =   "Excal"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmMain.frx":7484
               Key             =   "FileNet"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmMain.frx":74E2
               Key             =   "QuestionComp"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmMain.frx":7540
               Key             =   "NotConnected"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmMain.frx":759E
               Key             =   "Connected"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmMain.frx":75FC
               Key             =   "HelpFile"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmMain.frx":765A
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   2505
         Top             =   4770
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   17
         ImageHeight     =   18
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   12
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmMain.frx":76B8
               Key             =   "Cancel"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmMain.frx":7716
               Key             =   "OpenFolder"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmMain.frx":7774
               Key             =   "Ping"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmMain.frx":77D2
               Key             =   "FileCross"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmMain.frx":7830
               Key             =   "CancelCon"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmMain.frx":788E
               Key             =   "Cross"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmMain.frx":78EC
               Key             =   "Excal"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmMain.frx":794A
               Key             =   "FileNet"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmMain.frx":79A8
               Key             =   "QuestionComp"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmMain.frx":7A06
               Key             =   "NotConnected"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmMain.frx":7A64
               Key             =   "Connected"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmMain.frx":7AC2
               Key             =   "HelpFile"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageList3 
         Left            =   2445
         Top             =   1590
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmMain.frx":7B20
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmMain.frx":7B7E
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmMain.frx":7BDC
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmMain.frx":7C3A
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Version: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   8
         Left            =   2385
         TabIndex        =   35
         Top             =   420
         Width           =   1545
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Attempts: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   7
         Left            =   825
         TabIndex        =   30
         Top             =   1020
         Width           =   2835
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Remote Port: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   6
         Left            =   2385
         TabIndex        =   29
         Top             =   840
         Width           =   2835
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Local Port: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   5
         Left            =   825
         TabIndex        =   28
         Top             =   810
         Width           =   2835
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Remote Host: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   2385
         TabIndex        =   27
         Top             =   615
         Width           =   6390
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "PID: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   825
         TabIndex        =   26
         Top             =   615
         Width           =   2835
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Size: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   825
         TabIndex        =   25
         Top             =   420
         Width           =   1545
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Location: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   825
         TabIndex        =   24
         Top             =   240
         Width           =   5790
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Name: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   825
         TabIndex        =   23
         Top             =   60
         Width           =   2835
      End
      Begin ComctlLib.ImageList Iml16 
         Left            =   3945
         Top             =   1620
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         MaskColor       =   12632256
         _Version        =   327682
      End
      Begin ComctlLib.ImageList ImageList2 
         Left            =   3225
         Top             =   1590
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         MaskColor       =   12632256
         _Version        =   327682
      End
   End
   Begin Firewall.UserControl3 UserControl31 
      Height          =   2940
      Left            =   0
      TabIndex        =   1
      Top             =   2610
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   5186
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Program Info:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   75
         TabIndex        =   36
         Top             =   2730
         Width           =   1455
      End
   End
   Begin Firewall.UserControl2 UserControl21 
      Height          =   4995
      Index           =   0
      Left            =   2715
      TabIndex        =   0
      Top             =   555
      Width           =   6585
      _ExtentX        =   11615
      _ExtentY        =   8811
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   465
         Left            =   690
         ScaleHeight     =   435
         ScaleWidth      =   4890
         TabIndex        =   48
         Top             =   945
         Visible         =   0   'False
         Width           =   4920
         Begin VB.Label Label17 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "No Internet Connection Found"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   60
            TabIndex        =   49
            Top             =   45
            Width           =   4740
         End
      End
      Begin Firewall.UserControl7 UserControl72 
         Height          =   360
         Left            =   4920
         TabIndex        =   16
         Top             =   4440
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   635
         Hold_Caption    =   "Close Connection"
         Hold_Enabled    =   0   'False
      End
      Begin Firewall.UserControl7 UserControl71 
         Height          =   360
         Left            =   3360
         TabIndex        =   15
         Top             =   4440
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   635
         Hold_Caption    =   "Close Program"
         Hold_Enabled    =   0   'False
      End
      Begin ComctlLib.ListView ListView1 
         Height          =   3375
         Left            =   240
         TabIndex        =   14
         Top             =   405
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   5953
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         _Version        =   327682
         Icons           =   "Iml16"
         SmallIcons      =   "ImageList2"
         ForeColor       =   -2147483640
         BackColor       =   16777215
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Text            =   ""
            Object.Width           =   7832
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   1
            Key             =   ""
            Object.Tag             =   ""
            Text            =   ""
            Object.Width           =   7832
         EndProperty
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Personal Firewall Monitoring"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   360
         Left            =   255
         TabIndex        =   7
         Top             =   30
         Width           =   5310
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   255
         X2              =   5805
         Y1              =   360
         Y2              =   360
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private m_objIpHelper As CIpHelper

Dim oOld As Long
Dim oNew As Long
Dim aOld As Long
Dim aNew As Long

Dim objInterface2 As CInterface
Dim obJHelper As CInterface
Dim tValue As Long
Dim aValue As Long

Public Unloaded As Boolean
Private Processing As Boolean
Private IsOnline As Boolean

Private TVHost As Long
Private TVPath As String
Private TVTAG As Long
Private TVPI As Long

Public iphDNS As New CDictionary

Public Pub_BlockAll As Boolean

Dim xfrmAttempt As New frmAttempt

Public CurrentProcessing As String
Public Firewall_Enabled As Boolean
Public HideMe As Integer

Private Sub Exit_Click()
Unload Me
End Sub

Public Sub RefreshList()
  Dim i
  Dim Item As ListItem
If IsOnline = False Then Exit Sub
If Unloaded = True Then Exit Sub
Processing = True
    RefreshStack
    DoEvents
    LoadNTProcess
    DoEvents
ListView1.ListItems.Clear
ListView1.Sorted = False
ListView1.ColumnHeaders(1).Width = 2000
ListView1.ColumnHeaders(2).Width = ListView1.Width - 2000 - 600
SetTrayTip "Personal Firewall: Monitoring " & GetEntryCount & " Connections"
For i = 0 To GetEntryCount
        If Connection(i).State = "2" Then GoTo IsListening
            If Connection(i).FileName = "" Then
                Set Item = ListView1.ListItems.Add(, , "Unknown")
            Else
                Dim FileNameShort
                FileNameShort = Right(Connection(i).FileName, Len(Connection(i).FileName) - InStrRev(Connection(i).FileName, "\"))
                Set Item = ListView1.ListItems.Add(, , FileNameShort & " (" & GetPort(Connection(i).LocalPort) & ")")
            End If
            Item.Tag = i
IsListening:
Next i

ListView1.Sorted = True
GetAllIcons
DoEvents
ShowIcons
DoEvents
resolveIPs False
DoEvents
Finished:
Processing = False
If Unloaded = True Then Unload Me
End Sub

Private Sub resolveIPs(ShowHost As Boolean)
Dim Item As ListItem
    

For Each Item In ListView1.ListItems
If ShowHost = False Then
Item.SubItems(1) = GetIPAddress(Connection(Item.Tag).RemoteHost) & ":" & Connection(Item.Tag).RemotePort
Else
Item.SubItems(1) = iphDNS.CheckDictionary(GetIPAddress(Connection(Item.Tag).RemoteHost)) & ":" & Connection(Item.Tag).RemotePort
End If
DoEvents
Next

End Sub

Private Function GetIcon(FileName As String, Index As Long) As Long
'---------------------------------------------------------------------
'Extract an individual icon
'---------------------------------------------------------------------
On Error Resume Next
Dim hLIcon As Long, hSIcon As Long    'Large & Small Icons
Dim imgObj As ListImage               'Single bmp in imagelist.listimages collection
Dim r As Long


If Connection(ListView1.ListItems(Index).Tag).FileName = "" Then
Set imgObj = Iml16.ListImages.Add(Index, , PicQuestion.Image)
Exit Function
End If


'Get a handle to the small icon
hSIcon = SHGetFileInfo(FileName, 0&, ShInfo, Len(ShInfo), _
         BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
'Get a handle to the large icon
'hLIcon = SHGetFileInfo(FileName, 0&, ShInfo, Len(ShInfo), _
'         BASIC_SHGFI_FLAGS Or SHGFI_LARGEICON)

'If the handle(s) exists, load it into the picture box(es)
If hLIcon <> 0 Then

  'Large Icon
  'With Pic32
  '  Set .Picture = LoadPicture("")
  '  .AutoRedraw = True
  '  r = ImageList_Draw(hLIcon, ShInfo.iIcon, Pic32.hDC, 0, 0, ILD_TRANSPARENT)
  '  .Refresh
  'End With
  
    Else
  'Small Icon
  With Pic16
    Set .Picture = LoadPicture("")
    .AutoRedraw = True
    r = ImageList_Draw(hSIcon, ShInfo.iIcon, Pic16.hDC, 0, 0, ILD_TRANSPARENT)
    .Refresh
  End With
  
  Set imgObj = Iml16.ListImages.Add(Index, , Pic16.Image)
End If

End Function

Private Function GetLargeIcon(FileName As String) As Long
'---------------------------------------------------------------------
'Extract an individual icon
'---------------------------------------------------------------------
On Error Resume Next
Dim hLIcon As Long, hSIcon As Long    'Large & Small Icons
Dim imgObj As ListImage               'Single bmp in imagelist.listimages collection
Dim r As Long


If FileName = "" Then
'Set imgObj = Iml16.ListImages.Add(Index, , PicQuestion.Image)
Exit Function
End If


'Get a handle to the large icon
hLIcon = SHGetFileInfo(FileName, 0&, ShInfo, Len(ShInfo), _
         BASIC_SHGFI_FLAGS Or SHGFI_LARGEICON)

'If the handle(s) exists, load it into the picture box(es)
If hLIcon <> 0 Then

  'Large Icon
  With Pic32
    Set .Picture = LoadPicture("")
    .AutoRedraw = True
    r = ImageList_Draw(hLIcon, ShInfo.iIcon, Pic32.hDC, 0, 0, ILD_TRANSPARENT)
    .Refresh
  End With
  
    Else

End If

End Function

Private Sub ShowIcons()
'-----------------------------------------
'Show the icons in the lvw
'-----------------------------------------
On Error Resume Next

Dim Item As ListItem
With ListView1
  '.ListItems.Clear
  .SmallIcons = Iml16   'Small
  For Each Item In .ListItems
    Item.SmallIcon = Item.Index
  Next
End With

End Sub

Private Sub GetAllIcons()
'--------------------------------------------------
'Extract all icons
'--------------------------------------------------
Dim Item As ListItem
Dim FileName As String

    ListView1.SmallIcons = Nothing
    Iml16.ListImages.Clear
    
'On Local Error Resume Next
For Each Item In ListView1.ListItems
  FileName = Connection(Item.Tag).FileName

  GetIcon FileName, Item.Index
   
Next

End Sub

Private Sub Form_Load()
HideMe = 1
Load frmSystemTray
Firewall_Enabled = True
Pub_BlockAll = False
Set m_objIpHelper = New CIpHelper
Dim FP As FILE_PARAMS
Dim CurFile As Long
Dim AppPath As String
Dim fso As New FileSystemObject
    
If IsNetConnectOnline() = True Then
    Timer2.Enabled = True
    IsOnline = True
    Else
    ListView1.ListItems.Clear
    Timer2.Enabled = False
    IsOnline = False
End If

    
If Right(App.Path, 1) <> "\" Then AppPath = App.Path & "\" & App.EXEName & ".exe" Else AppPath = App.Path & App.EXEName & ".exe"

TVPath = AppPath

GetLargeIcon AppPath

   With FP
      .sFileNameExt = AppPath
   End With
   
CurFile = GetFileInformation(FP)

'Animation.Open App.Path & "\xpsearchinternet.avi"
'Animation.AutoPlay = True
Me.BackColor = 14078416
UserControl11.SubClassMe
UserControl41.AddButton "Monitoring"
UserControl41.AddButton "Status And Settings"
UserControl41.AddButton "Applications"
UserControl41.AddButton "Configure Ports"
UserControl12.SubClassMe
UserControl42.AddButton "Status And Setings"
UserControl42.AddButton "Statistics"
UserControl42.AddButton "Applications"
UserControl12.Top = UserControl11.Top + UserControl11.Height + 5
UserControl61.SubClassMe
UserControl71.SubClassMe
UserControl72.SubClassMe
UserControl73.SubClassMe
UserControl74.SubClassMe
UserControl75.SubClassMe
UserControl12.Reset
UserControl42.Reset
UserControl42.Visible = False

UserControl41.Left = UserControl11.Left
UserControl41.Top = UserControl11.Top + UserControl11.Height
UserControl41.Width = UserControl11.Width
UserControl41.ShowMe
UserControl41.Visible = True
UserControl12.Top = UserControl41.Top + UserControl41.Height + 5

UserControl41.ShowMenu 1
UserControl11.ForceClick

UpdatePrograms

UserControl21(0).Visible = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
If Processing = True Or HideMe = 1 Then
'Unloaded = True
Cancel = -1
Me.Visible = False
Exit Sub
End If

iphDNS.WriteCache

UserControl11.UnSubClassMe
UserControl41.UnSubClass
UserControl12.UnSubClassMe
UserControl42.UnSubClass
UserControl61.UnSubClassMe
UserControl71.UnSubClassMe
UserControl72.UnSubClassMe
UserControl73.UnSubClassMe
UserControl74.UnSubClassMe
UserControl75.UnSubClassMe
DoEvents
End Sub

Private Sub Form_Resize()
On Error Resume Next
ListView1.ColumnHeaders(1).Width = 1300
ListView1.ColumnHeaders(2).Width = 1100
ListView1.ColumnHeaders(4).Width = 1100
ListView1.ColumnHeaders(5).Width = 1100
ListView1.ColumnHeaders(6).Width = ListView1.Width \ 2 + 1000

End Sub

Private Sub ListView1_GotFocus()
UserControl71.Enabled = True
UserControl72.Enabled = True
End Sub

Private Sub ListView1_ItemClick(ByVal Item As ComctlLib.ListItem)
Picture2.Top = 0
Picture2.Left = 0
Picture2.BackColor = vbWhite
Picture2.Visible = True
DoEvents
Dim FP As FILE_PARAMS
Dim CurFile As Long

TVHost = Connection(ListView1.ListItems(Item.Index).Tag).RemoteHost
TVPath = Connection(ListView1.ListItems(Item.Index).Tag).FileName
TVTAG = ListView1.ListItems(Item.Index).Tag
TVPI = Connection(ListView1.ListItems(Item.Index).Tag).ProcessID
Label7(1).Caption = "Path: " & TVPath

Label7(3).Caption = "PID: " & TVPI
Label7(4).Caption = "Remote Host: " & iphDNS.CheckDictionary(GetIPAddress(Connection(Item.Tag).RemoteHost)) & " (" & GetIPAddress(TVHost) & ")"
Label7(5).Caption = "Local Port: " & Connection(ListView1.ListItems(Item.Index).Tag).LocalPort
Label7(6).Caption = "Remote Port: " & Connection(ListView1.ListItems(Item.Index).Tag).RemotePort

    Dim FileNameShort
    FileNameShort = Right(Connection(TVTAG).FileName, Len(Connection(TVTAG).FileName) - InStrRev(Connection(TVTAG).FileName, "\"))
Label7(0).Caption = "Name: " & FileNameShort

Dim xc
xc = CheckProgramID(TVPath)
If xc <> -1 Then
Picture3.Visible = True
Label7(7).Caption = "Attempts: " & Program(xc).Attempts
Label7(7).Visible = True
Else
Label7(7).Visible = False
Picture3.Visible = False
End If

GetLargeIcon (TVPath)

   With FP
      .sFileNameExt = TVPath
   End With
   
CurFile = GetFileInformation(FP)

DoEvents
'If ResolveHostchk.Value = 0 Then lblHost.Caption = "Remote Host : " & GetHostNameFromIP(GetIPAddress(TVHost)) Else lblHost.Caption = "Remote Host : " & GetIPAddress(TVHost)

'PopulateTreeview (Item.Index)
'item click

Picture2.Visible = False
End Sub

Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then
If ListView1.SelectedItem.Selected = False Then Exit Sub

End If
End Sub

Private Sub Timer1_Timer()
NotOnline (IsNetConnectOnline())
End Sub

Public Sub NotOnline(Online As Boolean)
If Online = False Then
    IsOnline = False
    Picture4.Visible = True
    Exit Sub
End If
If Online = True Then
    IsOnline = True
    Picture4.Visible = False
End If
CallRefresh:
If GetRefresh = True Then RefreshList
End Sub
Private Sub UpdateInterfaceInfo()
Dim objInterface        As CInterface
Static st_objInterface  As CInterface
Static lngBytesRecv     As Long
Static lngBytesSent     As Long
Dim blnIsRecv           As Boolean
Dim blnIsSent           As Boolean
If st_objInterface Is Nothing Then Set st_objInterface = New CInterface
Set objInterface = m_objIpHelper.Interfaces(1)
Select Case objInterface.InterfaceType
Case MIB_IF_TYPE_ETHERNET: Label16.Caption = "Connection: Ethernet"
Case MIB_IF_TYPE_FDDI: Label16.Caption = "Connection: FDDI"
Case MIB_IF_TYPE_LOOPBACK: Label16.Caption = "Connection: Loopback"
Case MIB_IF_TYPE_OTHER: Label16.Caption = "Connection: Other"
Case MIB_IF_TYPE_PPP: Label16.Caption = "Connection: PPP"
Case MIB_IF_TYPE_SLIP: Label16.Caption = "Connection: SLIP"
Case MIB_IF_TYPE_TOKENRING: Label16.Caption = "Connection: TokenRing"
End Select

If ShowTrafficInBytes = False Then
    Label10.Caption = "Received: " & GiveByteValues(Trim(Format(m_objIpHelper.BytesReceived, "###,###,###,###")))
    Label11.Caption = "Sent: " & GiveByteValues(Trim(Format(m_objIpHelper.BytesSent, "###,###,###,###")))
Else
    Label10.Caption = "Received: " & Trim(Format(m_objIpHelper.BytesReceived, "###,###,###,###"))
    Label11.Caption = "Sent: " & Trim(Format(m_objIpHelper.BytesSent, "###,###,###,###"))
End If
  '
    blnIsRecv = (m_objIpHelper.BytesReceived > lngBytesRecv)
    blnIsSent = (m_objIpHelper.BytesSent > lngBytesSent)
    '
    If blnIsRecv And blnIsSent Then
        ConStatus1.SetStatus 0
    ElseIf (Not blnIsRecv) And blnIsSent Then
        ConStatus1.SetStatus 3
    ElseIf blnIsRecv And (Not blnIsSent) Then
        ConStatus1.SetStatus 2
    ElseIf Not (blnIsRecv And blnIsSent) Then
        ConStatus1.SetStatus 1
    End If
    '
    lngBytesRecv = m_objIpHelper.BytesReceived
    lngBytesSent = m_objIpHelper.BytesSent
    '

    Set st_objInterface = objInterface

End Sub

Private Sub Timer2_Timer()
Call UpdateInterfaceInfo
End Sub

Private Function GetFileNameFromPath(ByVal sFullPath As String) As String
   Dim WFD As WIN32_FIND_DATA
   Dim hFile As Long
         If sFullPath = "" Then
         GetFileNameFromPath = "Unknown"
         Exit Function
         End If
   hFile = FindFirstFile(sFullPath, WFD)
   If hFile <> INVALID_HANDLE_VALUE Then
      GetFileNameFromPath = TrimNull(WFD.cFileName)
      Call FindClose(hFile)
   End If
End Function


Private Function TrimNull(startstr As String) As String
   Dim pos As Integer
   pos = InStr(startstr, Chr$(0))
   If pos Then
      TrimNull = Left$(startstr, pos - 1)
      Exit Function
   End If
   TrimNull = startstr
End Function

Public Function PingIP(IP As String)
Dim ECHO As ICMP_ECHO_REPLY
   Dim pos As Long
   Dim success As Long
   Dim sIPAddress As String
   If SocketsInitialize() Then
      sIPAddress = IP
      success = Ping(sIPAddress, "Echo This", ECHO)
      If GetStatusCode(success) = "ip success" Then PingIP = "Success - Round Time : " & ECHO.RoundTripTime & " ms" Else PingIP = GetStatusCode(success)
     
      If Left$(ECHO.Data, 1) <> Chr$(0) Then
         pos = InStr(ECHO.Data, Chr$(0))
         'Left$(ECHO.Data, pos - 1)
      End If
      SocketsCleanup
   Else
        MsgBox "Windows Sockets for 32 bit Windows " & _
               "is not successfully responding.", vbInformation, "Error"
   End If
End Function

Private Function GetFileInformation(FP As FILE_PARAMS) As Long

   Dim WFD As WIN32_FIND_DATA
   Dim hFile As Long
   Dim nSize As Long
   Dim sPath As String
   Dim sRoot As String
   Dim sTmp As String
   Dim itmx As ListItem
   Dim lv As Control

   sPath = FP.sFileNameExt
   hFile = FindFirstFile(sPath, WFD)
   If hFile <> INVALID_HANDLE_VALUE Then
         sTmp = TrimNull(WFD.cFileName)
         If Not (WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) _
            = FILE_ATTRIBUTE_DIRECTORY Then
            nSize = nSize + (WFD.nFileSizeHigh * (MAXDWORD + 1)) + WFD.nFileSizeLow
               'Set itmx = lv.ListItems.Add(, , LCase$(sTmp))
               Label7(8).Caption = "Version: " & GetFileVersion(sRoot & sTmp)
               Label7(2).Caption = "Size: " & GetFileSizeStr(WFD.nFileSizeHigh + WFD.nFileSizeLow)
               'itmx.SubItems(2) = GetFileDescription(sRoot & sTmp)
               'itmx.SubItems(4) = LCase$(sRoot)
                'If GetFileDescription(sPath) = "" Then Lblinfo.Caption = "Description : (No Description) " Else Lblinfo.Caption = "Description : " & GetFileDescription(sPath)
                'If GetFileVersion(sPath) = "" Then LblVersion.Caption = "Version : (No Version) " Else LblVersion.Caption = "Version : " & GetFileVersion(sPath)
                'LblSize.Caption = "Size : " & GetFileSizeStr(WFD.nFileSizeHigh + WFD.nFileSizeLow)
         End If
      hFile = FindClose(hFile)
   End If
   GetFileInformation = nSize
End Function


Private Function GetFileSizeStr(fsize As Long) As String
    GetFileSizeStr = GiveByteValues(Format$((fsize), "###,###,###"))  '& " kb"
End Function

Public Function GetTempDir() As String
   Dim tmp As String
   tmp = Space$(256)
   Call GetTempPath(Len(tmp), tmp)
   GetTempDir = TrimNull(tmp)
End Function

Public Function BasePath(ByVal fname As String, Optional delim As String = "\", Optional keeplast As Boolean = True) As String
    Dim outstr As String
    Dim llen As Long
    llen = InStrRev(fname, delim)
    If (Not keeplast) Then
        llen = llen - 1
    End If
    If (llen > 0) Then
        BasePath = Mid(fname, 1, llen)
    Else
        BasePath = fname
    End If
End Function


Private Sub Timer4_Timer()
If Pub_BlockAll = True Then
Status1.SetStatus 1
Label6.Caption = "Block All Programs: On"
Else
Status1.SetStatus 0
Label6.Caption = "Block All Programs: Off"
End If

If Firewall_Enabled = True Then
Status2.SetStatus 1
Label15.Caption = "Firewall Enabled"
Else
Status2.SetStatus 0
Label15.Caption = "Firewall Disabled"
End If
End Sub

Private Sub UserControl11_Clicked(State As Integer)
UserControl12.Reset
UserControl42.Reset
UserControl42.Visible = False
Select Case State
Case 0
UserControl41.Visible = False
UserControl12.Top = UserControl11.Top + UserControl11.Height + 5
Case 1
UserControl41.Left = UserControl11.Left
UserControl41.Top = UserControl11.Top + UserControl11.Height
UserControl41.Width = UserControl11.Width
UserControl41.ShowMe
UserControl41.Visible = True
UserControl12.Top = UserControl41.Top + UserControl41.Height + 5
End Select
End Sub

Private Sub UserControl12_Clicked(State As Integer)
UserControl11.Reset
UserControl41.Reset
UserControl41.Visible = False
UserControl12.Top = UserControl11.Top + UserControl11.Height + 5
Select Case State
Case 0
UserControl42.Visible = False
Case 1
UserControl42.Left = UserControl12.Left
UserControl42.Top = UserControl12.Top + UserControl12.Height
UserControl42.Width = UserControl12.Width
UserControl42.ShowMe
UserControl42.Visible = True
End Select
End Sub

Private Sub UserControl41_ButtonClick(Index As Integer)
Select Case Index
Case 1
    HideFrames
    UserControl21(0).Visible = True
    'UserControl21(0).Gradient &H52F18A
Case 2
    HideFrames
    UserControl21(1).Visible = True
    UserControl21(1).Gradient &H8080FF
Case 3
    HideFrames
    UpdatePrograms
    UserControl21(2).Visible = True
    UserControl21(2).Gradient &H80FFFF
End Select
End Sub

Private Sub UserControl42_ButtonClick(Index As Integer)
Select Case Index
Case 1
UserControl21(0).Gradient &HFF8080
Case 2
UserControl21(1).Gradient &H80C0FF
Case 3
UserControl21(2).Gradient &HFF80FF
End Select
End Sub

Private Sub UserControl61_ButtonClick(Index As Integer)
Select Case Index
Case 1
    If Pub_BlockAll = False Then
    Pub_BlockAll = True
    Else
    Pub_BlockAll = False
    End If
    RefreshList
Case 2
Case 3
End Select
End Sub

Function HideFrames()
Dim x
For x = 0 To UserControl21.Count - 1
UserControl42.ZOrder 0
UserControl41.ZOrder 0
UserControl21(x).Visible = False
UserControl21(x).Top = UserControl21(0).Top
UserControl21(x).Left = UserControl21(0).Left
UserControl21(x).Width = UserControl21(0).Width
UserControl21(x).Height = UserControl21(0).Height
Next
End Function

Private Sub UserControl72_Clicked()
TerminateThisConnection ListView1.SelectedItem.Tag
End Sub

Function UpdatePrograms()
ListView2.ListItems.Clear
    Dim Item As ListItem
    Dim x, z, y(4) As String
    x = GetSetting(App.Title & "Firewall", "Programs", "ProgramCount", 0)
    For z = 0 To x
        y(0) = GetSetting(App.Title & "Firewall", "Programs", "Name" & z, "[Name Not Found]")
        y(1) = GetSetting(App.Title & "Firewall", "Programs", "Path" & z, "c:\Program Files\Internet Explorer\iexplore.exe")
        y(2) = GetSetting(App.Title & "Firewall", "Programs", "Status" & z, "0")
        y(3) = GetSetting(App.Title & "Firewall", "Programs", "Attempts" & z, "0")
        y(4) = GetSetting(App.Title & "Firewall", "Programs", "Blocks" & z, "0")
        
        ListView2.ListItems.Add , , y(0)
        ListView2.ListItems(ListView2.ListItems.Count).ListSubItems.Add , , y(1)
        
        With Program(z)
        .FileName = y(0)
        .FilePath = y(1)
            If Int(y(2)) = 0 Then
            .Block = True
            ListView2.ListItems(ListView2.ListItems.Count).ListSubItems.Add , , "Block"
            Else
            .Block = False
            ListView2.ListItems(ListView2.ListItems.Count).ListSubItems.Add , , "Allow"
            End If
            .Attempts = y(3)
        .Blocked = y(4)
        .Count = x
        End With
        ListView2.ListItems(ListView2.ListItems.Count).Tag = z
        'ListView2.ListItems(ListView2.ListItems.Count).ListSubItems.Add , , y(3)
        'ListView2.ListItems(ListView2.ListItems.Count).ListSubItems.Add , , y(4)
    Next
End Function

Function CheckPrograms(ProgramPath As String, Index As Integer) As Boolean
Dim x
CheckPrograms = False
For x = 0 To Program(0).Count
    If UCase(Program(x).FilePath) = UCase(ProgramPath) Then
        Program(x).Attempts = Program(x).Attempts + 1
        SaveSetting App.Title & "Firewall", "Programs", "Attempts" & x, Program(x).Attempts
            If Program(x).Block = True Then
                Program(x).Blocked = Program(x).Blocked + 1
                SaveSetting App.Title & "Firewall", "Programs", "Blocks" & x, Program(x).Blocked
                If Firewall_Enabled = True Then CheckPrograms = True
            End If
        Exit Function
    End If
Next
If InStr(1, CurrentProcessing, Chr(1) & ProgramPath & Chr(1)) Then Exit Function
    SuspendThreads (Connection(Index).ProcessID)
    CurrentProcessing = CurrentProcessing & Chr(1) & ProgramPath & Chr(1)
    Set xfrmAttempt = New frmAttempt
    xfrmAttempt.ShowInfo ProgramPath, Index
End Function

Function CheckProgramID(ProgramPath) As Integer
Dim x
CheckProgramID = -1
For x = 1 To Program(0).Count
    If UCase(Program(x).FilePath) = UCase(ProgramPath) Then
        CheckProgramID = x
        Exit Function
    End If
Next
End Function

Function AddProgram(ProgramPath As String, Block As Integer)
    Dim FileNameShort
    FileNameShort = Right(ProgramPath, Len(ProgramPath) - InStrRev(ProgramPath, "\"))
    MsgBox "Are you sure you want to ALLAWAYS ALLOW this " & FileNameShort & " ?", vbYesNo
    Dim Xt
    Xt = GetSetting(App.Title & "Firewall", "Programs", "ProgramCount", 0)
    Xt = Xt + 1
    SaveSetting App.Title & "Firewall", "Programs", "Name" & Xt, UCase(FileNameShort)
    SaveSetting App.Title & "Firewall", "Programs", "Path" & Xt, UCase(ProgramPath)
    SaveSetting App.Title & "Firewall", "Programs", "Status" & Xt, Block
    SaveSetting App.Title & "Firewall", "Programs", "ProgramCount", Xt
End Function

Function DeleteProgram(Index As Integer)
    Dim Xt, Xp
    Xt = GetSetting(App.Title & "Firewall", "Programs", "ProgramCount", 0)
    If Index <> Xt Then
    For Xp = Index To Xt
    DeleteSetting App.Title & "Firewall", "Programs", "Name" & Xp
    DeleteSetting App.Title & "Firewall", "Programs", "Path" & Xp
    DeleteSetting App.Title & "Firewall", "Programs", "Status" & Xp
    If Xp <> Xt Then
    SaveSetting App.Title & "Firewall", "Programs", "Name" & Xp, GetSetting(App.Title & "Firewall", "Programs", "Name" & Xp + 1)
    SaveSetting App.Title & "Firewall", "Programs", "Path" & Xp, GetSetting(App.Title & "Firewall", "Programs", "Path" & Xp + 1)
    SaveSetting App.Title & "Firewall", "Programs", "Status" & Xp, GetSetting(App.Title & "Firewall", "Programs", "Status" & Xp + 1)
    SaveSetting App.Title & "Firewall", "Programs", "Attempts" & Xp, GetSetting(App.Title & "Firewall", "Programs", "Attempts" & Xp + 1, 0)
    SaveSetting App.Title & "Firewall", "Programs", "Blocks" & Xp, GetSetting(App.Title & "Firewall", "Programs", "Blocks" & Xp + 1, 0)
    End If
    Next
    Else
    DeleteSetting App.Title & "Firewall", "Programs", "Name" & Xt
    DeleteSetting App.Title & "Firewall", "Programs", "Path" & Xt
    DeleteSetting App.Title & "Firewall", "Programs", "Status" & Xt
    End If
    Xt = Xt - 1
    SaveSetting App.Title & "Firewall", "Programs", "ProgramCount", Xt
End Function

Private Sub UserControl73_Clicked()
SaveSetting App.Title & "Firewall", "Programs", "Status" & ListView2.SelectedItem.Index - 1, 1
UpdatePrograms
UserControl21(2).Visible = True
End Sub

Private Sub UserControl74_Clicked()
SaveSetting App.Title & "Firewall", "Programs", "Status" & ListView2.SelectedItem.Index - 1, 0
UpdatePrograms
UserControl21(2).Visible = True
End Sub

Private Sub UserControl75_Clicked()
DeleteProgram ListView2.SelectedItem.Index - 1
UpdatePrograms
UserControl21(2).Visible = True
End Sub
