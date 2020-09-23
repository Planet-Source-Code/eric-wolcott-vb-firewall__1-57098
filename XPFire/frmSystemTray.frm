VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmSystemTray 
   Caption         =   "frmSystemTray"
   ClientHeight    =   3090
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.PictureBox PicIco 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   810
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   0
      Top             =   390
      Width           =   270
   End
   Begin MSComctlLib.ImageList IL1 
      Left            =   75
      Top             =   135
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSystemTray.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuSystem 
      Caption         =   "SystemMenu"
      Begin VB.Menu mnuOpen 
         Caption         =   "Open Personal Firewall"
      End
      Begin VB.Menu mnuEnable 
         Caption         =   "Disable Firewall"
      End
      Begin VB.Menu mnuBlockAll 
         Caption         =   "Block All: Turn ON"
      End
      Begin VB.Menu mnuclose 
         Caption         =   "Close"
      End
   End
End
Attribute VB_Name = "frmSystemTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
AddToTray Me, mnuSystem
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
RemoveFromTray
End Sub

Private Sub mnuBlockAll_Click()
    If Pub_BlockAll = False Then
    Pub_BlockAll = True
    mnuBlockAll.Caption = "Block All: Turn OFF"
    Else
    Pub_BlockAll = False
    mnuBlockAll.Caption = "Block All: Turn ON"
    End If
    FrmMain.RefreshList
End Sub

Private Sub mnuclose_Click()
FrmMain.Unloaded = True
FrmMain.HideMe = 0
Unload FrmMain
Unload Me
End Sub

Private Sub mnuEnable_Click()
    If FrmMain.Firewall_Enabled = False Then
    FrmMain.Firewall_Enabled = True
    mnuEnable.Caption = "Disable Firewall"
    Else
    FrmMain.Firewall_Enabled = False
    mnuEnable.Caption = "Enable Firewall"
    End If
End Sub

Private Sub mnuOpen_Click()
FrmMain.Visible = True
End Sub
