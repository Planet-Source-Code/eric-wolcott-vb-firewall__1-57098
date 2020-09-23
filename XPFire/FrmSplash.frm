VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmSplash 
   BackColor       =   &H00D6D1D0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3885
   ClientLeft      =   5835
   ClientTop       =   4215
   ClientWidth     =   6375
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "FrmSplash.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   6375
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   5745
      Top             =   1080
   End
   Begin MSComCtl2.Animation Animation1 
      Height          =   825
      Left            =   1245
      TabIndex        =   0
      Top             =   1260
      Width           =   4035
      _ExtentX        =   7117
      _ExtentY        =   1455
      _Version        =   393216
      Center          =   -1  'True
      BackColor       =   14078416
      FullWidth       =   269
      FullHeight      =   55
   End
   Begin Firewall.UserControl3 UserControl31 
      Height          =   1650
      Left            =   0
      TabIndex        =   3
      Top             =   2235
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   2910
      Begin VB.Label LblLoad 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Loading..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1980
         TabIndex        =   4
         Top             =   1410
         Width           =   4365
      End
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Visual Basic"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   975
      TabIndex        =   5
      Top             =   435
      Width           =   1755
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "2005"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5160
      TabIndex        =   2
      Top             =   660
      Width           =   810
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Personal Firewall"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00CD663F&
      Height          =   990
      Left            =   1005
      TabIndex        =   1
      Top             =   615
      Width           =   4290
   End
End
Attribute VB_Name = "FrmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetTickCount& Lib "kernel32" ()
Public Start_Seconds
Function LoadAmination()
ret& = GetTickCount&
Start_Seconds = Int(ret& / 1000)
    
Dim TempData
Dim fso As New FileSystemObject
'If fso.FileExists(GetAppPath & "103.avi") = False Then
    TempData = BuildFileFromResource(GetAppPath & "103.avi", 171, "AVI")
        If TempData <> "" Then
            If fso.FileExists(TempData) = True Then
            Animation1.Open TempData
            Animation1.Play
            End If
        End If
'Else
'Animation1.Open GetAppPath & "103.avi"
'Animation1.Play
'End If

DoEvents


End Function
Private Sub Form_Load()
LoadAmination
End Sub

Private Sub LoadUP()
Dim fso As New FileSystemObject
LblLoad.Caption = "Loading... GUI"
Load FrmMain
DOUNTIL 1
LblLoad.Caption = "Loading... Checking Connections"
DOUNTIL 2
LblLoad.Caption = IIf(IsNetConnectOnline, "Connections Online", "Connections OffLine")
DOUNTIL 3
LblLoad.Caption = "Loading... Program List"
FrmMain.UpdatePrograms
DOUNTIL 4
LblLoad.Caption = IIf(IsNetConnectViaProxy, "Reading... Proxy: Found", "Reading... Proxy: Not Found")
DOUNTIL 5
FrmMain.UserControl21(0).Visible = True
FrmMain.Show
FrmMain.RefreshList
DoEvents
Unload Me
End Sub

Private Sub Timer1_Timer()
LoadUP
End Sub

Function DOUNTIL(Index As Long)
ret& = GetTickCount&
Do Until (ret& / 1000) - Start_Seconds > Index
ret& = GetTickCount&
DoEvents
Loop
End Function
