VERSION 5.00
Begin VB.Form frmAttempt 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Warning!"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4635
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   352
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   309
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1905
      Left            =   60
      TabIndex        =   8
      Top             =   1650
      Width           =   4515
      Begin VB.Label Label3 
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
         Height          =   345
         Index           =   2
         Left            =   45
         TabIndex        =   11
         Top             =   120
         Width           =   3900
      End
      Begin VB.Label Label3 
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
         Height          =   345
         Index           =   3
         Left            =   45
         TabIndex        =   10
         Top             =   315
         Width           =   3900
      End
      Begin VB.Label Label3 
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
         Height          =   420
         Index           =   4
         Left            =   45
         TabIndex        =   9
         Top             =   495
         Width           =   4425
      End
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   45
      TabIndex        =   6
      Text            =   "Only Block This Program Once"
      Top             =   4410
      Width           =   4545
   End
   Begin Firewall.UserControl7 UserControl71 
      Height          =   360
      Left            =   3330
      TabIndex        =   5
      Top             =   4785
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   635
      Hold_Caption    =   "Continue"
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   90
      ScaleHeight     =   34
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   36
      TabIndex        =   2
      Top             =   930
      Width           =   540
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Please Select An Action:"
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
      Height          =   240
      Left            =   60
      TabIndex        =   7
      Top             =   4185
      Width           =   2490
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Path: "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   1
      Left            =   690
      TabIndex        =   4
      Top             =   1155
      Width           =   3900
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Program: "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   690
      TabIndex        =   3
      Top             =   945
      Width           =   2760
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Warning! An unrecognized program is attempting to connect."
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
      Left            =   60
      TabIndex        =   1
      Top             =   585
      Width           =   4470
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Warning!"
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
      Left            =   270
      TabIndex        =   0
      Top             =   90
      Width           =   2595
   End
End
Attribute VB_Name = "frmAttempt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public xPath As String
Public xIndex As Integer
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40
Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

Private Function LoadBmpMenuLines(Legnth As Integer, ColorPallet As String, x As Integer, y As Integer) As Integer
    Dim Colors() As String, CurrentRow, CurrentColumn, Count, Rows
    Colors = Split(ColorPallet, ",")
    Rows = Int(Split(ColorPallet, ",")(0))
    For Count = 1 To UBound(Colors)
    If CurrentRow > (Rows) Then CurrentRow = 0: CurrentColumn = CurrentColumn + 1
    If Colors(Count) <> -1 Then
    Me.Line (x + CurrentColumn, y + CurrentRow)-(x + CurrentColumn + Legnth, y + CurrentRow), Colors(Count)
    End If
    CurrentRow = CurrentRow + 1
    Next
    LoadBmpMenuLines = CurrentColumn
End Function

Private Sub Form_Load()
Dim Color_Cent As String
Color_Cent = "36,9598839,10480895,10218495,9890559,9562623,9103615,8775679,8381951,7922943,7463679,6939135,6414335,5889791,5299455,4774655,4184319,3659775,3134975,2675710,2150909,1691388,1166331,969210,772088,509430,377588,246003,114417,113903,113389,112875,112361,111847,111333,110818,4342338,5592405"
LoadBmpMenuLines Me.ScaleWidth, Color_Cent, 0, 0

Combo1.AddItem "Only Block This Program Once"
Combo1.AddItem "Only Allow This Program Once"
Combo1.AddItem "Always Block This Program"
Combo1.AddItem "Always Allow This Program"

UserControl71.SubClassMe

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
UserControl71.UnSubClassMe
FrmMain.CurrentProcessing = Replace(FrmMain.CurrentProcessing, Chr(1) & xPath & Chr(1), "")
ResumeThreads Connection(xIndex).ProcessID
End Sub

Private Sub UserControl71_Clicked()
Select Case Combo1.ListIndex
Case -1
    TerminateThisConnection xIndex + 0
Case 0
    TerminateThisConnection xIndex + 0
Case 1
    ''
Case 2
    TerminateThisConnection xIndex + 0
    FrmMain.AddProgram xPath, 0
Case 3
    FrmMain.AddProgram xPath, 1
End Select
FrmMain.UpdatePrograms
UserControl71.UnSubClassMe
FrmMain.CurrentProcessing = Replace(FrmMain.CurrentProcessing, Chr(1) & xPath & Chr(1), "")
ResumeThreads Connection(xIndex).ProcessID
Unload Me
End Sub

Function ShowInfo(ProgramPath As String, intConnection As Integer)
xPath = ProgramPath
xIndex = intConnection
Dim FileNameShort
FileNameShort = Right(ProgramPath, Len(ProgramPath) - InStrRev(ProgramPath, "\"))

Label3(0).Caption = "Program: " & FileNameShort
Label3(1).Caption = "Path: " & ProgramPath

Label3(2).Caption = "Local Port: " & Connection(intConnection).LocalPort
Label3(3).Caption = "Remote Port: " & Connection(intConnection).RemotePort
Label3(4).Caption = "Remote Host: " & GetIPAddress(Connection(intConnection).RemoteHost) & " (" & FrmMain.iphDNS.CheckDictionary(GetIPAddress(Connection(intConnection).RemoteHost)) & ")"
GetLargeIcon ProgramPath
Me.Visible = True
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End Function

Private Function GetLargeIcon(FileName As String) As Long
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
    r = ImageList_Draw(hLIcon, ShInfo.iIcon, Picture1.hDC, 0, 0, ILD_TRANSPARENT)
    .Refresh
  End With
  
    Else

End If

End Function

