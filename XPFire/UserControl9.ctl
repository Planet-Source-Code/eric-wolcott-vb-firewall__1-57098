VERSION 5.00
Begin VB.UserControl Status 
   BackColor       =   &H00FFFFFF&
   BackStyle       =   0  'Transparent
   ClientHeight    =   3045
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3045
   ScaleWidth      =   4800
   ToolboxBitmap   =   "UserControl9.ctx":0000
   Begin VB.Image Image4 
      Height          =   375
      Left            =   135
      Top             =   105
      Width           =   465
   End
   Begin VB.Image Image3 
      Height          =   270
      Left            =   1515
      Picture         =   "UserControl9.ctx":0312
      Top             =   765
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Image Image2 
      Height          =   300
      Left            =   1050
      Picture         =   "UserControl9.ctx":05CE
      Top             =   735
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image Image1 
      Height          =   300
      Left            =   750
      Picture         =   "UserControl9.ctx":0A93
      Top             =   765
      Visible         =   0   'False
      Width           =   315
   End
End
Attribute VB_Name = "Status"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public pub_Status As Integer

Function SetStatus(Status As Integer)
If pub_Status = Status Then Exit Function
pub_Status = Status
Image4.Top = 0
Image4.Left = 0
Select Case Status
Case 0
Image4.Picture = Image2.Picture
Case 1
Image4.Picture = Image1.Picture
Case 2
Image4.Picture = Image3.Picture
End Select
UserControl.Height = Image4.Height
UserControl.Width = Image4.Width
End Function

Private Sub UserControl_Resize()
Image4.Top = 0
Image4.Left = 0
Image4.Picture = Image2.Picture
UserControl.Height = Image4.Height
UserControl.Width = Image4.Width
End Sub

Private Sub UserControl_Show()
Image4.Top = 0
Image4.Left = 0
Image4.Picture = Image2.Picture
UserControl.Height = Image4.Height
UserControl.Width = Image4.Width
End Sub
