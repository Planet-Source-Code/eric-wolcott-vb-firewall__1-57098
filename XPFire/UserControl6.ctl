VERSION 5.00
Begin VB.UserControl UserControl6 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0FF&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6480
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   432
   Begin Firewall.UserControl5 UserControl53 
      Height          =   555
      Left            =   3360
      TabIndex        =   2
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   979
      Hold_Caption    =   "Options"
      Hold_Icon       =   3
   End
   Begin Firewall.UserControl5 UserControl52 
      Height          =   555
      Left            =   1680
      TabIndex        =   1
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   979
      Hold_Caption    =   "Live Update"
      Hold_Icon       =   2
   End
   Begin Firewall.UserControl5 UserControl51 
      Height          =   555
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   979
      Hold_Caption    =   "Block all"
   End
End
Attribute VB_Name = "UserControl6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public Event ButtonClick(Index As Integer)

Function SubClassMe()
UserControl51.SubClassMe
UserControl52.SubClassMe
UserControl53.SubClassMe
End Function

Function UnSubClassMe()
UserControl51.UnSubClassMe
UserControl52.UnSubClassMe
UserControl53.UnSubClassMe
End Function

Private Sub UserControl_Show()
Dim Color_Cent As String
Color_Cent = "36,9598839,10480895,10218495,9890559,9562623,9103615,8775679,8381951,7922943,7463679,6939135,6414335,5889791,5299455,4774655,4184319,3659775,3134975,2675710,2150909,1691388,1166331,969210,772088,509430,377588,246003,114417,113903,113389,112875,112361,111847,111333,110818,4342338,5592405"
LoadBmpMenuLines UserControl.ScaleWidth, Color_Cent, 0, 0
End Sub
Private Function LoadBmpMenuLines(Legnth As Integer, ColorPallet As String, x As Integer, y As Integer) As Integer
    Dim Colors() As String, CurrentRow, CurrentColumn, Count, Rows
    Colors = Split(ColorPallet, ",")
    Rows = Int(Split(ColorPallet, ",")(0))
    For Count = 1 To UBound(Colors)
    If CurrentRow > (Rows) Then CurrentRow = 0: CurrentColumn = CurrentColumn + 1
    If Colors(Count) <> -1 Then
    UserControl.Line (x + CurrentColumn, y + CurrentRow)-(x + CurrentColumn + Legnth, y + CurrentRow), Colors(Count)
    End If
    CurrentRow = CurrentRow + 1
    Next
    LoadBmpMenuLines = CurrentColumn
End Function

Private Sub UserControl51_Clicked()
RaiseEvent ButtonClick(1)
UserControl52.Reset
UserControl53.Reset
End Sub

Private Sub UserControl52_Clicked()
RaiseEvent ButtonClick(2)
UserControl51.Reset
UserControl53.Reset
End Sub

Private Sub UserControl53_Clicked()
RaiseEvent ButtonClick(3)
UserControl51.Reset
UserControl52.Reset
End Sub
