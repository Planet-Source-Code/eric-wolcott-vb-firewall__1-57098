VERSION 5.00
Begin VB.UserControl UserControl2 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   4995
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8175
   ControlContainer=   -1  'True
   MaskColor       =   &H00FFFFFF&
   ScaleHeight     =   333
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   545
   Begin VB.Image Image1 
      Height          =   4995
      Left            =   0
      Picture         =   "UserControl2.ctx":0000
      Top             =   0
      Width           =   330
   End
End
Attribute VB_Name = "UserControl2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Color_Bottom     As String
Public Color_Top        As String
Dim Red, Green, Blue
Public Enum enumOrientation2
    Orientation_Horizontal = 0
    Orientation_Vertical = 1
End Enum

Public Function Gradient(EClr As ColorConstants)
Dim Orientation As enumOrientation2
Dim SClr As ColorConstants
SClr = vbWhite
Orientation = 1
UserControl.AutoRedraw = True: UserControl.ScaleMode = 3 '2 is interesting,too
Analyze (SClr): SRed = Red: SGreen = Green: SBlue = Blue
Analyze (EClr): ERed = Red: EGreen = Green: EBlue = Blue
DifR = ERed - SRed: DifG = EGreen - SGreen: DifB = EBlue - SBlue
Select Case Orientation
  Case Is = 0: Fora = UserControl.ScaleHeight / 2
  Case Is = 1: Fora = UserControl.ScaleWidth / 2
End Select
For Yi = Fora To Fora * 2
SRed = SRed + (DifR / Fora): If SRed < 0 Then SRed = 0
SGreen = SGreen + (DifG / Fora): If SGreen < 0 Then SGreen = 0
SBlue = SBlue + (DifB / Fora): If SBlue < 0 Then SBlue = 0
Select Case Orientation
  Case Is = 0: UserControl.Line (0, Yi)-(UserControl.ScaleWidth, Yi), RGB(SRed, SGreen, SBlue), B
  Case Is = 1: UserControl.Line (Yi, 0)-(Yi, UserControl.ScaleHeight), RGB(SRed, SGreen, SBlue), B
End Select
Next

LoadBMP
LoadBmpMenuLines UserControl.ScaleWidth, Color_Bottom, 5, Image1.Height - 4
LoadBmpMenuLines UserControl.ScaleWidth, Color_Top, 5, Image1.Top


End Function

Public Function Analyze(CConst As ColorConstants)
Dim rr, gr, br As Long
rr = 1: gr = 256: br = 65536
Dim rest As Long
rest = CConst \ br
Blue = rest
CConst = CConst Mod br
If Blue < 0 Then Blue = 0
rest = CConst \ gr
Green = rest
CConst = CConst Mod gr
If Green < 0 Then Green = 0
rest = CConst \ rr
Red = rest
CConst = CConst Mod rr
If Red < 0 Then Red = 0
End Function


Function LoadBMP()
    Color_Bottom = "3,5723991,7828853,8421504,11316139"
    Color_Top = "0,13158600"
End Function

Private Function LoadBmpMenuLines(Legnth As Integer, ColorPallet As String, x As Integer, Y As Integer) As Integer
    Dim Colors() As String, CurrentRow, CurrentColumn, Count, Rows
    Colors = Split(ColorPallet, ",")
    Rows = Int(Split(ColorPallet, ",")(0))
    For Count = 1 To UBound(Colors)
    If CurrentRow > (Rows) Then CurrentRow = 0: CurrentColumn = CurrentColumn + 1
    If Colors(Count) <> -1 Then
    UserControl.Line (x + CurrentColumn, Y + CurrentRow)-(x + CurrentColumn + Legnth, Y + CurrentRow), Colors(Count)
    End If
    CurrentRow = CurrentRow + 1
    Next
    LoadBmpMenuLines = CurrentColumn
End Function

Private Sub UserControl_Resize()
LoadBMP
LoadBmpMenuLines UserControl.ScaleWidth, Color_Bottom, 5, Image1.Height - 4
LoadBmpMenuLines UserControl.ScaleWidth, Color_Top, 5, Image1.Top

End Sub

Private Sub UserControl_Show()
LoadBMP
LoadBmpMenuLines UserControl.ScaleWidth, Color_Bottom, 5, Image1.Height - 4
LoadBmpMenuLines UserControl.ScaleWidth, Color_Top, 5, Image1.Top

End Sub
