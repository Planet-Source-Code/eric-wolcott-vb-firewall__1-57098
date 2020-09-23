VERSION 5.00
Begin VB.UserControl UserControl3 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ControlContainer=   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4800
End
Attribute VB_Name = "UserControl3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim Red, Green, Blue
Public Enum enumOrientation
    Orientation_Horizontal = 0
    Orientation_Vertical = 1
End Enum

Public Function Gradient(Orientation As enumOrientation, SClr As ColorConstants, EClr As ColorConstants)
UserControl.AutoRedraw = True: UserControl.ScaleMode = 3 '2 is interesting,too
Analyze (SClr): SRed = Red: SGreen = Green: SBlue = Blue
Analyze (EClr): ERed = Red: EGreen = Green: EBlue = Blue
DifR = ERed - SRed: DifG = EGreen - SGreen: DifB = EBlue - SBlue
Select Case Orientation
  Case Is = 0: Fora = UserControl.ScaleHeight
  Case Is = 1: Fora = UserControl.ScaleWidth
End Select
For Yi = 0 To Fora
SRed = SRed + (DifR / Fora): If SRed < 0 Then SRed = 0
SGreen = SGreen + (DifG / Fora): If SGreen < 0 Then SGreen = 0
SBlue = SBlue + (DifB / Fora): If SBlue < 0 Then SBlue = 0
Select Case Orientation
  Case Is = 0: UserControl.Line (0, Yi)-(UserControl.ScaleWidth, Yi), RGB(SRed, SGreen, SBlue), B
  Case Is = 1: UserControl.Line (Yi, 0)-(Yi, UserControl.ScaleHeight), RGB(SRed, SGreen, SBlue), B
End Select
Next
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

Private Sub UserControl_Resize()
Gradient 0, 14078416, vbWhite
End Sub

Private Sub UserControl_Show()
Gradient 0, 14078416, vbWhite
End Sub
