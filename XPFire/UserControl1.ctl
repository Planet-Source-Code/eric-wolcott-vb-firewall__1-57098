VERSION 5.00
Begin VB.UserControl UserControl1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ControlContainer=   -1  'True
   FillColor       =   &H00404040&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00404040&
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin Firewall.TrackMouse TrackMouse1 
      Left            =   1455
      Top             =   2400
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00D6D1D0&
      BorderStyle     =   0  'None
      Height          =   780
      Left            =   720
      ScaleHeight     =   52
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   91
      TabIndex        =   0
      Top             =   660
      Width           =   1365
   End
End
Attribute VB_Name = "UserControl1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private Declare Function DrawTextEx Lib "user32" Alias "DrawTextExA" (ByVal hDC As Long, ByVal lpsz As String, ByVal n As Long, lpRect As RECT, ByVal un As Long, lpDrawTextParams As DRAWTEXTPARAMS) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long

Private Const DT_BOTTOM = &H8
Private Const DT_CALCRECT = &H400
Private Const DT_LEFT = &H0
Private Const DT_CENTER = &H1
Private Const DT_RIGHT = &H2
Private Const DT_SINGLELINE = &H20
Private Const DT_TABSTOP = &H80
Private Const DT_TOP = &H0
Const RDW_INVALIDATE = &H1
Private Const DT_VCENTER = &H4
Private Const DT_WORDBREAK = &H10
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long

End Type
Private Type DRAWTEXTPARAMS
    cbSize As Long
    iTabLength As Long
    iLeftMargin As Long
    iRightMargin As Long
    uiLengthDrawn As Long

End Type

Public Color_Cap As String
Public Color_Cent As String

Public Color_Cap1 As String
Public Color_Cap2 As String
Public Color_Cent1 As String

Public CurrentState As Integer
Private strCaption As String

Public Event Clicked(State As Integer)

Public Property Get Caption() As String
    Caption = strCaption
End Property

Public Property Let Caption(strCaptions As String)
    strCaption = strCaptions
End Property

Function LoadColors()
Color_Cap = "22,8618883,8618883,8618883,8618883,8618883,8618883,8618883,8618883,8618883,8618883,8618883,8618883,8618883,8618883,8618883,8618883,8618883,8618883,8618883,8618883,8618883,7368302,9079176"
Color_Cent = "22,8618883,16383228,16120314,15792119,15463669,15135218,14741231,14412780,13953001,13559014,13164770,12770783,12376540,11982553,11588566,11194579,10866128,10537678,10209227,9946569,8618883,5000011,7368302"

Color_Cap1 = "24,8421504,8421504,8421504,8421504,8421504,8421504,8421504,8421504,8421504,8421504,8421504,8421504,8421504,8421504,8421504,8421504,8421504,8421504,8421504,8421504,8421504,8421504,8421504,8421504,8421504,8421504,15326939,15326939,15326939,15326939,15326939,15326939,15326939,15326939,15326939,15326939,15326939,15326939,15326939,15326939,15326939,15326939,15326939,15326939,15326939,15326939,15326939,15326939,15326939,15523804,8421504,15326939,8618883,8618883,8618883,8618883,8618883,8618883,8618883,8618883,8618883,8618883,8618883,8618883,8618883,8618883,8618883,8618883,8618883,8618883,8618883,8618883,8618883,15326939,15523804"
Color_Cap2 = "24,8421504,15326939,8618883,8618883,8618883,8618883,8618883,8618883,8618883,8618883,8618883,8618883,8618883,8618883,8618883,8618883,8618883,8618883,8618883,8618883,8618883,8618883,8618883,15326939,15523804,8421504,15326939,15326939,15326939,15326939,15326939,15326939,15326939,15326939,15326939,15326939,15326939,15326939,15326939,15326939,15326939,15326939,15326939,15326939,15326939,15326939,15326939,15326939,15326939,15523804,8421504,8421504,8421504,8421504,8421504,8421504,8421504,8421504,8421504,8421504,8421504,8421504,8421504,8421504,8421504,8421504,8421504,8421504,8421504,8421504,8421504,8421504,8421504,8421504,8421504"
Color_Cent1 = "24,8421504,15326939,8618883,15793405,15661819,15530233,15398647,15332598,15201268,15069682,14872303,14740717,14543338,14411752,14280166,14082787,14016994,13885408,13753822,13622236,13490650,13359064,8618883,15326939,15523804"
End Function

Private Function LoadBmpMenuLines(Legnth As Integer, ColorPallet As String, x As Integer, y As Integer) As Integer
    Dim Colors() As String, CurrentRow, CurrentColumn, Count, Rows
    Colors = Split(ColorPallet, ",")
    Rows = Int(Split(ColorPallet, ",")(0))
    For Count = 1 To UBound(Colors)
    If CurrentRow > (Rows) Then CurrentRow = 0: CurrentColumn = CurrentColumn + 1
    If Colors(Count) <> -1 Then
    Picture1.Line (x + CurrentColumn, y + CurrentRow)-(x + CurrentColumn + Legnth, y + CurrentRow), Colors(Count)
    End If
    CurrentRow = CurrentRow + 1
    Next
    LoadBmpMenuLines = CurrentColumn
End Function


Private Sub TrackMouse1_MouseLeftDown()
Dim z As Integer
Picture1.Cls
z = LoadBmpMenuLines(1, Color_Cap1, 0, 0)
z = LoadBmpMenuLines(UserControl.ScaleWidth - (z * 2) - 1, Color_Cent1, z + 1, 0)
LoadBmpMenuLines 1, Color_Cap2, UserControl.ScaleWidth - (z * 2) - 3, 0
Dim dword As String, dlen As Long
dword = "  " & strCaption
dlen = Len(dword)
Picture1.ForeColor = &H404040
DrawTextTohWnd dword, dlen
If CurrentState <> 0 Then
CurrentState = 0
Else
CurrentState = 1
End If
End Sub

Private Sub TrackMouse1_MouseLeftUp()
RaiseEvent Clicked(CurrentState)
End Sub

Private Sub TrackMouse1_MouseOut()
LoadColors
Picture1.Cls
If CurrentState = 0 Then
LoadBmpMenuLines 1, Color_Cap, 2, 2
LoadBmpMenuLines UserControl.ScaleWidth - 5, Color_Cent, 3, 2
LoadBmpMenuLines 1, Color_Cap, UserControl.ScaleWidth - 3, 2
Else
Dim z As Integer
z = LoadBmpMenuLines(1, Color_Cap1, 0, 0)
z = LoadBmpMenuLines(UserControl.ScaleWidth - (z * 2) - 1, Color_Cent1, z + 1, 0)
LoadBmpMenuLines 1, Color_Cap2, UserControl.ScaleWidth - (z * 2) - 3, 0
End If

Dim dword As String, dlen As Long
dword = "  " & strCaption
dlen = Len(dword)
Picture1.ForeColor = &H404040
DrawTextTohWnd dword, dlen
End Sub

Private Sub TrackMouse1_MouseOver()
LoadColors
Picture1.Cls
If CurrentState = 0 Then
LoadBmpMenuLines 1, Color_Cap, 2, 2
LoadBmpMenuLines UserControl.ScaleWidth - 5, Color_Cent, 3, 2
LoadBmpMenuLines 1, Color_Cap, UserControl.ScaleWidth - 3, 2
Else
Dim z As Integer
z = LoadBmpMenuLines(1, Color_Cap1, 0, 0)
z = LoadBmpMenuLines(UserControl.ScaleWidth - (z * 2) - 1, Color_Cent1, z + 1, 0)
LoadBmpMenuLines 1, Color_Cap2, UserControl.ScaleWidth - (z * 2) - 3, 0
End If

Dim dword As String, dlen As Long
dword = "  " & strCaption
dlen = Len(dword)
Picture1.ForeColor = vbRed
DrawTextTohWnd dword, dlen
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
strCaption = PropBag.ReadProperty("strCaption", UserControl.Name)
End Sub

Private Sub UserControl_Resize()
Picture1.Width = UserControl.ScaleWidth
Picture1.Height = UserControl.ScaleHeight
Picture1.Top = 0
Picture1.Left = 0
Picture1.Cls
LoadColors
LoadBmpMenuLines 1, Color_Cap, 2, 2
LoadBmpMenuLines UserControl.ScaleWidth - 5, Color_Cent, 3, 2
LoadBmpMenuLines 1, Color_Cap, UserControl.ScaleWidth - 3, 2
UserControl.Height = 375
End Sub

Private Sub UserControl_Show()
Picture1.Width = UserControl.ScaleWidth
Picture1.Height = UserControl.ScaleHeight
Picture1.Top = 0
Picture1.Left = 0
Picture1.Cls
'CurrentState = 0
LoadColors
If CurrentState = 0 Then
LoadBmpMenuLines 1, Color_Cap, 2, 2
LoadBmpMenuLines UserControl.ScaleWidth - 5, Color_Cent, 3, 2
LoadBmpMenuLines 1, Color_Cap, UserControl.ScaleWidth - 3, 2
Else
z = LoadBmpMenuLines(1, Color_Cap1, 0, 0)
z = LoadBmpMenuLines(UserControl.ScaleWidth - (z * 2) - 1, Color_Cent1, z + 1, 0)
LoadBmpMenuLines 1, Color_Cap2, UserControl.ScaleWidth - (z * 2) - 3, 0
End If
Dim dword As String, dlen As Long
dword = "  " & strCaption
dlen = Len(dword)
Picture1.ForeColor = &H404040
DrawTextTohWnd dword, dlen
End Sub

Private Sub DrawTextTohWnd(htext As String, lentext As Long)
    Dim vh As Integer
    Dim hrect As RECT
    SetRect hrect, 4, 0, ScaleWidth - 4, ScaleHeight
    vh = DrawText(Picture1.hDC, htext, lentext, hrect, DT_CALCRECT Or DT_CENTER Or DT_WORDBREAK)
    SetRect hrect, 4, (ScaleHeight * 0.5) - (vh * 0.5), ScaleWidth - 4, (ScaleHeight * 0.5) + (vh * 0.5)
    DrawText Picture1.hDC, htext, lentext, hrect, DT_LEFT Or DT_WORDBREAK
    RedrawWindow hwnd, hrect, ByVal 0&, RDW_INVALIDATE
End Sub

Function SubClassMe()
TrackMouse1.Watch Picture1
End Function

Function UnSubClassMe()
TrackMouse1.CloseWatch
End Function

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
PropBag.WriteProperty "strCaption", strCaption
End Sub

Function Reset()
Picture1.Cls
CurrentState = 0
LoadBmpMenuLines 1, Color_Cap, 2, 2
LoadBmpMenuLines UserControl.ScaleWidth - 5, Color_Cent, 3, 2
LoadBmpMenuLines 1, Color_Cap, UserControl.ScaleWidth - 3, 2
Dim dword As String, dlen As Long
dword = "  " & strCaption
dlen = Len(dword)
Picture1.ForeColor = &H404040
DrawTextTohWnd dword, dlen
End Function

Function ForceClick()
CurrentState = 1
Picture1.Cls
Dim z As Integer
z = LoadBmpMenuLines(1, Color_Cap1, 0, 0)
z = LoadBmpMenuLines(UserControl.ScaleWidth - (z * 2) - 1, Color_Cent1, z + 1, 0)
LoadBmpMenuLines 1, Color_Cap2, UserControl.ScaleWidth - (z * 2) - 3, 0
Dim dword As String, dlen As Long
dword = "  " & strCaption
dlen = Len(dword)
Picture1.ForeColor = &H404040
DrawTextTohWnd dword, dlen
End Function
