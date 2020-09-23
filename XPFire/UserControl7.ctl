VERSION 5.00
Begin VB.UserControl UserControl7 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   180
   ScaleMode       =   2  'Point
   ScaleWidth      =   240
   Begin Firewall.TrackMouse TrackMouse1 
      Left            =   2715
      Top             =   1800
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   540
      ScaleHeight     =   40
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   129
      TabIndex        =   0
      Top             =   420
      Width           =   1935
      Begin VB.Shape Shape1 
         BorderColor     =   &H00404040&
         BorderStyle     =   3  'Dot
         Height          =   165
         Left            =   60
         Top             =   90
         Visible         =   0   'False
         Width           =   600
      End
   End
End
Attribute VB_Name = "UserControl7"
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

Dim Color_Left As String
Dim Color_Center As String
Dim Color_Right As String

Dim Color_Left1 As String
Dim Color_Center1 As String
Dim Color_Right1 As String

Dim Color_Left2 As String
Dim Color_Center2 As String
Dim Color_Right2 As String

Dim Color_Left3 As String
Dim Color_Center3 As String
Dim Color_Right3 As String

Dim ButtonDown As Integer
Dim X_Cord As Integer
Dim Y_Cord As Integer

Public Event Clicked()

Dim Hold_Caption As String
Dim Hold_Enabled As Boolean
Public Property Get Caption() As String
    Caption = Hold_Caption
    Call UserControl_Show
End Property

Public Property Let Caption(strCaptions As String)
    Hold_Caption = strCaptions
    Call UserControl_Show
End Property
Public Property Get Enabled() As Boolean
    Enabled = Hold_Enabled
    Call UserControl_Show
End Property

Public Property Let Enabled(strEnabled As Boolean)
    Hold_Enabled = strEnabled
    Call UserControl_Show
End Property
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

Function LoadColors()
Color_Left = "23,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,12435134,12435134,12435134,12435134,12435134,12435134,12435134,12435134,12435134,12435134,12435134,12435134,12435134,12435134,12435134,12435134,12435134,12435134,12435134,12435134,12435134,12435134,-1,-1,12435134,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,12435134,-1"

Color_Center = "23,-1,12435134,16777215,16777215,16777215,16448250,16448249,16185334,16185334,15922162,15724783,15395562,15066854,14737889,14474716,14079958,13816530,13553613,13158857,12961477,12764099,12500928,12435134,-1"

Color_Right = "23,-1,12435134,12435134,12435134,12435134,12435134,12435134,12435134,12435134,12435134,12435134,12435134,12435134,12435134,12435134,12435134,12435134,12435134,12435134,12435134,12435134,12435134,12435134,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1"
'''''''''''''''''''''
Color_Left1 = "23,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,10856102,10856102,10856102,10856102,10856102,10856102,10856102,10856102,10856102,10856102,10856102,10856102,10856102,10856102,10856102,10856102,10856102,10856102,10856102,10856102,10856102,10856102,-1,-1,10856102,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,10856102,-1"

Color_Center1 = "23,-1,10856102,16777215,16777215,16777215,16777215,16777215,16777215,16448250,16448250,16185334,15725040,15000805,14211545,13487821,13027271,12632513,12237755,11777204,11448239,11250860,10987688,10856102,-1"

Color_Right1 = "23,-1,10856102,10856102,10856102,10856102,10856102,10856102,10856102,10856102,10856102,10856102,10856102,10856102,10856102,10856102,10856102,10856102,10856102,10856102,10856102,10856102,10856102,10856102,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1"
'''''''''''''''''''''
Color_Left2 = "23,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,12435134,12435134,12435134,12435134,12435134,12435134,12435134,12435134,12435134,12435134,12435134,12435134,12435134,12435134,12435134,12435134,12435134,12435134,12435134,12435134,12435134,12435134,-1"

Color_Center2 = "23,-1,12435134,12632513,13553358,14211545,14737889,14737889,14737889,14737889,14737889,14737889,14737889,14737889,14737889,14737889,14737889,14737889,14737889,14737889,14737889,14737889,14737889,12435134,-1"

Color_Right2 = "23,-1,12435134,12435134,12435134,12435134,12435134,12435134,12435134,12435134,12435134,12435134,12435134,12435134,12435134,12435134,12435134,12435134,12435134,12435134,12435134,12435134,12435134,12435134,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1"
'''''''''''''''''''''
Color_Left3 = "23,12632256,12632256,12632256,12632256,12632256,12632256,12632256,12632256,12632256,12632256,12632256,12632256,12632256,12632256,12632256,12632256,12632256,12632256,12632256,12632256,12632256,12632256,12632256,12632256,12632256,14343131,14343131,14343131,14343131,14343131,14343131,14343131,14343131,14343131,14343131,14343131,14343131,14343131,14343131,14343131,14343131,14343131,14343131,14343131,14343131,14343131,14343131,12632256"

Color_Center3 = "23,12632256,14145495,15922162,15922162,15922162,15922162,15922162,15922162,15922162,15922162,15922162,15922162,15922162,15922162,15922162,15922162,15922162,15922162,15922162,15922162,15922162,15922162,15132648,12632256"

Color_Right3 = "23,12632256,15132648,15132648,15132648,15132648,15132648,15132648,15132648,15132648,15132648,15132648,15132648,15132648,15132648,15132648,15132648,15132648,15132648,15132648,15132648,15132648,15132648,15132648,12632256,12632256,12632256,12632256,12632256,12632256,12632256,12632256,12632256,12632256,12632256,12632256,12632256,12632256,12632256,12632256,12632256,12632256,12632256,12632256,12632256,12632256,12632256,12632256,12632256"
End Function



Private Sub TrackMouse1_MouseLeftDown()
ButtonDown = -1
Picture1.Cls
LoadColors
Dim x As Integer, y
If Hold_Enabled = True Then
Picture1.ForeColor = vbBlack
x = LoadBmpMenuLines(1, Color_Left2, 0, 0)
y = LoadBmpMenuLines(Picture1.ScaleWidth - x - 2, Color_Center2, x, 0)
x = LoadBmpMenuLines(1, Color_Right2, Picture1.ScaleWidth - 1, 0)
X_Cord = 1
Y_Cord = 1
Shape1.Top = 3
Shape1.Left = 3
Shape1.Width = Picture1.ScaleWidth - 7
Shape1.Height = Picture1.ScaleHeight - 6
Shape1.Visible = True
Else
ShowDisable
End If

DrawTextTohWnd Hold_Caption

End Sub

Private Sub TrackMouse1_MouseLeftUp()
ButtonDown = 0
Picture1.Cls
LoadColors
Dim x As Integer, y
If Hold_Enabled = True Then
Picture1.ForeColor = vbBlack
x = LoadBmpMenuLines(1, Color_Left, 0, 0)
y = LoadBmpMenuLines(Picture1.ScaleWidth - x - 2, Color_Center, x, 0)
x = LoadBmpMenuLines(1, Color_Right, Picture1.ScaleWidth - x, 0)
X_Cord = 0
Y_Cord = 0
RaiseEvent Clicked
Else
ShowDisable
End If
DrawTextTohWnd Hold_Caption
Shape1.Visible = False

End Sub

Private Sub TrackMouse1_MouseOut()
If ButtonDown = -1 Then Exit Sub
Picture1.Cls
LoadColors
Dim x As Integer, y
If Hold_Enabled = True Then
Picture1.ForeColor = vbBlack
x = LoadBmpMenuLines(1, Color_Left, 0, 0)
y = LoadBmpMenuLines(Picture1.ScaleWidth - x - 2, Color_Center, x, 0)
x = LoadBmpMenuLines(1, Color_Right, Picture1.ScaleWidth - x, 0)
Else
ShowDisable
End If
DrawTextTohWnd Hold_Caption
End Sub

Private Sub TrackMouse1_MouseOver()
If ButtonDown = -1 Then Exit Sub
Picture1.Cls
LoadColors

Dim x As Integer, y
If Hold_Enabled = True Then
Picture1.ForeColor = vbBlack
x = LoadBmpMenuLines(1, Color_Left1, 0, 0)
y = LoadBmpMenuLines(Picture1.ScaleWidth - x - 2, Color_Center1, x, 0)
x = LoadBmpMenuLines(1, Color_Right1, Picture1.ScaleWidth - x, 0)
Else
ShowDisable
End If
DrawTextTohWnd Hold_Caption
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Hold_Caption = PropBag.ReadProperty("Hold_Caption", "null")
Hold_Enabled = PropBag.ReadProperty("Hold_Enabled", True)
End Sub

Private Sub UserControl_Show()
Picture1.Cls
ButtonDown = 0
LoadColors
Picture1.Top = 0
Picture1.Left = 0
Picture1.Width = UserControl.ScaleWidth
Picture1.Height = 18
UserControl.Height = 24 * 15

If Hold_Enabled = True Then
Dim x As Integer, y
Picture1.ForeColor = vbBlack
x = LoadBmpMenuLines(1, Color_Left, 0, 0)
y = LoadBmpMenuLines(Picture1.ScaleWidth - x - 2, Color_Center, x, 0)
x = LoadBmpMenuLines(1, Color_Right, Picture1.ScaleWidth - x, 0)
Else
ShowDisable
End If
DrawTextTohWnd Hold_Caption

End Sub

Function ShowDisable()
Dim x As Integer, y
x = LoadBmpMenuLines(1, Color_Left3, 0, 0)
y = LoadBmpMenuLines(Picture1.ScaleWidth - x - 2, Color_Center3, x, 0)
x = LoadBmpMenuLines(1, Color_Right3, Picture1.ScaleWidth - 2, 0)
X_Cord = 2
Y_Cord = 2
Picture1.ForeColor = vbWhite

DrawTextTohWnd Hold_Caption

Picture1.ForeColor = &HC0C0C0
X_Cord = 0
Y_Cord = 0
End Function

Function SubClassMe()
TrackMouse1.Watch Picture1
End Function

Function UnSubClassMe()
TrackMouse1.CloseWatch
End Function

Private Sub DrawTextTohWnd(htext2 As String)
    Dim lentext As Long
    Dim vh As Integer
    Dim hrect As RECT
    Dim htext As String
    htext = String(0, " ") & htext2
    lentext = Len(htext)
    SetRect hrect, 4, 0, Picture1.ScaleWidth - 4, Picture1.ScaleHeight
    vh = DrawText(Picture1.hDC, htext, lentext, hrect, DT_CALCRECT Or DT_CENTER Or DT_WORDBREAK)
    SetRect hrect, X_Cord + 4, Y_Cord + (Picture1.ScaleHeight * 0.5) - (vh * 0.5), Picture1.ScaleWidth - 4, (Picture1.ScaleHeight * 0.5) + (vh * 0.5)
    DrawText Picture1.hDC, htext, lentext, hrect, DT_CENTER Or DT_WORDBREAK
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
PropBag.WriteProperty "Hold_Caption", Hold_Caption, "null"
PropBag.WriteProperty "Hold_Enabled", Hold_Enabled, True
End Sub
