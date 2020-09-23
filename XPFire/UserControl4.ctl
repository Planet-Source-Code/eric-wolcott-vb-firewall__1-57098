VERSION 5.00
Begin VB.UserControl UserControl4 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0FF&
   BackStyle       =   0  'Transparent
   ClientHeight    =   2460
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2865
   ScaleHeight     =   164
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   191
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   270
      Index           =   0
      Left            =   60
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   1
      Top             =   -255
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1920
      Left            =   15
      ScaleHeight     =   128
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   191
      TabIndex        =   0
      Top             =   -165
      Width           =   2865
      Begin Firewall.TrackMouse TrackMouse1 
         Index           =   0
         Left            =   1005
         Top             =   1575
         _ExtentX        =   741
         _ExtentY        =   741
      End
   End
End
Attribute VB_Name = "UserControl4"
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

Public SelectedOption As Integer

Public Color_Cap As String

Public Color_Cent As String

Public Color_Corner As String

Public Event ButtonClick(Index As Integer)
Function ShowMe()
Picture1.Top = 0
Picture1.Left = 0
Picture1.Width = UserControl.ScaleWidth
Picture1.Height = Picture2(Picture2.Count - 1).Top + Picture2(Picture2.Count - 1).Height + 5

Dim He, Wi
He = Picture1.ScaleHeight
Wi = Picture1.ScaleWidth
Picture1.BackColor = 15523804
Picture1.Line (0, 0)-(0, He), 8421504
Picture1.Line (0, He - 1)-(Wi, He - 1), 8421504
Picture1.Line (Wi - 1, He)-(Wi - 1, -1), 8421504

UserControl.Width = UserControl.Width + 200
UserControl.Height = (Picture2(Picture2.Count - 1).Top + Picture2(Picture2.Count - 1).Height + 5) * 15
End Function

Function AddButton(Caption As String)
SelectedOption = -1
Load Picture2(Picture2.Count)
Load TrackMouse1(TrackMouse1.Count)
With Picture2(Picture2.Count - 1)
    Dim htext As String
    Dim lentext As Long
    htext = Caption
    lentext = Len(Caption)
    Dim vh As Integer
    Dim hrect As RECT
    .Top = Picture2(Picture2.Count - 2).Top + Picture2(Picture2.Count - 2).Height
    .BackColor = 15523804
    Picture1.ZOrder 1
    .Visible = True
    .Tag = Caption
    
    SetRect hrect, 4, 0, .ScaleWidth - 4, .ScaleHeight
    vh = DrawText(.hDC, htext, lentext, hrect, DT_CALCRECT Or DT_CENTER Or DT_WORDBREAK)
    SetRect hrect, 4, (.ScaleHeight * 0.5) - (vh * 0.5), .ScaleWidth - 4, (.ScaleHeight * 0.5) + (vh * 0.5)
    DrawText .hDC, htext, lentext, hrect, DT_LEFT Or DT_WORDBREAK
End With
TrackMouse1(TrackMouse1.Count - 1).Watch Picture2(Picture2.Count - 1)
End Function

Private Sub TrackMouse1_MouseLeftDown(Index As Integer)
ShowMenu Index
End Sub

Private Sub TrackMouse1_MouseLeftUp(Index As Integer)
RaiseEvent ButtonClick(Index)
End Sub

Private Sub TrackMouse1_MouseOut(Index As Integer)
If SelectedOption <> Index Then
Picture2(Index).Cls
End If
Picture2(Index).ForeColor = &H808080
Picture2(Index).FontUnderline = False

With Picture2(Index)
Dim htext As String
Dim lentext As Long
Dim vh As Integer
Dim hrect As RECT
htext = .Tag
lentext = Len(.Tag)
SetRect hrect, 4, 0, .ScaleWidth - 4, .ScaleHeight
vh = DrawText(.hDC, htext, lentext, hrect, DT_CALCRECT Or DT_CENTER Or DT_WORDBREAK)
SetRect hrect, 4, (.ScaleHeight * 0.5) - (vh * 0.5), .ScaleWidth - 4, (.ScaleHeight * 0.5) + (vh * 0.5)
DrawText .hDC, htext, lentext, hrect, DT_LEFT Or DT_WORDBREAK
End With

End Sub

Private Sub TrackMouse1_MouseOver(Index As Integer)
If SelectedOption <> Index Then
Picture2(Index).Cls
End If
Picture2(Index).ForeColor = vbBlack
Picture2(Index).FontUnderline = True

With Picture2(Index)
Dim htext As String
Dim lentext As Long
Dim vh As Integer
Dim hrect As RECT
htext = .Tag
lentext = Len(.Tag)
SetRect hrect, 4, 0, .ScaleWidth - 4, .ScaleHeight
vh = DrawText(.hDC, htext, lentext, hrect, DT_CALCRECT Or DT_CENTER Or DT_WORDBREAK)
SetRect hrect, 4, (.ScaleHeight * 0.5) - (vh * 0.5), .ScaleWidth - 4, (.ScaleHeight * 0.5) + (vh * 0.5)
DrawText .hDC, htext, lentext, hrect, DT_LEFT Or DT_WORDBREAK
End With

End Sub

Function UnSubClass()
Dim f
For f = 1 To TrackMouse1.Count - 1
TrackMouse1(f).CloseWatch
Next
End Function

Function LoadColors()
Color_Cap = "1,6052956,8421504,6052956,13289672"

Color_Cent = "1,6052956,8421504"

Color_Corner = "1,15523804,15523804,6052956,15523804"
End Function

Private Function LoadBmpMenuLines(Index As Integer, Legnth As Integer, ColorPallet As String, x As Integer, y As Integer) As Integer
    Dim Colors() As String, CurrentRow, CurrentColumn, Count, Rows
    Colors = Split(ColorPallet, ",")
    Rows = Int(Split(ColorPallet, ",")(0))
    For Count = 1 To UBound(Colors)
    If CurrentRow > (Rows) Then CurrentRow = 0: CurrentColumn = CurrentColumn + 1
    If Colors(Count) <> -1 Then
    Picture2(Index).Line (x + CurrentColumn, y + CurrentRow)-(x + CurrentColumn + Legnth, y + CurrentRow), Colors(Count)
    End If
    CurrentRow = CurrentRow + 1
    Next
    LoadBmpMenuLines = CurrentColumn
End Function

Function Reset()
If SelectedOption <> -1 Then
Picture2(SelectedOption).Cls
Picture2(SelectedOption).BackColor = 15523804
Picture2(SelectedOption).Width = Picture2(0).Width
Picture2(SelectedOption).Height = Picture2(Index).Height
Dim htext As String
Dim lentext As Long
Dim vh As Integer
Dim hrect As RECT
With Picture2(SelectedOption)
.ForeColor = &H808080
.FontUnderline = False
htext = .Tag
lentext = Len(.Tag)
SetRect hrect, 4, 0, .ScaleWidth - 4, .ScaleHeight
vh = DrawText(.hDC, htext, lentext, hrect, DT_CALCRECT Or DT_CENTER Or DT_WORDBREAK)
SetRect hrect, 4, (.ScaleHeight * 0.5) - (vh * 0.5), .ScaleWidth - 4, (.ScaleHeight * 0.5) + (vh * 0.5)
DrawText .hDC, htext, lentext, hrect, DT_LEFT Or DT_WORDBREAK
End With
SelectedOption = -1
End If
End Function

Function ShowMenu(Index As Integer)
If Picture2.Count < Index Then Exit Function
If SelectedOption = Index Then Exit Function
Dim htext As String
Dim lentext As Long
Dim vh As Integer
Dim hrect As RECT
If SelectedOption <> -1 Then
Picture2(SelectedOption).Cls
Picture2(SelectedOption).BackColor = 15523804
Picture2(SelectedOption).Width = Picture2(0).Width
Picture2(SelectedOption).Height = Picture2(Index).Height

With Picture2(SelectedOption)
.ForeColor = &H808080
.FontUnderline = False
htext = .Tag
lentext = Len(.Tag)
SetRect hrect, 4, 0, .ScaleWidth - 4, .ScaleHeight
vh = DrawText(.hDC, htext, lentext, hrect, DT_CALCRECT Or DT_CENTER Or DT_WORDBREAK)
SetRect hrect, 4, (.ScaleHeight * 0.5) - (vh * 0.5), .ScaleWidth - 4, (.ScaleHeight * 0.5) + (vh * 0.5)
DrawText .hDC, htext, lentext, hrect, DT_LEFT Or DT_WORDBREAK
End With

End If
SelectedOption = Index
LoadColors
Picture2(Index).BackColor = vbWhite
Picture2(Index).Width = UserControl.Width
Picture2(Index).Height = Picture2(Index).Height + 2
LoadBmpMenuLines Index, Picture1.ScaleWidth, Color_Cent, 0, Picture2(Index).ScaleHeight - 2
LoadBmpMenuLines Index, 1, Color_Corner, 0, Picture2(Index).ScaleHeight - 2
LoadBmpMenuLines Index, 1, Color_Cap, Picture1.ScaleWidth, Picture2(Index).ScaleHeight - 2

With Picture2(Index)
htext = .Tag
lentext = Len(.Tag)
SetRect hrect, 4, 0, .ScaleWidth - 4, .ScaleHeight
vh = DrawText(.hDC, htext, lentext, hrect, DT_CALCRECT Or DT_CENTER Or DT_WORDBREAK)
SetRect hrect, 4, (.ScaleHeight * 0.5) - (vh * 0.5), .ScaleWidth - 4, (.ScaleHeight * 0.5) + (vh * 0.5)
DrawText .hDC, htext, lentext, hrect, DT_LEFT Or DT_WORDBREAK
End With


End Function
