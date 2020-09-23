Attribute VB_Name = "Module2"
'Tray Module
'Simple traymodule for Anim Icons
'by Scythe
'scythe@cablenet
'www.scythe-tools.de

Option Explicit

'Tray
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

'Subclass
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

'Allways on top
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const HWND_TOPMOST = -1

'Tray data & events
Private Type NOTIFYICONDATA
 cbSize As Long
 hwnd As Long
 uID As Long
 uFlags As Long
 uCallbackMessage As Long
 hIcon As Long
 szTip As String * 64
End Type

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const WM_MOUSEMOVE = &H200
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const TRAY_CALLBACK = (&H7E9)
Private Const GWL_WNDPROC = (-4)

Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205

Private Const ILD_TRANSPARENT = &H1       'Display transparent

Public Const icoLock = "11,-1,-1,-1,-1,-1,5479640,5808602,5611737,5611737,5611737,4625365,13488079,-1,-1,11579568,11776173,9149873,4296148,8834804,7455217,6208241,6933748,4300002,13027528,-1,10263708,9474192,9999505,8424859,5085401,10547198,8773630,7196926,8250366,4826858,11185067,11053224,10000536,5855577,5395026,5528159,4561114,9104126,7789054,6670590,7394558,4497642,11185067,10790052,10263708,5723733,8289919,11186615,4561114,8250366,7196926,6078462,6670590,4168938,11185067,11579568,10263708,7171436,10264222,10070713,4101596,7197694,6407166,5683454,6078462,3971306,11185067,-1,9737364,10000536,9999505,8883089,3837911,6342654,5683454,5025790,5420798,3708394,11185067,-1,15264491,7434609,5855577,5134434,2982611,6342654,5683454,5025790,5420798,3246041,11185067,-1,-1,15264491,12698306,11185067,2982611,3708133,3510499,3378915,3379173,3246041,11185067,-1,-1,-1,-1,-1,14869989,12698306,11185067,11185067,11185067,11185067,13356493"

Dim OldWindowProc As Long
Dim TrayDat As NOTIFYICONDATA
Dim TrayForm As Form
Dim TrayMenu As Menu

Public Sub AddToTray(frm As Form, mnu As Menu)
 LoadBmpMenuLines 1, icoLock, 3, 1
 frm.PicIco.Line (0, 0)-(frm.PicIco.ScaleWidth, 0), &HC0C0C0
 frm.PicIco.Line (frm.PicIco.ScaleWidth - 1, 0)-(frm.PicIco.ScaleWidth - 1, frm.PicIco.ScaleHeight), &HC0C0C0
 frm.PicIco.Line (0, 0)-(0, frm.PicIco.ScaleHeight), &HC0C0C0
 frm.PicIco.Line (0, frm.PicIco.ScaleHeight - 1)-(frm.PicIco.ScaleWidth, frm.PicIco.ScaleHeight - 1), &HC0C0C0
 
 frm.IL1.ListImages.Add 2, , frm.PicIco.Image
 'Get it back as Icon
 frm.PicIco.Picture = frm.IL1.ListImages(1).ExtractIcon
 'Delete new createt Image from list
 frm.IL1.ListImages.Remove (2)
 'Add form to tray
 
 Set TrayMenu = mnu
 Set TrayForm = frm

 'Subclass
 OldWindowProc = SetWindowLong(frm.hwnd, GWL_WNDPROC, AddressOf NewWindowProc)

 'Set the Tray Icon
 With TrayDat
 .uID = 0
 .hwnd = frm.hwnd
 .cbSize = Len(TrayDat)
 'We need a picture on the form to get the Icon from it
 .hIcon = frm.PicIco.Picture
 .uFlags = NIF_ICON
 .uCallbackMessage = TRAY_CALLBACK
 .uFlags = .uFlags Or NIF_MESSAGE Or ILD_TRANSPARENT
 .cbSize = Len(TrayDat)
 End With
 
 'DO it
 Shell_NotifyIcon NIM_ADD, TrayDat

End Sub

'Subclass function
Public Function NewWindowProc(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
 
 'We pressed any Button ?
 If Msg = TRAY_CALLBACK Then
  If lParam = WM_RBUTTONUP Or lParam = WM_LBUTTONUP Then
   'Show the hidden Menu from form
   TrayForm.PopupMenu TrayMenu
   Exit Function
  End If
 End If
 
 'Go back to old routine
 NewWindowProc = CallWindowProc(OldWindowProc, hwnd, Msg, wParam, lParam)
End Function

'Delete from tray
Public Sub RemoveFromTray()
 'remove TrayIcon
 TrayDat.uFlags = 0
 Shell_NotifyIcon NIM_DELETE, TrayDat
 'End Subclassing
 SetWindowLong frmSystemTray.hwnd, GWL_WNDPROC, OldWindowProc
End Sub

'Show the new TrayIcon
Public Sub UpdateIcon()
 TrayDat.hIcon = frmSystemTray.PicIco.Picture
 TrayDat.uFlags = NIF_ICON
 Shell_NotifyIcon NIM_MODIFY, TrayDat
End Sub
'Show the New Tooltip
Public Sub SetTrayTip(tip As String)
 TrayDat.szTip = tip & vbNullChar
 TrayDat.uFlags = NIF_TIP
 Shell_NotifyIcon NIM_MODIFY, TrayDat
End Sub

Private Function LoadBmpMenuLines(Legnth As Integer, ColorPallet As String, x As Integer, y As Integer) As Integer
    Dim Colors() As String, CurrentRow, CurrentColumn, Count, Rows
    Colors = Split(ColorPallet, ",")
    Rows = Int(Split(ColorPallet, ",")(0))
    For Count = 1 To UBound(Colors)
    If CurrentRow > (Rows) Then CurrentRow = 0: CurrentColumn = CurrentColumn + 1
    If Colors(Count) <> -1 Then
    frmSystemTray.PicIco.Line (x + CurrentColumn, y + CurrentRow)-(x + CurrentColumn + Legnth, y + CurrentRow), Colors(Count)
    End If
    CurrentRow = CurrentRow + 1
    Next
    LoadBmpMenuLines = CurrentColumn
End Function
