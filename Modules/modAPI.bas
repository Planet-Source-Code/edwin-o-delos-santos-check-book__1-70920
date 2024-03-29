Attribute VB_Name = "modAPI"
Option Explicit


Public Declare Function SendMessage Lib "user32" _
Alias "SendMessageA" (ByVal hWnd As Long, ByVal _
wMsg As Integer, ByVal wParam As Integer, lParam _
As Any) As Long

'<< use by autoaligncolheaders >>
Public Const LVM_SETCOLUMNWIDTH = (&H1000 + 30)
Public Const LVSCW_AUTOSIZE = -1
Public Const LVSCW_AUTOSIZE_USEHEADER = -2
'<<--------------------->>
'====================================================
'** FLAT HEADERS
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" _
   (ByVal hWnd As Long, _
    ByVal Msg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long) As Long
Public Declare Function GetWindowLong Lib "user32" _
   Alias "GetWindowLongA" _
   (ByVal hWnd As Long, _
   ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" _
   Alias "SetWindowLongA" _
   (ByVal hWnd As Long, _
   ByVal nIndex As Long, _
   ByVal dwNewLong As Long) As Long
Public Const GWL_STYLE = (-16)
Public Const LVM_FIRST = &H1000
Public Const LVM_GETHEADER = (LVM_FIRST + 31)
Public Const HDS_BUTTONS = &H2
'**====END FLAT HEADERS==================================
'[=================================]
'< Win32 Declarations for DisableX >
'[=================================]
Public Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Public Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Public Const MF_BYPOSITION = &H400&
'-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-
'_______________ always on top ____________________
'__________________________________________________
Public Declare Function SetWindowPos Lib "user32" _
(ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, _
ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const HWND_TOPMOST = -1

'//this is how ...//
'*Private Sub Form_Resize()
'*SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
'*End Sub

'_________________________________________________
'________ Drag any control________________________
'_________________________________________________

'Public Declare Function SendMessage Lib "user32" Alias "SendMessageA"  (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long

Public Const WM_NCLBUTTONDOWN = &HA1

Public Const HTCAPTION = 2
'<< syntax >>
'---------- on conctrol mousemove --------------
'    If Button = vbLeftButton Then
'        Call DragIt(picture1.hwnd)   'picture1 is the control
'    End If
'-----------------------------------------------
'//local declaration
'Private Sub DragIt(ByVal lngHwnd As Long)
'Dim lngReturn As Long
'    lngReturn = ReleaseCapture()
'    lngReturn = SendMessage(lngHwnd, WM_NCLBUTTONDOWN, HTCAPTION, CLng(0))
'End Sub
'-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-


'_________________________________________________
'________ Form Round Corner_______________________
'_________________________________________________
Global Const winding = 2
Global Const alternate = 1
Global Const rgn_or = 2

Type pointapi
   X As Long
   Y As Long
End Type

Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As pointapi, ByVal nCount As Long, ByVal nPolyfillMode As Long) As Long
Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal x3 As Long, ByVal y3 As Long) As Long
Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
'-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-

Public Sub FormRndCorner(ByRef frm As Form, _
                         ByVal wd As Long, _
                         ByVal ht As Long)
'// round corner
Dim X(2) As pointapi
Dim lRegion As Long
Dim lRegion1 As Long
Dim lRegion2 As Long
Dim lResult As Long
    frm.Width = wd * Screen.TwipsPerPixelX
    frm.Height = ht * Screen.TwipsPerPixelY

    lRegion = CreatePolygonRgn(X(0), 3, alternate)

    lRegion1 = CreatePolygonRgn(X(0), 3, alternate)
    '4=Left/2=Top/wd=Width/ht=Height/20=curve/20=curve
    lRegion2 = CreateRoundRectRgn(4, 2, wd, ht, 20, 20)
    lResult = CombineRgn(lRegion, lRegion1, lRegion2, rgn_or)
    DeleteObject lRegion1
    DeleteObject lRegion2
    lResult = SetWindowRgn(frm.hWnd, lRegion, True)
End Sub




