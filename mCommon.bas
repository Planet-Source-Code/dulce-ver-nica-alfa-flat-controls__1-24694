Attribute VB_Name = "mCommon"
Option Explicit

Private Const PS_SOLID = 0

' ..:: Public Constants ::..
Public Const GWL_STYLE = (-16)
Public Const WM_PAINT = &HF
Public Const WM_TIMER = &H113
Public Const WM_MOUSEMOVE = &H200
Public Const WM_SETFOCUS = &H7
Public Const WM_KILLFOCUS = &H8
Public Const WM_COMMAND = &H111

' ..:: Public types ::..
Public Type POINTAPI
  x As Long
  y As Long
End Type
Public Type RECT
  Left     As Long
  Top      As Long
  Right    As Long
  Bottom   As Long
End Type

' ..:: Private API's ::..
Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

' ..:: Public API's ::..
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Public Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function GetFocus Lib "user32" () As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Public Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Public Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long

Public Function TranslateColor(ByVal clr As OLE_COLOR, Optional hPal As Long = 0) As Long
  If OleTranslateColor(clr, hPal, TranslateColor) Then
    TranslateColor = -1
  End If
End Function

Public Function RoundRect(ByVal hdc As Long, rcItem As RECT, ByVal oTopLeftColor As OLE_COLOR, ByVal oBottomRightColor As OLE_COLOR)
  Dim hPen As Long
  Dim hPenOld As Long
  Dim tp As POINTAPI

  hPen = CreatePen(PS_SOLID, 1, TranslateColor(oTopLeftColor))
  hPenOld = SelectObject(hdc, hPen)

  MoveToEx hdc, rcItem.Left + 1, rcItem.Bottom - 2, tp
  LineTo hdc, rcItem.Left, rcItem.Bottom - 3
  LineTo hdc, rcItem.Left, rcItem.Top + 2
  LineTo hdc, rcItem.Left + 3, rcItem.Top - 1
  MoveToEx hdc, rcItem.Left + 3, rcItem.Top, tp
  LineTo hdc, rcItem.Right - 2, rcItem.Top

  SelectObject hdc, hPenOld
  DeleteObject hPen


  hPen = CreatePen(PS_SOLID, 1, TranslateColor(oBottomRightColor))
  hPenOld = SelectObject(hdc, hPen)

  MoveToEx hdc, rcItem.Right - 2, rcItem.Top + 1, tp
  LineTo hdc, rcItem.Right - 1, rcItem.Top + 3
  LineTo hdc, rcItem.Right - 1, rcItem.Bottom - 3
  LineTo hdc, rcItem.Right - 3, rcItem.Bottom - 1
  LineTo hdc, rcItem.Left + 1, rcItem.Bottom - 1

  SelectObject hdc, hPenOld
  DeleteObject hPen
End Function

Public Function EraseRect(ByVal hdc As Long, rcItem As RECT, ByVal oBackColor As OLE_COLOR)
  Dim hPen As Long
  Dim hPenOld As Long
  Dim tp As POINTAPI
  hPen = CreatePen(PS_SOLID, 1, TranslateColor(oBackColor))
  hPenOld = SelectObject(hdc, hPen)

  MoveToEx hdc, rcItem.Left, rcItem.Top, tp
  LineTo hdc, rcItem.Right - 1, rcItem.Top
  LineTo hdc, rcItem.Right - 1, rcItem.Bottom - 1
  LineTo hdc, rcItem.Left, rcItem.Bottom - 1
  LineTo hdc, rcItem.Left, rcItem.Top

  SelectObject hdc, hPenOld
  DeleteObject hPen
End Function
