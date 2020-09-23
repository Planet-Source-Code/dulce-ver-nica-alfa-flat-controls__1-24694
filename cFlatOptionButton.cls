VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFlatOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ..:: cFlatOption delivered to you by _aLfa_ ::..

' =================================================================
' Filename: cFlatOption.cls
' Author:   _aLfa_ (Joao Dias - joao-dias@usa.net)
'           based on the cFlatControl.cls from Steve McMahon (steve@dogma.demon.co.uk)
' Requires: SSUBTMR.DLL or VBWSUBCLASS.DLL
' =================================================================

Implements ISubclass

Private m_hWnd             As Long
Private m_hWndParent       As Long
Private m_bSubclass        As Boolean
Private m_bMouseOver       As Boolean

Private Const BN_CLICKED = 0
Private Const BM_SETSTATE = &HF3
Private Const BM_SETCHECK = &HF1

Private Declare Function Arc Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long

Public Sub Attach(ByRef objthis As Object)
  Dim lStyle As Long
  Dim lhWnd As Long

  pRelease

  On Error Resume Next
  lhWnd = objthis.hwnd
  If (Err.Number <> 0) Then
    Err.Raise vbObjectError + 1049, _
        App.EXEName & ".cEdit", _
        "Incorrect control type passed to 'Attach' parameter - must be a control with a hWnd property."
    Exit Sub
  End If

  ' Check for standard VB types:
  If TypeName(objthis) = "OptionButton" Then
    m_hWndParent = GetParent(lhWnd)
  Else
    Exit Sub
  End If

  m_hWnd = lhWnd
  If (m_hWnd <> 0) Then
    lStyle = GetWindowLong(m_hWnd, GWL_STYLE)
    AttachMessage Me, m_hWnd, WM_PAINT
    AttachMessage Me, m_hWnd, BM_SETSTATE
    AttachMessage Me, m_hWnd, BM_SETCHECK
    AttachMessage Me, m_hWnd, WM_SETFOCUS
    AttachMessage Me, m_hWnd, WM_KILLFOCUS
    AttachMessage Me, m_hWnd, WM_MOUSEMOVE
    AttachMessage Me, m_hWnd, WM_TIMER
    AttachMessage Me, m_hWndParent, WM_COMMAND
    m_bSubclass = True
  End If
End Sub

Private Sub pRelease()
  If (m_bSubclass) Then
    DetachMessage Me, m_hWnd, WM_PAINT
    DetachMessage Me, m_hWnd, BM_SETSTATE
    DetachMessage Me, m_hWnd, BM_SETCHECK
    DetachMessage Me, m_hWnd, WM_SETFOCUS
    DetachMessage Me, m_hWnd, WM_KILLFOCUS
    DetachMessage Me, m_hWnd, WM_MOUSEMOVE
    DetachMessage Me, m_hWnd, WM_TIMER
    DetachMessage Me, m_hWndParent, WM_COMMAND
  End If
  m_hWnd = 0: m_hWndParent = 0
End Sub

Private Sub DrawOption(clrTopLeft As OLE_COLOR, clrBottomRight As OLE_COLOR)
  Dim rcItem As RECT
  Dim pDC As Long
  Dim tp As POINTAPI
  Dim hWndP As Long
  
  GetClientRect m_hWnd, rcItem
  pDC = GetDC(m_hWnd)
    
  rcItem.Left = 1
  rcItem.Right = 13
  rcItem.Top = (rcItem.Bottom - rcItem.Top - 12) \ 2
  rcItem.Bottom = rcItem.Bottom - rcItem.Top - 1
  
  EraseRect pDC, rcItem, vbButtonFace
  InflateRect rcItem, -1, -1
  EraseRect pDC, rcItem, vbButtonFace
  InflateRect rcItem, -1, -1
  
  SetPixel pDC, rcItem.Left, rcItem.Top, TranslateColor(vbButtonFace)
  SetPixel pDC, rcItem.Left + 1, rcItem.Top, TranslateColor(vbButtonFace)
  SetPixel pDC, rcItem.Left, rcItem.Top + 1, TranslateColor(vbButtonFace)

  SetPixel pDC, rcItem.Right - 1, rcItem.Top, TranslateColor(vbButtonFace)
  SetPixel pDC, rcItem.Right - 1, rcItem.Top + 1, TranslateColor(vbButtonFace)
  SetPixel pDC, rcItem.Right - 2, rcItem.Top, TranslateColor(vbButtonFace)

  SetPixel pDC, rcItem.Right - 1, rcItem.Bottom - 1, TranslateColor(vbButtonFace)
  SetPixel pDC, rcItem.Right - 1, rcItem.Bottom - 2, TranslateColor(vbButtonFace)
  SetPixel pDC, rcItem.Right - 2, rcItem.Bottom - 1, TranslateColor(vbButtonFace)

  SetPixel pDC, rcItem.Left, rcItem.Bottom - 1, TranslateColor(vbButtonFace)
  SetPixel pDC, rcItem.Left + 1, rcItem.Bottom - 1, TranslateColor(vbButtonFace)
  SetPixel pDC, rcItem.Left, rcItem.Bottom - 2, TranslateColor(vbButtonFace)

  InflateRect rcItem, 1, 1
  If (IsWindowEnabled(m_hWnd) = 0) Then
    SetPixel pDC, rcItem.Left, rcItem.Bottom - 2, TranslateColor(vb3DHighlight)
    SetPixel pDC, rcItem.Left, rcItem.Bottom - 3, TranslateColor(vb3DHighlight)
    SetPixel pDC, rcItem.Left - 1, rcItem.Bottom - 4, TranslateColor(vb3DHighlight)
    SetPixel pDC, rcItem.Left - 1, rcItem.Bottom - 5, TranslateColor(vb3DHighlight)
    SetPixel pDC, rcItem.Left - 1, rcItem.Top + 4, TranslateColor(vb3DHighlight)
    SetPixel pDC, rcItem.Left - 1, rcItem.Top + 3, TranslateColor(vb3DHighlight)
    SetPixel pDC, rcItem.Left, rcItem.Top + 2, TranslateColor(vb3DHighlight)
    SetPixel pDC, rcItem.Left, rcItem.Top + 1, TranslateColor(vb3DHighlight)
    SetPixel pDC, rcItem.Left + 1, rcItem.Top, TranslateColor(vb3DHighlight)
    SetPixel pDC, rcItem.Left + 2, rcItem.Top, TranslateColor(vb3DHighlight)
    SetPixel pDC, rcItem.Left + 3, rcItem.Top - 1, TranslateColor(vb3DHighlight)
    SetPixel pDC, rcItem.Left + 4, rcItem.Top - 1, TranslateColor(vb3DHighlight)
    SetPixel pDC, rcItem.Left + 5, rcItem.Top - 1, TranslateColor(vb3DHighlight)
    SetPixel pDC, rcItem.Right - 4, rcItem.Top - 1, TranslateColor(vb3DHighlight)
    SetPixel pDC, rcItem.Right - 3, rcItem.Top, TranslateColor(vb3DHighlight)
    SetPixel pDC, rcItem.Right - 2, rcItem.Top, TranslateColor(vb3DHighlight)
  
    SetPixel pDC, rcItem.Right - 1, rcItem.Top + 1, TranslateColor(vb3DHighlight)
    SetPixel pDC, rcItem.Right - 1, rcItem.Top + 2, TranslateColor(vb3DHighlight)
    SetPixel pDC, rcItem.Right, rcItem.Top + 3, TranslateColor(vb3DHighlight)
    SetPixel pDC, rcItem.Right, rcItem.Top + 4, TranslateColor(vb3DHighlight)
    SetPixel pDC, rcItem.Right, rcItem.Bottom - 5, TranslateColor(vb3DHighlight)
    SetPixel pDC, rcItem.Right, rcItem.Bottom - 4, TranslateColor(vb3DHighlight)
    SetPixel pDC, rcItem.Right - 1, rcItem.Bottom - 3, TranslateColor(vb3DHighlight)
    SetPixel pDC, rcItem.Right - 1, rcItem.Bottom - 2, TranslateColor(vb3DHighlight)
    SetPixel pDC, rcItem.Right - 2, rcItem.Bottom - 1, TranslateColor(vb3DHighlight)
    SetPixel pDC, rcItem.Right - 3, rcItem.Bottom - 1, TranslateColor(vb3DHighlight)
    SetPixel pDC, rcItem.Right - 4, rcItem.Bottom, TranslateColor(vb3DHighlight)
    SetPixel pDC, rcItem.Right - 5, rcItem.Bottom, TranslateColor(vb3DHighlight)
    SetPixel pDC, rcItem.Right - 6, rcItem.Bottom, TranslateColor(vb3DHighlight)
    SetPixel pDC, rcItem.Left + 3, rcItem.Bottom, TranslateColor(vb3DHighlight)
    SetPixel pDC, rcItem.Left + 2, rcItem.Bottom - 1, TranslateColor(vb3DHighlight)
    SetPixel pDC, rcItem.Left + 1, rcItem.Bottom - 1, TranslateColor(vb3DHighlight)
  
    InflateRect rcItem, 1, 1
    
    SetPixel pDC, rcItem.Left, rcItem.Bottom - 2, TranslateColor(vb3DHighlight)
    SetPixel pDC, rcItem.Left, rcItem.Bottom - 3, TranslateColor(vb3DHighlight)
    SetPixel pDC, rcItem.Left - 1, rcItem.Bottom - 4, TranslateColor(vb3DHighlight)
    SetPixel pDC, rcItem.Left - 1, rcItem.Bottom - 5, TranslateColor(vb3DHighlight)
    SetPixel pDC, rcItem.Left - 1, rcItem.Top + 4, TranslateColor(vb3DHighlight)
    SetPixel pDC, rcItem.Left - 1, rcItem.Top + 3, TranslateColor(vb3DHighlight)
    SetPixel pDC, rcItem.Left, rcItem.Top + 2, TranslateColor(vb3DHighlight)
    SetPixel pDC, rcItem.Left, rcItem.Top + 1, TranslateColor(vb3DHighlight)
    SetPixel pDC, rcItem.Left + 1, rcItem.Top, TranslateColor(vb3DHighlight)
    SetPixel pDC, rcItem.Left + 2, rcItem.Top, TranslateColor(vb3DHighlight)
    SetPixel pDC, rcItem.Left + 3, rcItem.Top - 1, TranslateColor(vb3DHighlight)
    SetPixel pDC, rcItem.Left + 4, rcItem.Top - 1, TranslateColor(vb3DHighlight)
    SetPixel pDC, rcItem.Left + 5, rcItem.Top - 1, TranslateColor(vb3DHighlight)
    SetPixel pDC, rcItem.Right - 4, rcItem.Top - 1, TranslateColor(vb3DHighlight)
    SetPixel pDC, rcItem.Right - 3, rcItem.Top, TranslateColor(vb3DHighlight)
    SetPixel pDC, rcItem.Right - 2, rcItem.Top, TranslateColor(vb3DHighlight)
  
    SetPixel pDC, rcItem.Right - 1, rcItem.Top + 1, TranslateColor(vb3DHighlight)
    SetPixel pDC, rcItem.Right - 1, rcItem.Top + 2, TranslateColor(vb3DHighlight)
    SetPixel pDC, rcItem.Right, rcItem.Top + 3, TranslateColor(vb3DHighlight)
    SetPixel pDC, rcItem.Right, rcItem.Top + 4, TranslateColor(vb3DHighlight)
    SetPixel pDC, rcItem.Right, rcItem.Bottom - 5, TranslateColor(vb3DHighlight)
    SetPixel pDC, rcItem.Right, rcItem.Bottom - 4, TranslateColor(vb3DHighlight)
    SetPixel pDC, rcItem.Right - 1, rcItem.Bottom - 3, TranslateColor(vb3DHighlight)
    SetPixel pDC, rcItem.Right - 1, rcItem.Bottom - 2, TranslateColor(vb3DHighlight)
    SetPixel pDC, rcItem.Right - 2, rcItem.Bottom - 1, TranslateColor(vb3DHighlight)
    SetPixel pDC, rcItem.Right - 3, rcItem.Bottom - 1, TranslateColor(vb3DHighlight)
    SetPixel pDC, rcItem.Right - 4, rcItem.Bottom, TranslateColor(vb3DHighlight)
    SetPixel pDC, rcItem.Right - 5, rcItem.Bottom, TranslateColor(vb3DHighlight)
    SetPixel pDC, rcItem.Right - 6, rcItem.Bottom, TranslateColor(vb3DHighlight)
    SetPixel pDC, rcItem.Left + 3, rcItem.Bottom, TranslateColor(vb3DHighlight)
    SetPixel pDC, rcItem.Left + 2, rcItem.Bottom - 1, TranslateColor(vb3DHighlight)
    SetPixel pDC, rcItem.Left + 1, rcItem.Bottom - 1, TranslateColor(vb3DHighlight)
  Else
    SetPixel pDC, rcItem.Left, rcItem.Bottom - 2, TranslateColor(clrTopLeft)
    SetPixel pDC, rcItem.Left, rcItem.Bottom - 3, TranslateColor(clrTopLeft)
    SetPixel pDC, rcItem.Left - 1, rcItem.Bottom - 4, TranslateColor(clrTopLeft)
    SetPixel pDC, rcItem.Left - 1, rcItem.Bottom - 5, TranslateColor(clrTopLeft)
    SetPixel pDC, rcItem.Left - 1, rcItem.Top + 4, TranslateColor(clrTopLeft)
    SetPixel pDC, rcItem.Left - 1, rcItem.Top + 3, TranslateColor(clrTopLeft)
    SetPixel pDC, rcItem.Left, rcItem.Top + 2, TranslateColor(clrTopLeft)
    SetPixel pDC, rcItem.Left, rcItem.Top + 1, TranslateColor(clrTopLeft)
    SetPixel pDC, rcItem.Left + 1, rcItem.Top, TranslateColor(clrTopLeft)
    SetPixel pDC, rcItem.Left + 2, rcItem.Top, TranslateColor(clrTopLeft)
    SetPixel pDC, rcItem.Left + 3, rcItem.Top - 1, TranslateColor(clrTopLeft)
    SetPixel pDC, rcItem.Left + 4, rcItem.Top - 1, TranslateColor(clrTopLeft)
    SetPixel pDC, rcItem.Left + 5, rcItem.Top - 1, TranslateColor(clrTopLeft)
    SetPixel pDC, rcItem.Right - 4, rcItem.Top - 1, TranslateColor(clrTopLeft)
    SetPixel pDC, rcItem.Right - 3, rcItem.Top, TranslateColor(clrTopLeft)
    SetPixel pDC, rcItem.Right - 2, rcItem.Top, TranslateColor(clrTopLeft)
  
    SetPixel pDC, rcItem.Right - 1, rcItem.Top + 1, TranslateColor(clrBottomRight)
    SetPixel pDC, rcItem.Right - 1, rcItem.Top + 2, TranslateColor(clrBottomRight)
    SetPixel pDC, rcItem.Right, rcItem.Top + 3, TranslateColor(clrBottomRight)
    SetPixel pDC, rcItem.Right, rcItem.Top + 4, TranslateColor(clrBottomRight)
    SetPixel pDC, rcItem.Right, rcItem.Bottom - 5, TranslateColor(clrBottomRight)
    SetPixel pDC, rcItem.Right, rcItem.Bottom - 4, TranslateColor(clrBottomRight)
    SetPixel pDC, rcItem.Right - 1, rcItem.Bottom - 3, TranslateColor(clrBottomRight)
    SetPixel pDC, rcItem.Right - 1, rcItem.Bottom - 2, TranslateColor(clrBottomRight)
    SetPixel pDC, rcItem.Right - 2, rcItem.Bottom - 1, TranslateColor(clrBottomRight)
    SetPixel pDC, rcItem.Right - 3, rcItem.Bottom - 1, TranslateColor(clrBottomRight)
    SetPixel pDC, rcItem.Right - 4, rcItem.Bottom, TranslateColor(clrBottomRight)
    SetPixel pDC, rcItem.Right - 5, rcItem.Bottom, TranslateColor(clrBottomRight)
    SetPixel pDC, rcItem.Right - 6, rcItem.Bottom, TranslateColor(clrBottomRight)
    SetPixel pDC, rcItem.Left + 3, rcItem.Bottom, TranslateColor(clrBottomRight)
    SetPixel pDC, rcItem.Left + 2, rcItem.Bottom - 1, TranslateColor(clrBottomRight)
    SetPixel pDC, rcItem.Left + 1, rcItem.Bottom - 1, TranslateColor(clrBottomRight)
  End If
  
  DeleteDC pDC
End Sub

Private Sub OnPaint(ByVal bFocus As Boolean)
  If m_hWnd = 0 Then Exit Sub
  'used for paint
  If bFocus Then
    DrawOption vbButtonShadow, vb3DHighlight
  Else
    DrawOption vbButtonFace, vbButtonFace
  End If
End Sub

Private Sub Class_Terminate()
  pRelease
End Sub

Private Sub OnTimer(ByVal bCheckMouse As Boolean)
  Dim bOver As Boolean
  Dim rcItem As RECT
  Dim tp As POINTAPI

  If (bCheckMouse) Then
    bOver = True
    GetCursorPos tp
    GetWindowRect m_hWnd, rcItem
    If (PtInRect(rcItem, tp.x, tp.y) = 0) Then
      bOver = False
    End If
  End If

  If Not (bOver) Then
    KillTimer m_hWnd, 1
    m_bMouseOver = False
  End If
End Sub

Private Property Let ISubclass_MsgResponse(ByVal RHS As vbwSubClass.EMsgResponse)
  ' do not remove this comment
End Property

Private Property Get ISubclass_MsgResponse() As vbwSubClass.EMsgResponse
  Select Case CurrentMessage
    Case WM_PAINT, BM_SETSTATE, BM_SETCHECK
      ISubclass_MsgResponse = emrPreprocess
    Case Else
      ISubclass_MsgResponse = emrPostProcess
  End Select
End Property

Private Function ISubclass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Dim bFocus As Boolean

  Select Case iMsg
    Case WM_COMMAND
      OnPaint False
      Select Case wParam \ &H10000
      Case BN_CLICKED
        OnPaint True
      End Select
      OnTimer False
    
    Case WM_PAINT, BM_SETSTATE, BM_SETCHECK
      bFocus = (m_hWnd = GetFocus())
      OnPaint bFocus
      If (bFocus) Then
        OnTimer False
      End If

    Case WM_SETFOCUS
      OnPaint True
      OnTimer False

    Case WM_KILLFOCUS
      OnPaint False

    Case WM_MOUSEMOVE
      If Not (m_bMouseOver) Then
        If Not (m_hWnd = GetFocus()) Then
          OnPaint True
          m_bMouseOver = True
          ' Start checking to see if mouse is no longer over.
          SetTimer m_hWnd, 1, 10, 0
        End If
      End If

    Case WM_TIMER
      OnTimer True
      If Not (m_bMouseOver) Then
        OnPaint False
      End If
  End Select
End Function
