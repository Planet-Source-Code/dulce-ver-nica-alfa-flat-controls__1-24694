VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFlatEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ..:: cFlatEdit delivered to you by _aLfa_ ::..

' =================================================================
' Filename: cFlatEdit.cls
' Author:   _aLfa_ (Joao Dias - joao-dias@usa.net)
'           based on the cFlatControl.cls from Steve McMahon (steve@dogma.demon.co.uk)
' Date:     3 July 2001
' Requires: SSUBTMR.DLL or VBWSUBCLASS.DLL
'
' Turns edit boxes into a Office 97 style flat editboxes.
' Thanks to Peter Siebling and Paul Wilde for modification
' suggestions.
'
' Originally Based on MFC code by Kirk Stowell (see below)
'
' Visit vbAccelerator - the vb Programmer's resource
' http://vbaccelerator.com
' =================================================================

'
'// Copyright © 1998 Kirk Stowell
'// All Rights Reserved.
'//
'// Email: kstowel@sprynet.com
'// URL:   www.geocities.com/SiliconValley/Haven/8230
'//
'/////////////////////////////////////////////////////////////////////////////


Implements ISubclass

Private m_hWnd             As Long
Private m_bSubclass        As Boolean
Private m_bMouseOver       As Boolean

Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long

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
  If TypeName(objthis) = "TextBox" Then
    With objthis
      .Move .Left + 3 * Screen.TwipsPerPixelX, .Top + 2 * Screen.TwipsPerPixelY, .Width - 6 * Screen.TwipsPerPixelX, .Height - 4 * Screen.TwipsPerPixelY
    End With
  Else
    Exit Sub
  End If

  m_hWnd = lhWnd
  If (m_hWnd <> 0) Then
    lStyle = GetWindowLong(m_hWnd, GWL_STYLE)
    AttachMessage Me, m_hWnd, WM_PAINT
    AttachMessage Me, m_hWnd, WM_SETFOCUS
    AttachMessage Me, m_hWnd, WM_KILLFOCUS
    AttachMessage Me, m_hWnd, WM_MOUSEMOVE
    AttachMessage Me, m_hWnd, WM_TIMER
    m_bSubclass = True
  End If
End Sub

Private Sub pRelease()
  If (m_bSubclass) Then
    DetachMessage Me, m_hWnd, WM_PAINT
    DetachMessage Me, m_hWnd, WM_SETFOCUS
    DetachMessage Me, m_hWnd, WM_KILLFOCUS
    DetachMessage Me, m_hWnd, WM_MOUSEMOVE
    DetachMessage Me, m_hWnd, WM_TIMER
  End If
  m_hWnd = 0
End Sub

Private Sub DrawEdit(clrTopLeft As OLE_COLOR, clrBottomRight As OLE_COLOR)
  Dim rcItem As RECT
  Dim pDC As Long
  Dim tp As POINTAPI
  Dim hWndP As Long

  hWndP = GetParent(m_hWnd)
  GetWindowRect m_hWnd, rcItem
  tp.x = rcItem.Left: tp.y = rcItem.Top
  ScreenToClient hWndP, tp
  rcItem.Left = tp.x: rcItem.Top = tp.y
  tp.x = rcItem.Right: tp.y = rcItem.Bottom
  ScreenToClient hWndP, tp
  rcItem.Right = tp.x: rcItem.Bottom = tp.y
  InflateRect rcItem, 1, 1
  pDC = GetDC(hWndP)
  
  EraseRect pDC, rcItem, vbWindowBackground
  RoundRect pDC, rcItem, vbWindowBackground, vbWindowBackground
  InflateRect rcItem, 1, 1
  RoundRect pDC, rcItem, vbButtonFace, vbButtonFace
  
  If (IsWindowEnabled(m_hWnd) = 0) Then
    RoundRect pDC, rcItem, vb3DHighlight, vb3DHighlight
  Else
    InflateRect rcItem, 1, 1
    RoundRect pDC, rcItem, clrTopLeft, clrBottomRight
  End If
  
  DeleteDC pDC
End Sub

Private Sub OnPaint(ByVal bFocus As Boolean)
  If m_hWnd = 0 Then Exit Sub
  'used for paint
  If bFocus Then
    DrawEdit vbButtonShadow, vb3DHighlight
  Else
    DrawEdit vbButtonFace, vbButtonFace
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
  If (CurrentMessage = WM_PAINT) Then
    ISubclass_MsgResponse = emrPreprocess
  Else
    ISubclass_MsgResponse = emrPostProcess
  End If
End Property

Private Function ISubclass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Dim bFocus As Boolean

  Select Case iMsg
    Case WM_PAINT
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
