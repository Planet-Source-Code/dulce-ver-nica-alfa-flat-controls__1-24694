VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFlatCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ..:: cFlatCombo delivered to you by _aLfa_ ::..

' =================================================================
' Filename: cFlatCombo.cls
' Author:   _aLfa_ (Joao Dias - joao-dias@usa.net)
'           based on the cFlatControl.cls from Steve McMahon (steve@dogma.demon.co.uk)
' Date:     2 July 2001
' Requires: SSUBTMR.DLL or VBWSUBCLASS.DLL
'
' Turns combo boxes into a Office 97 style flat controls.
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
Private m_hWndEdit         As Long
Private m_hWndParent       As Long
Private m_bSubclass        As Boolean
Private m_bMouseOver       As Boolean

Private Enum EDrawStyle
  FC_DRAWNORMAL = &H1
  FC_DRAWRAISED = &H2
  FC_DRAWPRESSED = &H4
End Enum

Private Const SM_CXHTHUMB = 10

Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Const CBS_DROPDOWN = &H2&
Private Const CBS_DROPDOWNLIST = &H3&
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, lpsz2 As Any) As Long
Private Const CBN_CLOSEUP = 8
Private Const CB_GETDROPPEDSTATE = &H157
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Const GW_CHILD = 5

Public Sub Attach(ByRef objthis As Object)
  Dim lStyle As Long
  Dim lhWnd As Long

  pRelease

  On Error Resume Next
  lhWnd = objthis.hwnd
  If (Err.Number <> 0) Then
    Err.Raise vbObjectError + 1049, _
        App.EXEName & ".cFlatCombo", _
        "Incorrect control type passed to 'Attach' parameter - must be a control with a hWnd property."
    Exit Sub
  End If

  ' ..:: See if the control is a valid ComboBox ::..
  If TypeName(objthis) = "ImageCombo" Then
    ' Support VB6 ImageCombos.  Thanks to Paul Wilde for the tip.
    ' WM_COMMAND notifications to the control:
    m_hWndParent = lhWnd
    ' Get the actual ComboBox in the control:
    lhWnd = FindWindowEx(lhWnd, 0&, "ComboBox", ByVal 0&)
  ElseIf TypeName(objthis) = "ComboBox" Then
    ' WM_COMMAND notifications go to the parent:
    m_hWndParent = GetParent(objthis.hwnd)
  ElseIf TypeName(objthis) = "OwnerDrawComboList" Then
    ' WM_COMMAND notifications to the control:
    m_hWndParent = lhWnd
  Else
    lStyle = GetWindowLong(lhWnd, GWL_STYLE)
    If ((lStyle And CBS_DROPDOWN) = CBS_DROPDOWN) Or ((lStyle And CBS_DROPDOWNLIST) = CBS_DROPDOWNLIST) Then
      ' An API combo box of some sort.
      ' Assume WM_COMMAND notifications go to the parent form:
      m_hWndParent = objthis.Parent.hwnd
    Else
      Exit Sub
    End If
  End If

  ' ..:: Attaching ::..
  m_hWnd = lhWnd
  If (m_hWnd <> 0) Then
    lStyle = GetWindowLong(m_hWnd, GWL_STYLE)
    If (lStyle And CBS_DROPDOWN) = CBS_DROPDOWN Then
      m_hWndEdit = GetWindow(m_hWnd, GW_CHILD)
    End If
    AttachMessage Me, m_hWnd, WM_PAINT
    AttachMessage Me, m_hWnd, WM_SETFOCUS
    AttachMessage Me, m_hWnd, WM_KILLFOCUS
    AttachMessage Me, m_hWnd, WM_MOUSEMOVE
    AttachMessage Me, m_hWnd, WM_TIMER
    If (m_hWndEdit <> 0) Then
      AttachMessage Me, m_hWndEdit, WM_SETFOCUS
      AttachMessage Me, m_hWndEdit, WM_KILLFOCUS
      AttachMessage Me, m_hWndEdit, WM_MOUSEMOVE
    End If
    AttachMessage Me, m_hWndParent, WM_COMMAND
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
    If (m_hWndEdit <> 0) Then
      DetachMessage Me, m_hWndEdit, WM_SETFOCUS
      DetachMessage Me, m_hWndEdit, WM_KILLFOCUS
      DetachMessage Me, m_hWndEdit, WM_MOUSEMOVE
    End If
    DetachMessage Me, m_hWndParent, WM_COMMAND
  End If
  m_hWnd = 0: m_hWndEdit = 0: m_hWndParent = 0
End Sub

Private Sub DrawCombo(ByVal dwStyle As EDrawStyle, clrTopLeft As OLE_COLOR, clrBottomRight As OLE_COLOR)
  Dim rcItem As RECT
  Dim rcItem2 As RECT
  Dim pDC As Long
  Dim hWndFocus As Long
  Dim tp As POINTAPI


  GetClientRect m_hWnd, rcItem
  pDC = GetDC(m_hWnd)

  '// Cover up dark 3D shadow on drop arrow.
  InflateRect rcItem, -2, -2
  rcItem.Left = rcItem.Right - Offset()
  EraseRect pDC, rcItem, vbButtonFace
  RoundRect pDC, rcItem, vbButtonFace, vbButtonFace

  '// Cover up normal 3D shadow on drop arrow.
  InflateRect rcItem, -1, -1
  EraseRect pDC, rcItem, vbButtonFace
  RoundRect pDC, rcItem, vbButtonFace, vbButtonFace

  Select Case dwStyle
    Case FC_DRAWNORMAL
      rcItem.Top = rcItem.Top - 1
      rcItem.Bottom = rcItem.Bottom + 1
      EraseRect pDC, rcItem, vbButtonFace
      RoundRect pDC, rcItem, vb3DHighlight, vb3DHighlight
      If (IsWindowEnabled(m_hWnd) = 0) Then
        EraseRect pDC, rcItem, vb3DHighlight
        InflateRect rcItem, 1, 1
        RoundRect pDC, rcItem, vb3DHighlight, vb3DHighlight
      End If
      rcItem.Left = rcItem.Left - 1
      rcItem.Right = rcItem.Left
      EraseRect pDC, rcItem, vbWindowBackground
      RoundRect pDC, rcItem, vbWindowBackground, vbWindowBackground
      SetPixel pDC, rcItem.Left + 1, rcItem.Bottom - 1, TranslateColor(vbWindowBackground)
      SetPixel pDC, rcItem.Left + 2, rcItem.Bottom - 1, TranslateColor(vbWindowBackground)
    Case FC_DRAWRAISED
      rcItem.Top = rcItem.Top - 1
      rcItem.Bottom = rcItem.Bottom + 1
      rcItem.Right = rcItem.Right + 1
      EraseRect pDC, rcItem, vbButtonFace
      RoundRect pDC, rcItem, vb3DHighlight, vbButtonShadow
      SetPixel pDC, rcItem.Left - 1, rcItem.Top, TranslateColor(vbWindowBackground)
      SetPixel pDC, rcItem.Left - 1, rcItem.Bottom - 1, TranslateColor(vbWindowBackground)
    Case FC_DRAWPRESSED
      rcItem.Left = rcItem.Left - 1
      rcItem.Top = rcItem.Top - 2
      OffsetRect rcItem, 1, 1
      RoundRect pDC, rcItem, vbButtonShadow, vb3DHighlight
      SetPixel pDC, rcItem.Left - 1, rcItem.Top, TranslateColor(vbWindowBackground)
      SetPixel pDC, rcItem.Left - 1, rcItem.Bottom - 1, TranslateColor(vbWindowBackground)
  End Select

  '****************************************
  GetClientRect m_hWnd, rcItem
  pDC = GetDC(m_hWnd)

  InflateRect rcItem, -1, -1
  If (IsWindowEnabled(m_hWnd) = 0) Then
    EraseRect pDC, rcItem, vbButtonFace
    RoundRect pDC, rcItem, vb3DHighlight, vb3DHighlight
  Else
    EraseRect pDC, rcItem, vbButtonFace
    RoundRect pDC, rcItem, vbButtonFace, vbButtonFace
  End If

  '// Cover up dark 3D shadow.
  InflateRect rcItem, 1, 1
  EraseRect pDC, rcItem, vbButtonFace
  RoundRect pDC, rcItem, clrTopLeft, clrBottomRight

  DeleteDC pDC
End Sub

Private Function Offset() As Long
  ' Thanks to Todd Brannam for this suggestion...
  Offset = GetSystemMetrics(SM_CXHTHUMB)
End Function

Public Property Get DroppedDown() As Boolean
  If m_hWnd <> 0 Then
    DroppedDown = (SendMessageLong(m_hWnd, CB_GETDROPPEDSTATE, 0, 0) <> 0)
  End If
End Property

Private Sub OnPaint(ByVal bFocus As Boolean, ByVal bDropped As Boolean)
  If m_hWnd = 0 Then Exit Sub
    
  'used for paint
  If bFocus Then
    If (bDropped) Then
      DrawCombo FC_DRAWPRESSED, vbButtonShadow, vb3DHighlight
    Else
      DrawCombo FC_DRAWRAISED, vbButtonShadow, vb3DHighlight
    End If
  Else
    DrawCombo FC_DRAWNORMAL, vbButtonFace, vbButtonFace
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
  Dim bDown As Boolean
  Dim bFocus As Boolean

  Select Case iMsg
    Case WM_COMMAND
      If (m_hWnd = lParam) Then
        ' Type of notification is in the hiword of wParam:
        Select Case wParam \ &H10000
          Case CBN_CLOSEUP
            OnPaint (m_hWnd = GetFocus() Or m_hWndEdit = GetFocus() Or bDown), bDown
        End Select
        OnTimer False
      End If

    Case WM_PAINT
      bDown = DroppedDown()
      bFocus = (m_hWnd = GetFocus() Or m_hWndEdit = GetFocus() Or bDown)
      OnPaint (bFocus), bDown
      If (bFocus) Then
        OnTimer False
      End If

    Case WM_SETFOCUS
      OnPaint True, False
      OnTimer False

    Case WM_KILLFOCUS
      OnPaint False, False

    Case WM_MOUSEMOVE
      If Not (m_bMouseOver) Then
        bDown = DroppedDown()
        If Not (m_hWnd = GetFocus() Or m_hWndEdit = GetFocus() Or bDown) Then
          OnPaint True, False
          m_bMouseOver = True
          ' Start checking to see if mouse is no longer over.
          SetTimer m_hWnd, 1, 10, 0
        End If
      End If

    Case WM_TIMER
      OnTimer True
      If Not (m_bMouseOver) Then
        OnPaint False, False
      End If
  End Select
End Function
