VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   840
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   1560
   LinkTopic       =   "Form1"
   ScaleHeight     =   840
   ScaleWidth      =   1560
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   75
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   450
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   75
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   75
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private cFlatCombo As clsFlatControl
Private cFlatCombo2 As cFlatControl

Private Sub Form_Load()
  Set cFlatCombo = New clsFlatControl
  cFlatCombo.Attach Combo1.hwnd
  
  Set cFlatCombo2 = New cFlatControl
  cFlatCombo2.Attach Combo2
End Sub
