VERSION 5.00
Begin VB.Form frmHol 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Holidays"
   ClientHeight    =   4590
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   2280
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   306
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   152
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox List1 
      Height          =   4350
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "frmHol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const SWW_HPARENT = (-8)

Private Sub Form_Load()
SetWindowLong Me.hwnd, SWW_HPARENT, Form1.hwnd
Form_Resize
End Sub

Private Sub Form_Resize()
List1.Width = ScaleWidth - 2
List1.Height = ScaleHeight - 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
SetWindowLong Me.hwnd, SWW_HPARENT, 0
End Sub


