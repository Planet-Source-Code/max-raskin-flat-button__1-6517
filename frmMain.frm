VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Flat Button"
   ClientHeight    =   1665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4740
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   4740
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFlat 
      Caption         =   "Flat Command Button"
      Height          =   525
      Left            =   1110
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   930
      Width           =   2295
   End
   Begin VB.CommandButton cmdDefault 
      Caption         =   "Default Command Button"
      Height          =   525
      Left            =   1110
      TabIndex        =   0
      Top             =   180
      Width           =   2295
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Flat Button, by Max Raskin
'March 2000

'Flat Buttons, just like in Visual C++


'The SetWindowLong API call used to change an attribute of a control in runtime -
'just like you set the control's style (style/extended) in the CreateWindowEx
'API, this one work in runtime.
'(It can change some other attributes except styles)
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

'The button style BS_FLAT used to change a button to a Flat one
Private Const BS_FLAT = &H8000&
'GWL_Style is the attribute we will use for changing the style of the button
Private Const GWL_STYLE = (-16)
'To set the button as a child window and not as a self dependent window
Private Const WS_CHILD = &H40000000

Private Sub cmdDefault_Click()
    MsgBox "This is the default button in Visual Basic"
End Sub

Private Sub cmdFlat_Click()
    MsgBox "This is a Flat Button, just like in Visual C++"
End Sub

Private Sub Form_Load()
    btnFlat cmdFlat
End Sub

'Here is a small function to change button to flat:-
Function btnFlat(Button As CommandButton)
    SetWindowLong cmdFlat.hwnd, GWL_STYLE, WS_CHILD Or BS_FLAT
    cmdFlat.Visible = True 'Make the button visible (its automaticly hidden when the SetWindowLong call is executed because we reset the button's Attributes)
End Function
