VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tony's AutoClicker"
   ClientHeight    =   1425
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3960
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1425
   ScaleWidth      =   3960
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrCountdown 
      Interval        =   1000
      Left            =   1560
      Top             =   120
   End
   Begin VB.Timer tmrClicker 
      Interval        =   100
      Left            =   240
      Top             =   120
   End
   Begin VB.Timer tmrSS 
      Interval        =   100
      Left            =   840
      Top             =   120
   End
   Begin VB.TextBox txtSpeed 
      Height          =   285
      Left            =   2520
      TabIndex        =   3
      Text            =   "txtSpeed"
      ToolTipText     =   "1 - 65535 ms"
      Top             =   840
      Width           =   615
   End
   Begin VB.Frame frmOptions 
      Caption         =   "Click options"
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1575
      Begin VB.ComboBox cboButton 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "Form1.frx":0442
         Left            =   120
         List            =   "Form1.frx":0444
         Style           =   2  'Dropdown List
         TabIndex        =   2
         ToolTipText     =   "Pick left or right click button"
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame frmSpeed 
      Caption         =   "Speed in milliseconds"
      Height          =   855
      Left            =   1920
      TabIndex        =   4
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label lbl1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "Press F12 to Start and Stop"
      Top             =   0
      Width           =   3975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Sub mouse_event Lib "user32.dll" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)

Private Const MOUSEEVENTF_RIGHTDOWN = &H8
Private Const MOUSEEVENTF_RIGHTUP = &H10

Private Const MOUSEEVENTF_LEFTDOWN = &H2
Private Const MOUSEEVENTF_LEFTUP = &H4

''''''''''''''''' ALWAYS PUT WINDOW ON TOP ''''''''''''''''''''''''''
Private Declare Function SetWindowPos Lib "user32" _
                    (ByVal hwnd As Long, ByVal hWndInsertAfter As WndInsertAfterEnum, _
                    ByVal X As Long, ByVal Y As Long, ByVal cx As Long, _
                    ByVal cy As Long, ByVal wFlags As SetWindowPosFlagsEnum) As Long

Private Enum WndInsertAfterEnum
    HWND_BOTTOM = 1
    HWND_BROADCAST = &HFFFF&
    HWND_DESKTOP = 0
    HWND_NOTOPMOST = -2
    HWND_TOP = 0
    HWND_TOPMOST = -1
End Enum

Private Enum SetWindowPosFlagsEnum
    SWP_FRAMECHANGED = &H20
    SWP_DRAWFRAME = SWP_FRAMECHANGED
    SWP_HIDEWINDOW = &H80
    SWP_NOACTIVATE = &H10
    SWP_NOCOPYBITS = &H100
    SWP_NOMOVE = &H2
    SWP_NOOWNERZORDER = &H200
    SWP_NOREDRAW = &H8
    SWP_NOREPOSITION = SWP_NOOWNERZORDER
    SWP_NOSIZE = &H1
    SWP_NOZORDER = &H4
    SWP_SHOWWINDOW = &H40
End Enum
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub Mouse_RightClick()
  mouse_event MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0
  mouse_event MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0
End Sub

Private Sub Mouse_LeftClick()
  mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
  mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
End Sub

Private Function FnKeyDown() As Boolean
  If (GetAsyncKeyState(vbKeyF12)) Then
    FnKeyDown = True
  Else
    FnKeyDown = False
  End If
End Function

Private Sub Form_Load()
If App.PrevInstance = True Then
   Unload Me
   Exit Sub
End If

lbl1.BackColor = vbRed
lbl1.ForeColor = vbWhite
lbl1.Caption = "      Stopped! F12 to Start"

cboButton.AddItem ("Left button")
cboButton.AddItem ("Right button")
cboButton.ListIndex = 0

txtSpeed.Text = "1000"        'sets click interval speed text to 1 second or 1000 ms
tmrSS.Interval = "100"        'sets the start/stop interval to check for F12 key press
tmrClicker.Interval = "0"     'sets the click interval to zero
tmrClicker.Enabled = False    'turn off the clicking when form is loaded
txtSpeed.MaxLength = 5        'sets the max length of chars in box
tmrCountdown.Enabled = False

Call enableAll

SetWindowPos Form1.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE    ' puts on top
'SetWindowPos Form.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE  ' stop being on top
End Sub

Private Sub lbl1_Click()
   lbl1.BackColor = vbRed
   lbl1.ForeColor = vbWhite
   lbl1.Caption = "      Stopped! F12 to Start"
   tmrClicker.Enabled = False
   Form1.Height = 1860
   Call enableAll
End Sub

Private Sub tmrSS_Timer()     'StartStop timer
If FnKeyDown = True Then
tmrClicker.Interval = "0"
tmrClicker.Enabled = Not (tmrClicker.Enabled)

If tmrClicker.Enabled = True Then
Call txtSanityCheck
tmrClicker.Interval = Val(txtSpeed.Text)
   lbl1.BackColor = vbGreen
   lbl1.ForeColor = vbBlack
   lbl1.Caption = "      Running! F12 to Stop"
   Form1.Height = 810
   Call disableAll

ElseIf tmrClicker.Enabled = False Then
   lbl1.BackColor = vbRed
   lbl1.ForeColor = vbWhite
   lbl1.Caption = "      Stopped! F12 to Start"
   tmrClicker.Enabled = False
   Form1.Height = 1860
   Call enableAll
End If
End If
End Sub

Private Sub tmrClicker_Timer()
If cboButton.ListIndex = 0 Then  'left button is selected
  Mouse_LeftClick
Else
  Mouse_RightClick
End If
End Sub

Private Sub disableAll()
cboButton.Enabled = False
txtSpeed.Enabled = False
End Sub

Private Sub txtSanityCheck()
If txtSpeed.Text = "0" Or txtSpeed.Text = "" Then
txtSpeed.Text = "1"
End If

If Val(txtSpeed.Text) > 65535 Then
txtSpeed.Text = "65535"
End If

End Sub

Private Sub enableAll()
cboButton.Enabled = True
txtSpeed.Enabled = True
End Sub

Private Sub txtSpeed_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then       'if Enter is pressed, go to the next box
    End If
    If KeyAscii < 48 Or KeyAscii > 57 Then  'allows 0-9
      If KeyAscii <> 8 Then                 'allow the backspace
       KeyAscii = 0
      End If
    End If
End Sub
