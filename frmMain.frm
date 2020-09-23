VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   1800
      Top             =   1320
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Counts in the Timer event
Private counter As Byte
'Refresh rate of the Screen Saver, pulled from the
'Registry
Private RefreshRate As Integer

'Used to get a colour to paint the background
'0 = White, 1 = Light Gray, 2 = Gray, 3 = Dark Gray
'4 = Black
Private Declare Function GetStockObject Lib "gdi32" _
    (ByVal nIndex As Long) As Long
    
'Used to fill the Screen with the appropriate colour
Private Declare Sub FillRect Lib "user32" (ByVal _
    hdc As Long, lpRect As RECT, ByVal hBrush As Long)

'Already defined in Common.Bas, but also as a private.
'Used also on this form, but can be removed for your
'own Screen Saver.
'Returns the handle to the Desktop Window. The desktop
'window covers the entire screen. The desktop window is
'the area on top of which all icons and other windows
'are painted.
Private Declare Function GetDesktopWindow Lib _
    "user32" () As Long
Private Sub Form_Initialize()
    'Here I chose to set the Timer's interval to 1 second
    '(1000 milliseconds). You can set it to whatever
    'value you wish. This timer will be used to handle
    'any events in your Screen Saver (e.g. animating
    'sprites).
    RefreshRate = GetSetting(App.Title, "Settings", "Refresh", 1)
    
    'Timer is defaulted to 1 second if no settings have
    'been previously saved
    Timer1.Interval = RefreshRate * 1000
    
    'EraseBackground will set frmMain's Picture to
    'whatever background you have deemed necessary
    EraseBackground
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'I want the Screen Saver shut down no matter what
    'after a key is pressed
    ShutDown
End Sub


Private Sub Form_Load()
    'Enable the Timer now, and it should handle the
    'rest
    Timer1.Enabled = True
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    'I want the Screen Saver to shut down on a mouse
    'down no matter what
    ShutDown
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Check to see if the program will shut down
    ShutDown
End Sub

Private Sub Timer1_Timer()
    'This is where you place all of your events
    
    'Simple events here, loop from 0 to 4 each time,
    'and display a new colour
    counter = counter + 1
    
    If counter > 4 Then counter = 0

    FillRect Me.hdc, ScrSaverRect, GetStockObject(counter)
    Me.Refresh
End Sub


