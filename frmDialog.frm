VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDialog 
   Caption         =   "Screen Saver Properties"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   1680
      TabIndex        =   6
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   3120
      TabIndex        =   5
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Frame fraSpeed 
      Caption         =   "Redraw Rate"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3855
      Begin MSComctlLib.Slider sldRefresh 
         Height          =   495
         Left            =   480
         TabIndex        =   1
         Top             =   1080
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   873
         _Version        =   393216
         LargeChange     =   1
         Min             =   1
         SelStart        =   1
         Value           =   1
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Select the time interval you wish between redraws"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   600
         TabIndex        =   4
         Top             =   360
         Width           =   2445
      End
      Begin VB.Label lblMax 
         AutoSize        =   -1  'True
         Caption         =   "10 seconds"
         Height          =   195
         Left            =   2880
         TabIndex        =   3
         Top             =   1680
         Width           =   825
      End
      Begin VB.Label lblMin 
         AutoSize        =   -1  'True
         Caption         =   "1 second"
         Height          =   195
         Left            =   600
         TabIndex        =   2
         Top             =   1680
         Width           =   660
      End
   End
End
Attribute VB_Name = "frmDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This will hold the value of the rate at which the
'screen saver will refresh
Private RefreshRate As Integer

Private Sub cmdCancel_Click()
    'Cancel clicked, no changes are saved, just shut
    'down
    ShutDownDialog
End Sub

Private Sub cmdOK_Click()
    'OK is clicked and we will save the changes in the
    'Dialog
    RefreshRate = sldRefresh.Value
    
    'Writing to Registry here
    SaveSetting App.Title, "Settings", "Refresh", RefreshRate
    
    'Shut down when complete
    ShutDownDialog
End Sub



Private Sub Form_Initialize()
    'Get the current refresh rate from the Windows
    'Registry
    RefreshRate = GetSetting(App.Title, "Settings", "Refresh", 0)
    
    'If no value is returned (0 is the default), it is
    'the first time that the App is being run or the
    'registry has been cleared. We'll save the default
    'now
    If RefreshRate = 0 Then
        SaveSetting App.Title, "Settings", "Refresh", 1
        RefreshRate = 1
    End If
End Sub

Private Sub Form_Load()
    'Set up the Slider to be that saved in the Registry
    sldRefresh.Value = RefreshRate
End Sub



Private Sub ShutDownDialog()
    'This is the final shutdown for this procedure
    'If you declared any variables, release them in the
    'frmDialog_UnLoad event
    Dim FRM As Form
    
    'Clears all forms started in the App
    For Each FRM In Forms
        Unload FRM
    Next
    
    End
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Release any variables here
End Sub


