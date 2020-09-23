Attribute VB_Name = "Common"
Option Explicit
'Project: Screen Saver Module
'Author: Tim Fournier
'Contact: tim_fournier@hotmail.com
'
'Last Update: February 2002
'
'If you are distributing this module to anyone, please
'do not modify its contents without contacting me. I
'would like to have any changes to it so that I can add
'it to my own source.
'
'Also, please don't post it up somewhere and say its
'your own, people like to receive (and give) credit
'where its due. I had to do quite a bit of research to
'get all of this together.
'
'This module, when combined with a Screen Saver of your
'own design, should handle all of the windows settings
'and options normally associated with a Screen Saver.
'I felt that something similar to this was needed as
'many people have posted some cool Screen Savers,
'graphically, but I haven't seen one that allows you
'to do all the tasks a Screen Saver normally performs
'(change passwords and change settings, etc).
'
'UPDATE: This program has been updated to work in the
'WinNT environment. Special thanks to Kyle Burns for
'these additions.
'
'Enjoy!

'Shows the mouse position
Public CurrentMouseX As Long
Public CurrentMouseY As Long

'Basic RECT structure, used for setting the parameters
'of the Screen Saver window
Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

'This will hold the display parameters for the screen
'saver
Public ScrSaverRect As RECT

'Device Mode type, will hold all the values for any
'given device (e.g. Display)
Private Type DEVMODE
        dmDeviceName As String * 32
        dmSpecVersion As Integer
        dmDriverVersion As Integer
        dmSize As Integer
        dmDriverExtra As Integer
        dmFields As Long
        dmOrientation As Integer
        dmPaperSize As Integer
        dmPaperLength As Integer
        dmPaperWidth As Integer
        dmScale As Integer
        dmCopies As Integer
        dmDefaultSource As Integer
        dmPrintQuality As Integer
        dmColor As Integer
        dmDuplex As Integer
        dmYResolution As Integer
        dmTTOption As Integer
        dmCollate As Integer
        dmFormName As String * 32
        dmUnusedPadding As Integer
        dmBitsPerPel As Long
        dmPelsWidth As Long
        dmPelsHeight As Long
        dmDisplayFlags As Long
        dmDisplayFrequency As Long
End Type

'This type will hold the value of a cursor position,
'it can be used to point to any single point on a
'display
Private Type POINTAPI
        x As Long
        y As Long
End Type

'Constants go here
Private Const SWP_SHOWWINDOW = &H40
Private Const HWND_TOP = 0
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOZORDER = &H4
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1

'These are used to set the preview as a child of the
'Control Panel's Display
Private Const GWL_HWNDPARENT = (-8)
Private Const GWL_STYLE = (-16)
Private Const WS_CHILD = &H40000000

'Used when first displaying the Screen Saver, will
'place frmMain as the highest zorder form.
Private Const HWND_TOPMOST = -1
Private Const SPI_SETSCREENSAVEACTIVE = 17
Private Const SPI_SCREENSAVERRUNNING = 97

'This will be used to retrieve the window parameters
'if the user wishes to see a screen saver preview
Private Declare Sub GetClientRect Lib "user32" _
    (ByVal hwnd As Long, lpRect As RECT)

'This allows us to enable or disable ALL input to any
'window, this includes all keypresses and mouseclicks
Private Declare Function EnableWindow Lib "user32" _
    (ByVal hwnd As Long, ByVal fEnable As Long) As Long

'This will tell us if any given window has mouse and
'keyboard inputs enabled (True or False returned)
Private Declare Function IsWindowEnabled Lib "user32" _
    (ByVal hwnd As Long) As Long

'This returns the handle of the first window found that
'matches the class name and window name given
Private Declare Function FindWindow Lib "user32" _
    Alias "FindWindowA" (ByVal lpClassName As String, _
    ByVal lpWindowName As String) As Long

'This will be used to enable or disable the current
'Windows cursor
Private Declare Function ShowCursor Lib "user32" _
    (ByVal bShow As Long) As Long

'This will return the current location of the cursor
Private Declare Sub GetCursorPos Lib "user32" (lpPoint _
    As POINTAPI)

'Used to tell the system that a Screen Saver is now
'running on the system
Private Declare Sub SystemParametersInfo Lib _
    "user32" Alias "SystemParametersInfoA" (ByVal _
    uAction As Long, ByVal uParam As Long, ByRef _
    lpvParam As Any, ByVal fuWinIni As Long)

'This function is used to connect to the Windows
'Registry
Private Declare Function RegConnectRegistry Lib _
    "advapi32.dll" Alias "RegConnectRegistryA" _
    (ByVal lpMachineName As String, ByVal hKey As _
    Long, phkResult As Long) As Long

'Returns the handle to the Desktop Window. The desktop
'window covers the entire screen. The desktop window is
'the area on top of which all icons and other windows
'are painted.
Private Declare Function GetDesktopWindow Lib _
    "user32" () As Long

'Used in the preview to make frmMain a child of the
'preview display window
Private Declare Sub SetParent Lib "user32" _
    (ByVal hWndChild As Long, ByVal hWndNewParent _
    As Long)

'Used to get any current style flags for any given
'window
Private Declare Function GetWindowLong Lib "user32" _
    Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal _
    nIndex As Long) As Long

'Used to set the style flags of the child window (the
'Screen Saver Preview)
Private Declare Sub SetWindowLong Lib "user32" _
    Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal _
    nIndex As Long, ByVal dwNewLong As Long)

'This is used to set the position of the window (frmMain)
'based on the parameters retrieved by the GetClientRect
'API call (See above)
Private Declare Sub SetWindowPos Lib "user32" _
    (ByVal hwnd As Long, ByVal hWndInsertAfter As _
    Long, ByVal x As Long, ByVal y As Long, ByVal cx _
    As Long, ByVal cy As Long, ByVal wFlags As Long)

'This is used to create a device context for a specific
'device using the specified name (Display)
Private Declare Function CreateDC Lib "gdi32" Alias _
    "CreateDCA" (ByVal lpDriverName As String, ByVal _
    lpDeviceName As Variant, ByVal lpOutput As _
    Variant, lpInitData As DEVMODE) As Long

'Bitblt here to transfer the current Windows Display
'to the image value of frmMain.
'If you are going to use this in your Screen Saver,
'SET IT TO PUBLIC
Private Declare Sub BitBlt Lib "gdi32" (ByVal _
    hDestDC As Long, ByVal x As Long, ByVal y As _
    Long, ByVal nWidth As Long, ByVal nHeight As _
    Long, ByVal hSrcDC As Long, ByVal xSrc As Long, _
    ByVal ySrc As Long, ByVal dwRop As Long)

'This procedure is called when the user clicks the
'change password button in the Screen Saver section
'of Control Panel. It handles all the password rules
'and write the encrypted password to the Registry
Private Declare Sub PwdChangePassword Lib "mpr.dll" _
    Alias "PwdChangePasswordA" (ByVal lpcRegkeyname _
    As String, ByVal hwnd As Long, ByVal uiReserved1 _
    As Long, ByVal uiReserved2 As Long)

'This Function is called when the user does something
'which would normally disable a Screen Saver. If the
'user chose the password protection option, this will
'be called. It returns True if the user entered a
'correct password, or False if they exited the dialog
'(ESC Button). Incorrect passwords simply loop through
'and the user is asked to enter it again.
Private Declare Function VerifyScreenSavePwd Lib _
    "password.cpl" (ByVal hwnd As Long) As Boolean

'The following three calls are used to get the information
'used in the APIFunctionPresent Function, allowing WinNT
'Kernel Systems to use this screensaver without crashing
'due to non-present API calls being used
'Added February 25, 2002
Private Declare Function LoadLibrary Lib "kernel32" _
  Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long

Private Declare Function GetProcAddress Lib "kernel32" _
  (ByVal hModule As Long, ByVal lpProcName As String) As Long

Private Declare Function FreeLibrary Lib "kernel32" _
  (ByVal hLibModule As Long) As Long
  
Public Function APIFunctionPresent(ByVal FunctionName As String, ByVal DllName As String) As Boolean
'This function is used to determine whether or not an API
'Call used in this application is present on the current
'operating system. Specifically, it allows WinNT Systems
'to run this program without crashing due to a call to
'the system's password handler not being necessary
'Added February 25, 2002
    Dim libHandle As Long
    Dim libAddr  As Long

    libHandle = LoadLibrary(DllName)
    If libHandle <> 0 Then
        libAddr = GetProcAddress(libHandle, FunctionName)
        FreeLibrary libHandle
    End If

    'Returns True if the API is present
    APIFunctionPresent = (libAddr <> 0)
End Function



Public Sub Main()
    Dim CommandLine As String
    'First thing we do is check for previous instances
    StartUp
    
    'Now, based on the Command Line, do the appropriate
    'command
    CommandLine = UCase(Trim(Command()))

    Select Case Left(CommandLine, 2)
        Case "/S"
            'This basically tells us to start up the
            'screen saver in full screen mode
            StartScreenSaver
        Case ""
            'This will probably happen if you are trying
            'to run it through VB or if you execute the
            '.SCR file
            StartScreenSaver
        Case "/P"
            'This is the call for the preview
            'The additional value is the handle to the
            'object windows wishes you to draw to
            
            'If statement used here in case a user with
            'some idea of a Screen Saver tries to run
            'their own command line argument without a
            'handle; just keeps it from crashing
            If Len(CommandLine) = 0 Then
                StartScreenSaver
            Else
                StartScreenSaver Int(Mid(CommandLine, 4, Len(CommandLine)))
            End If
        Case "/A"
            'This tells us that the user wishes to set
            'a new password for the screen saver.
            'Code courtesy of Marco Bellinaso from
            '101 Tech Tips Vol. 10 from the VBPJ
            
            'Windows also passes and additional message
            'through this string, a number which I assume
            'is a handle for the password box to link to
            
            'The first string must be "SCRSAVE" because
            'it is used by the function and will fail
            'otherwise
            PwdChangePassword "SCRSAVE", Int(Mid(CommandLine, 4, Len(CommandLine))), 0, 0
        Case "/C"
            'This is called if the user chose Settings
            'from the Screen Saver Menu in Control
            'Panel
            'If you have any options for your Screen
            'saver, add a call here, e.g.:
            frmDialog.Show vbModal
    End Select
End Sub


Private Sub StartScreenSaver(Optional ByVal targethWnd As Variant)
    'This is where the Screen Saver begins from
    'If a handle is passed, we will set the Screen
    'Saver's size parameters to that of the preview
    'display, otherwise, we will be setting it to
    'that of the Desktop Window
    
    'Used to set window styles
    Dim WindowStyle As Long
    'Temporary variable used in the SPI to disable
    'system keys
    Dim TempBoolean As Boolean
    'Used to get various values pertaining to the
    'Windows Task Bar
    Dim Taskbar As Long, TaskbarhWnd As Long, TaskBarEnabled As Integer
            
    If IsMissing(targethWnd) Then
        'Regular mode
        frmMain.Caption = "SCREEN SAVER"
        
        'Get the desktop's size, used to set the
        'window's size
        GetClientRect GetDesktopWindow, ScrSaverRect
        
        'Show the Screen Saver
        SetWindowPos frmMain.hwnd, 0, 0, 0, ScrSaverRect.Right, ScrSaverRect.Bottom, SWP_SHOWWINDOW
        
        'Make the cursor dissapear
        ShowCursor False
        
        'Initialize these two mouse variables, used in
        'the shut down procedure
        CurrentMouseX = 0
        CurrentMouseY = 0
        
        'Send a message telling Windows that a Screen
        'Saver is now active
        SystemParametersInfo SPI_SETSCREENSAVEACTIVE, 0, ByVal 0&, 0
        
        'Disable System Keys
        SystemParametersInfo SPI_SCREENSAVERRUNNING, True, TempBoolean, 0
        
        'Disable the system taskbar
        TaskbarhWnd = FindWindow("Shell_traywnd", "")

        'If you're not getting a handle, you have some
        'hardcore system problems to work through.
        'Nonetheless, here is a bit of error trapping
        If TaskbarhWnd <> 0 Then
            TaskBarEnabled = IsWindowEnabled(TaskbarhWnd)
            
            'Check if the Task bar is already disabled
            If TaskBarEnabled = 1 Then
                Taskbar = EnableWindow(TaskbarhWnd, 0)
            End If
        End If
    Else
        'Preview mode
        frmMain.Caption = "PREVIEW"
        GetClientRect targethWnd, ScrSaverRect
        
        
        'Pull the current style of frmMain
        WindowStyle = GetWindowLong(frmMain.hwnd, GWL_STYLE)
        
        'Next, add an or value to allow you to do a
        'switch between the previous style and WS_CHILD
        WindowStyle = WindowStyle Or WS_CHILD
        
        'Now we'll switch the style over to thew new
        'style (WS_CHILD)
        SetWindowLong frmMain.hwnd, GWL_STYLE, WindowStyle
        
        'Set the preview window as frmMain's parent
        SetParent frmMain.hwnd, targethWnd
        
        'Set the Display to be a parent window
        SetWindowLong frmMain.hwnd, GWL_HWNDPARENT, targethWnd
        
        'Show the Preview
        SetWindowPos frmMain.hwnd, HWND_TOP, 0, 0, ScrSaverRect.Right, ScrSaverRect.Bottom, SWP_NOZORDER Or SWP_NOACTIVATE Or SWP_SHOWWINDOW
    End If
End Sub

Public Sub EraseBackground()
    'If you want to make your Screen Saver's background
    'that of windows at this moment, uncomment the next
    'line:
    'DesktopWindow
    
    'Otherwise, set the background of frmMain (Picture
    'or Background, depending on whether the background
    'is a colour or a picture)
    frmMain.BackColor = vbWhite
End Sub

Private Sub DesktopWindow()
    'I still can't quite get this Sub Procedure to
    'work. If you know of a way to pull the current
    'Desktop's image, please get in touch with me

    'This will hold the value of the current Display
    'Device Context
    Dim ScreenDC As Long
    Dim DisplayMode As DEVMODE
    
    'Get the Display Device Context
    ScreenDC = CreateDC("DISPLAY" & Chr(0), 0, 0, DisplayMode)
    
    'Transfer the Display DC to frmmain.Picture
    BitBlt frmMain.hdc, 0, 0, (Screen.Width / Screen.TwipsPerPixelX), (Screen.Height / Screen.TwipsPerPixelY), ScreenDC, 0, 0, vbSrcCopy
End Sub

Public Sub ShutDown()
    'Used to check the current mode (Regular or Preview)
    Dim CurrentDesktopRect As RECT
    'Will tell us whether or not a correct password
    'was entered
    Dim CorrectPassword As Boolean
    'Used in the SPI call to enable system keys, not
    'used for anything else
    Dim TempBoolean As Boolean
    'Used to get the handle and status of the Task Bar
    Dim Taskbar As Long, TaskbarhWnd As Long
    'Holds the current cursor position
    Dim MousePos As POINTAPI
    'Used to shutdown all forms
    Dim FRM As Form

    'First we check to make sure we are running in
    'regular mode
    GetClientRect GetDesktopWindow, CurrentDesktopRect

    'Only in regular mode will the Screen Saver's rect
    'structure be equal to that of the Current Desktop
    If Not (ScrSaverRect.Right = CurrentDesktopRect.Right) Then
        Exit Sub
    End If
        
    'This will allow one grace move
    GetCursorPos MousePos
    
    If (CurrentMouseX = 0) And (CurrentMouseY = 0) Then
        'Exit if this is just the first move
        CurrentMouseX = MousePos.x
        CurrentMouseY = MousePos.y
        Exit Sub
    End If
    
    'Pause all Screen Saver Activities
    frmMain.Timer1.Enabled = False
    
    'Show the cursor
    ShowCursor True

    'Check to see if VerifyScreenSavePwd is available (only on Win9x)
    'If function is available, call the function, otherwise just
    'return True
    'Added February 25, 2002
    If APIFunctionPresent("VerifyScreenSavePwd", "password.cpl") Then
        CorrectPassword = VerifyScreenSavePwd(frmMain.hwnd)
    Else
        CorrectPassword = True
    End If
    
    'If it is correct, shut down, otherwise, reset
    If CorrectPassword = False Then
        'Reset everything
        
        'Diable the form or it will trap a MouseMove
        'Event again
        frmMain.Enabled = False
        
        'Reset cursor variables
        CurrentMouseX = 0
        CurrentMouseY = 0
        
        'Hide cursor
        ShowCursor False
        
        'Enable Screen Saver
        frmMain.Enabled = True
        
        'Restart events
        frmMain.Timer1.Enabled = True
        
        Exit Sub
    End If
    
    'Tell the system that the Screen Saver is shutting
    'down
    SystemParametersInfo SPI_SETSCREENSAVEACTIVE, 1, ByVal 0&, 0
    
    'Enable the system taskbar
    TaskbarhWnd = FindWindow("Shell_traywnd", "")
    
    If TaskbarhWnd <> 0 Then
        Taskbar = EnableWindow(TaskbarhWnd, 1)
    End If
        
    'Enable System Keys
    SystemParametersInfo SPI_SCREENSAVERRUNNING, False, TempBoolean, 0

    'Clear whatever forms you may have opened
    For Each FRM In Forms
        Unload FRM
    Next
    
    End
End Sub

Private Sub StartUp()
    'If App.PrevInstance = True Then End
End Sub
