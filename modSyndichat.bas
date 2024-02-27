Attribute VB_Name = "modSyndichat"
Option Explicit

Public gstrFind As String, gblnHasFocus As Boolean, gblnOptOut As Boolean

Public gstrUserID As String, gstrPassword As String, gstrNewText As String
Public gblnScrollbackForm As Boolean

' Used for launching web pages and e-mail.
Public Declare Function ShellExecute Lib "shell32.dll" _
    Alias "ShellExecuteA" _
    (ByVal hWnd As Long, ByVal lpOperation As String, _
     ByVal lpFile As String, ByVal lpParameters As String, _
     ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

' Used for checking state of a specific key on keyboard
Declare Function GetAsyncKeyState Lib "user32" _
    (ByVal vKey As Long) As Integer

' Used to test for application having focus or lost focus
Global Const GWL_WNDPROC = -4
Global Const WM_ACTIVATEAPP = &H1C
Public lpPrevWndProc As Long
Public gHW As Long

Public bln_frmChatHasFocus As Boolean, _
       bln_frmScrollbackHasFocus As Boolean

' Establishes a hook to capture messages for a window
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
    (ByVal hWnd As Long, ByVal nIndex As Long, _
    ByVal dwNewLong As Long) As Long

' Passes message to original window message handler
Declare Function CallWindowProc Lib "user32" Alias _
    "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
    ByVal hWnd As Long, ByVal Msg As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long
    

Public Function Messages_frmChat(ByVal hw As Long, ByVal uMsg As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long
    'Check for the ActivateApp message
    
    If uMsg = WM_ACTIVATEAPP Then
        'Check to see if Activating the application
        If wParam <> 0 Then
            'Application Received Focus
            gblnHasFocus = True
            frmChat.Caption = "SyndiChat"
        Else
            'Application Lost Focus
            gblnHasFocus = False
        End If
    End If
    
    'Pass message on to the original window message handler
    Messages_frmChat = CallWindowProc(lpPrevWndProc, hw, uMsg, wParam, _
        lParam)
        
End Function



Public Function FileExists(strFileName As String) As Boolean

'   Checks to see if a file exists

    If Dir$(strFileName) <> "" Then
        FileExists = True
    Else
        FileExists = False
    End If

End Function

Public Function ShiftKeyHeld() As Boolean

    Dim intKeyState As Integer

    intKeyState = GetAsyncKeyState(16)  ' check for shift key pressed
    
    If intKeyState = 0 Then
        ShiftKeyHeld = False
    Else
        ShiftKeyHeld = True
    End If

End Function
