VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmChat 
   Caption         =   "SyndiChat"
   ClientHeight    =   6525
   ClientLeft      =   2535
   ClientTop       =   3225
   ClientWidth     =   8940
   Icon            =   "frmChat.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6525
   ScaleWidth      =   8940
   Begin MSComDlg.CommonDialog dlgCommon 
      Left            =   8400
      Top             =   4920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer tmrInactivity 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   7320
      Top             =   5520
   End
   Begin VB.Timer tmrConnectTime 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7800
      Top             =   5520
   End
   Begin RichTextLib.RichTextBox rtfOutput 
      Height          =   4335
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   7646
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      TextRTF         =   $"frmChat.frx":0E42
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.StatusBar staStatus 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   6270
      Width           =   8940
      _ExtentX        =   15769
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1764
            MinWidth        =   1764
            Text            =   "Disconnect"
            TextSave        =   "Disconnect"
            Key             =   "Status"
            Object.ToolTipText     =   "Connect Status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1693
            MinWidth        =   176
            Text            =   "<Unknown>"
            TextSave        =   "<Unknown>"
            Key             =   "UserID"
            Object.ToolTipText     =   "User ID"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1296
            MinWidth        =   1288
            Text            =   "00:00:00"
            TextSave        =   "00:00:00"
            Key             =   "ConnectTime"
            Object.ToolTipText     =   "Connect Time"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "0 B"
            TextSave        =   "0 B"
            Key             =   "Skb"
            Object.ToolTipText     =   "Scrollback Buffer Size"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1032
            MinWidth        =   176
            Text            =   "0 char."
            TextSave        =   "0 char."
            Key             =   "Screen"
            Object.ToolTipText     =   "Line Length"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1244
            MinWidth        =   1235
            Text            =   "Anti-Idle"
            TextSave        =   "Anti-Idle"
            Key             =   "AntiIdle"
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtSend 
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      Locked          =   -1  'True
      MaxLength       =   32000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   5400
      Width           =   5775
   End
   Begin MSWinsockLib.Winsock tcpClient 
      Left            =   8400
      Top             =   5520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileConnect 
         Caption         =   "&Connect"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileDisconnect 
         Caption         =   "&Disconnect"
         Enabled         =   0   'False
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuFileHyphen1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSaveScrollback 
         Caption         =   "&Save Scrollback"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileHyphen2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileScrollback 
         Caption         =   "Open Scrollback Window"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuFileHyphen3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cu&t"
         Enabled         =   0   'False
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
         Enabled         =   0   'False
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
         Enabled         =   0   'False
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEditCopytoChat 
         Caption         =   "Copy to Chat &Window"
         Enabled         =   0   'False
         Shortcut        =   ^W
      End
      Begin VB.Menu mnuEditClear 
         Caption         =   "C&lear Clipboard"
      End
      Begin VB.Menu mnuEditHyphen1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditFind 
         Caption         =   "&Find"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuEditFindNext 
         Caption         =   "Find &Next"
         Enabled         =   0   'False
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuEditHyphen2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditSelectAll 
         Caption         =   "Select &All"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditHyphen3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditClearScrollback 
         Caption         =   "Clear &Scrollback"
         Shortcut        =   ^K
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuOptionSettings 
         Caption         =   "&User ID Settings"
      End
      Begin VB.Menu MnuOptionsFontSize 
         Caption         =   "&Font Size"
         Begin VB.Menu MnuOptionsFontSizeNum 
            Caption         =   "&6"
            Index           =   6
         End
         Begin VB.Menu MnuOptionsFontSizeNum 
            Caption         =   "&7"
            Index           =   7
         End
         Begin VB.Menu MnuOptionsFontSizeNum 
            Caption         =   "&8"
            Index           =   8
         End
         Begin VB.Menu MnuOptionsFontSizeNum 
            Caption         =   "&9"
            Index           =   9
         End
         Begin VB.Menu MnuOptionsFontSizeNum 
            Caption         =   "1&0"
            Index           =   10
         End
         Begin VB.Menu MnuOptionsFontSizeNum 
            Caption         =   "1&1"
            Index           =   11
         End
         Begin VB.Menu MnuOptionsFontSizeNum 
            Caption         =   "1&2"
            Index           =   12
         End
         Begin VB.Menu MnuOptionsFontSizeNum 
            Caption         =   "1&3"
            Index           =   13
         End
         Begin VB.Menu MnuOptionsFontSizeNum 
            Caption         =   "1&4"
            Index           =   14
         End
         Begin VB.Menu MnuOptionsFontSizeNum 
            Caption         =   "1&5"
            Index           =   15
         End
      End
      Begin VB.Menu mnuOptionsAntiIdle 
         Caption         =   "Anti-&Idle"
         Begin VB.Menu mnuOptionsAntiIdleNum 
            Caption         =   "Disabled"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuOptionsAntiIdleNum 
            Caption         =   "1 min."
            Index           =   1
         End
         Begin VB.Menu mnuOptionsAntiIdleNum 
            Caption         =   "2 mins."
            Index           =   2
         End
         Begin VB.Menu mnuOptionsAntiIdleNum 
            Caption         =   "3 mins"
            Index           =   3
         End
         Begin VB.Menu mnuOptionsAntiIdleNum 
            Caption         =   "4 mins."
            Index           =   4
         End
         Begin VB.Menu mnuOptionsAntiIdleNum 
            Caption         =   "5 mins."
            Index           =   5
         End
         Begin VB.Menu mnuOptionsAntiIdleNum 
            Caption         =   "7 mins."
            Index           =   7
         End
         Begin VB.Menu mnuOptionsAntiIdleNum 
            Caption         =   "10 mins"
            Index           =   10
         End
         Begin VB.Menu mnuOptionsAntiIdleNum 
            Caption         =   "15 mins."
            Index           =   15
         End
         Begin VB.Menu mnuOptionsAntiIdleNum 
            Caption         =   "20 mins"
            Index           =   20
         End
         Begin VB.Menu mnuOptionsAntiIdleNum 
            Caption         =   "30 mins."
            Index           =   30
         End
      End
      Begin VB.Menu mnuOptionsHyphen1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptionsDebug 
         Caption         =   "&Debug"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAreYouThere 
         Caption         =   "Are You &There?"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuHelpHyphen1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpVisitSyndicomm 
         Caption         =   "&Visit Syndicomm Online's Web Site..."
      End
      Begin VB.Menu mnuHelpHyphen2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpSyndiChat 
         Caption         =   "Visit Syndi&Chat's Web Site..."
      End
      Begin VB.Menu mneuHelpSupport 
         Caption         =   "&E-mail Technical Support..."
      End
      Begin VB.Menu mnuHelpHyphen3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About SyndiChat..."
      End
   End
End
Attribute VB_Name = "frmChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lngBegSearch As Long, intDebugFile As Integer
Dim blnDebug As Boolean, blnFirstData As Boolean, _
    blnLoggedIn As Boolean
Dim bytFontSize As Byte
Dim bytFontWidth(6 To 15) As Byte
Dim dtmConnectTime As Date
Dim strAppPath As String
Dim intAntiIdle As Integer

Private Declare Function LockWindowUpdate Lib "user32" _
    (ByVal hWnd As Long) As Long
    


Private Sub Hook()
    
    'Establish a hook to capture messages to this window
    lpPrevWndProc = SetWindowLong(gHW, GWL_WNDPROC, _
        AddressOf Messages_frmChat)

End Sub




Private Sub unHook()
    
    Dim temp As Long
    
    'Reset the message handler for this window
    temp = SetWindowLong(gHW, GWL_WNDPROC, lpPrevWndProc)
    
End Sub
Private Sub Form_Load()

    blnLoggedIn = False
    mnuOptionsAntiIdleNum(0).Checked = True
    intAntiIdle = 0
    staStatus.Panels("AntiIdle").Visible = False
    gblnScrollbackForm = False
    
    strAppPath = App.Path
    
    If Right(strAppPath, 1) = "\" Then
        strAppPath = Left(strAppPath, Len(strAppPath) - 1)
    End If
    
    OpenCloseDebug
    
    Me.Move (Screen.Width - Me.Width) \ 2, (Screen.Height - Me.Height) \ 2

    bytFontWidth(6) = 5
    bytFontWidth(7) = 5
    bytFontWidth(8) = 7
    bytFontWidth(9) = 7
    bytFontWidth(10) = 8
    bytFontWidth(11) = 9
    bytFontWidth(12) = 10
    bytFontWidth(13) = 10
    bytFontWidth(14) = 11
    bytFontWidth(15) = 12
    
    bytFontSize = 8
    MnuOptionsFontSizeNum(bytFontSize).Checked = True
    
    ' Get default settings
    LoadINI
    
    ' If ID not set, ask for it
    
    If gstrUserID = "" And gstrPassword = "" And gblnOptOut = False Then
        frmSettings.Show vbModal
    End If
    
    'Store handle to this form's window
    gHW = Me.hWnd

    'Call procedure to beging capturing messages for this window
    Hook
    
    tcpClient.RemoteHost = "syndicomm.com"
    tcpClient.RemotePort = 23

    staStatus.Panels("Status").Text = "Disconnect"
    
    TCPConnect
    
    If Clipboard.GetText() = "" Or tcpClient.State = sckClosed Then
        mnuEditPaste.Enabled = False
    Else
        mnuEditPaste.Enabled = True
    End If

End Sub


Private Sub Form_Resize()

    Static sintDefaultSendHeight As Integer
    
    If frmChat.ScaleHeight <> 0 And frmChat.ScaleWidth <> 0 Then
    
        If sintDefaultSendHeight = 0 Then
            sintDefaultSendHeight = txtSend.Height
        End If
    
        txtSend.Left = frmChat.ScaleLeft
        txtSend.Top = frmChat.ScaleHeight - staStatus.Height - sintDefaultSendHeight
        txtSend.Width = frmChat.ScaleWidth
        txtSend.Height = sintDefaultSendHeight
    
        rtfOutput.Left = frmChat.ScaleLeft
        rtfOutput.Width = frmChat.ScaleWidth
        
        If frmChat.ScaleHeight >= staStatus.Height + txtSend.Height Then
            rtfOutput.Height = frmChat.ScaleHeight - staStatus.Height - txtSend.Height
        End If
        
        Set_Window
    
    End If

End Sub


Private Sub Form_Unload(Cancel As Integer)

    Dim intStatus As Integer
    
    If tcpClient.State <> sckClosed Then
    
        intStatus = MsgBox("Do you really wish to exit?", vbOKCancel + vbExclamation, "Warning")
    
        If intStatus = vbOK Then
            Close_Connection
        Else
            Cancel = 1
        End If
        
    End If
    
    If blnDebug = True Then
        WriteDebug "*** Closing debug ..."
        Close #intDebugFile
    End If
    
    ' Unload frmScrollback

    'Call procedure to stop intercepting the messages for this window
    unHook

    ' Create INI and save preferences
    SaveINI

End Sub




Private Sub mneuHelpSupport_Click()
    
    Const SW_SHOWNORMAL = 1
    
    ShellExecute 0, vbNullString, "mailto:mark@syndicomm.com?subject=SyndiChat Technical Support Query", vbNullString, vbNullString, SW_SHOWNORMAL

End Sub

Private Sub mnuEdit_Click()
    
    If Clipboard.GetText() = "" Or tcpClient.State = sckClosed Then
        mnuEditPaste.Enabled = False
        mnuEditClear.Enabled = False
    Else
        mnuEditPaste.Enabled = True
        mnuEditClear.Enabled = True
    End If

End Sub

Private Sub mnuEditClear_Click()

    Clipboard.Clear

End Sub

Private Sub mnuEditClearScrollback_Click()

    rtfOutput.Text = ""
    ShowSkbSize

End Sub

Private Sub mnuEditCopy_Click()

    CopyText

End Sub

Private Sub mnuEditCopytoChat_Click()

    CopyText
    PasteText

End Sub

Private Sub mnuEditCut_Click()

    CopyText
    
    txtSend.SelText = ""

End Sub

Private Sub mnuEditFind_Click()

    frmFind.Show vbModal
    Set frmFind = Nothing

    If rtfOutput.SelLength <> 0 Then
        rtfOutput.SetFocus
    End If

End Sub

Private Sub mnuEditFindNext_Click()

    Dim lngPosFound As Long, intRetVal As Integer
    
    If lngBegSearch <> 1 Then
        lngBegSearch = rtfOutput.SelStart + 2
    End If
    
    lngPosFound = InStr(lngBegSearch, rtfOutput.Text, gstrFind, vbTextCompare)

    If lngPosFound = 0 Then
        intRetVal = MsgBox("End of scrollback.", vbOKOnly + vbInformation, "Find Next")
        txtSend.SetFocus
        lngBegSearch = 1
    Else
        frmChat.rtfOutput.SelStart = lngPosFound - 1
        frmChat.rtfOutput.SelLength = Len(gstrFind)
        rtfOutput.SetFocus
        lngBegSearch = 0
    End If

End Sub


Private Sub mnuEditPaste_Click()

    PasteText

End Sub


Private Sub mnuEditSelectAll_Click()

    If bln_frmChatHasFocus = False Then
        rtfOutput.SelStart = 0
        rtfOutput.SelLength = Len(rtfOutput.Text)
        rtfOutput.SetFocus
    
    Else
        txtSend.SelStart = 0
        txtSend.SelLength = Len(txtSend.Text)
        txtSend.SetFocus
    
    End If
    
    Check4SelectedText

End Sub

Private Sub mnuFileConnect_Click()
    
    TCPConnect

End Sub

Private Sub mnuFileDisconnect_Click()

    Close_Connection

End Sub

Private Sub mnuFileExit_Click()

    Unload Me

End Sub

Private Sub mnuFileSaveScrollback_Click()
    
    Dim intFileHandle As Integer
    Static strFileName As String
    
    'Set error conditoin for cancel
    On Error GoTo mnuFileSaveScrollback_Click_Err
    
    'Set default file name
    If strFileName = "" Then
        strFileName = "*.txt"
    End If
    
    'Set common dialogue options
    dlgCommon.DialogTitle = "Save Scrollback As"
    dlgCommon.CancelError = True
    dlgCommon.Flags = cdlOFNOverwritePrompt + cdlOFNPathMustExist
    dlgCommon.FileName = strFileName
    dlgCommon.Filter = "Text Documents (*.txt)|*.txt|All Files|*.*"
    'Show save dialogue
    dlgCommon.ShowSave
    
    'Get free file handle
    intFileHandle = FreeFile
    
    'Save to file
    Open dlgCommon.FileName For Output As #intFileHandle
    Print #intFileHandle, rtfOutput.Text
    Close #intFileHandle
    
    Exit Sub
    
mnuFileSaveScrollback_Click_Err:

    'Cancel is a valid error
    If Err.Number <> cdlCancel Then
        MsgBox Err.Number & " - " & Err.Description
    End If

End Sub

Private Sub mnuFileScrollback_Click()

    mnuFileScrollback.Enabled = False
    frmScrollback.Show

End Sub

Private Sub mnuHelpAbout_Click()
    
    frmAbout.Show vbModal, Me
    Set frmAbout = Nothing

End Sub

Private Sub mnuHelpAreYouThere_Click()

    tcpClient.SendData Chr(255) & Chr(246)
    WriteDebug "Are you there?"

End Sub

Private Sub mnuHelpSyndiChat_Click()
    
    Const SW_SHOWNORMAL = 1
    
    ShellExecute 0, vbNullString, "http://www.syndicomm.com/~Mark/SyndiChat/", vbNullString, vbNullString, SW_SHOWNORMAL

End Sub

Private Sub mnuHelpVisitSyndicomm_Click()

    Const SW_SHOWNORMAL = 1
    
    ShellExecute 0, vbNullString, "http://www.syndicomm.com", vbNullString, vbNullString, SW_SHOWNORMAL

End Sub

Private Sub mnuOptionsAntiIdleNum_Click(Index As Integer)

    mnuOptionsAntiIdleNum(intAntiIdle).Checked = False
    mnuOptionsAntiIdleNum(Index).Checked = True
    
    intAntiIdle = Index
    
    If intAntiIdle = 0 Then
        tmrInactivity.Enabled = False
        staStatus.Panels("AntiIdle").Visible = False
    Else
        tmrInactivity.Enabled = True
        staStatus.Panels("AntiIdle").Visible = True
    End If

End Sub

Private Sub mnuOptionSettings_Click()

    frmSettings.Show vbModal
    Set frmSettings = Nothing

End Sub

Private Sub MnuOptionsFontSizeNum_Click(Index As Integer)

    MnuOptionsFontSizeNum(bytFontSize).Checked = False
    MnuOptionsFontSizeNum(Index).Checked = True
    
    bytFontSize = Index
    
    Set_Window

End Sub


Private Sub rtfOutput_Click()

    Check4SelectedText

End Sub


Private Sub rtfOutput_DblClick()
    
    Check4SelectedText
    
End Sub


Private Sub rtfOutput_GotFocus()
    
    bln_frmChatHasFocus = False
    
    txtSend.SelLength = 0
    
    Check4SelectedText

    If Clipboard.GetText() = "" Or tcpClient.State = sckClosed Then
        mnuEditPaste.Enabled = False
    Else
        mnuEditPaste.Enabled = True
    End If

End Sub

Private Sub rtfOutput_KeyPress(KeyAscii As Integer)

    If KeyAscii >= 32 And KeyAscii <= 127 Then
        txtSend.Text = txtSend.Text & Chr(KeyAscii)
        txtSend.SelStart = Len(txtSend.Text)
        txtSend.SelLength = 0
        txtSend.SetFocus
    End If
    
    If KeyAscii = 13 Then
        txtSend.SetFocus
    End If

End Sub

Private Sub rtfOutput_KeyUp(KeyCode As Integer, Shift As Integer)

    Check4SelectedText

End Sub

Private Sub tcpClient_Close()

    Dim intStatus As Integer
    
    Close_Connection
    
    intStatus = MsgBox("Disconnected by host.", vbOKOnly + vbInformation, "Disconnect")

End Sub

Private Sub tcpClient_DataArrival(ByVal bytesTotal As Long)

    Dim strTCPData As String, strTemp As String, strIACData As String * 2, strUnit As String * 2
    Dim intI As Integer, strChar As String * 1, intChar As Integer
    Dim lngScrollBack As Long, lngSelStart As Long, lngSelLength As Long
    Dim sngScrollback As Single
    
    If blnFirstData = True Then
        StatusMsg "Connected to IP " & tcpClient.RemoteHostIP
        blnFirstData = False
    End If
        
    tcpClient.GetData strTCPData
    
    WriteDebug " Received " & bytesTotal & " bytes."
    
    For intI = 1 To bytesTotal
    
        strChar = Mid(strTCPData, intI, 1)
        intChar = Asc(strChar)
               
        Select Case intChar
    
            Case 0 To 6                                         ' Ctrl characters
            
            Case 7                                              ' Beep
                Beep
                
            Case 9, 11, 12                                      ' Ctrl characters
            
            Case 14 To 31                                       ' Ctrl characters
            
            Case 128 To 254                                     ' High bit characters
            
            Case 255                                            ' IAC
                strIACData = Mid(strTCPData, intI + 1, 2)       ' IAC data next 2 bytes.
                intI = Process_IAC(intI, strIACData)            ' Process IAC
                   
            Case Else                                           ' Writeable data
                strTemp = strTemp & strChar                     ' Add valid character to output string
        
        End Select
        
    Next intI
        
    If frmChat.WindowState <> vbMinimized And _
        gblnHasFocus = True Then                                ' Prevent desktop refresh when minimized
        LockWindowUpdate rtfOutput.hWnd                         ' Lock output window to prevent window repositioning
        frmChat.Caption = "SyndiChat"
    Else
        frmChat.Caption = "SyndiChat **"                        ' If not front window, indicate update in form caption eith '**'
    End If
        
    rtfOutput.Text = rtfOutput.Text & strTemp                   ' Append new data to output window
    
    rtfOutput.SelStart = Len(rtfOutput.Text)                    ' Reposition cursor to end of window
    rtfOutput.SelLength = 0                                     ' Select length 0
    
    If frmChat.WindowState <> vbMinimized And _
        gblnHasFocus = True Then
        LockWindowUpdate 0                                      ' Unlock window to update
    End If
    
    ShowSkbSize                                                 ' Update scrollback size display
    
    If gblnScrollbackForm = True Then
        gstrNewText = gstrNewText & strTemp
    Else
        gstrNewText = Empty
    End If
    
    If blnLoggedIn = False Then
        If Trim(Right(rtfOutput.Text, 7)) = "login:" And _
            gstrUserID <> "" Then
            tcpClient.SendData gstrUserID & vbCr
            If gstrPassword = "" Then
                blnLoggedIn = True
            End If
        End If
        If Trim(Right(rtfOutput.Text, 10)) = "Password:" And _
            gstrPassword <> "" Then
            tcpClient.SendData gstrPassword & vbCr
            blnLoggedIn = True
        End If
    End If
                                
End Sub

Private Sub tcpClient_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

    StatusMsg "Winsock Err: " & Number & "-" & Description
    
    Close_Connection

End Sub


Public Sub StatusMsg(strMsg As String)

    rtfOutput.Text = rtfOutput.Text & vbNewLine & "Msg : " & strMsg
    WriteDebug strMsg
    rtfOutput.SelStart = Len(rtfOutput)
    rtfOutput.SelLength = 0

End Sub



Private Sub Timer1_Timer()

End Sub

Private Sub tmrConnectTime_Timer()

        
    staStatus.Panels("ConnectTime").Text = Format(Now - dtmConnectTime, "[d:]hh:mm:ss")
    

End Sub


Private Sub tmrInactivity_Timer()
    
    Static intMinutes As Integer
    
    intMinutes = intMinutes + 1
    
    If intMinutes >= intAntiIdle And _
        staStatus.Panels("AntiIdle").Enabled = True Then
'       tcpClient.SendData Chr(255) & Chr(246)  'Send "Are You There?"
        tcpClient.SendData Chr(255) & Chr(241)  'Send NOP
        intMinutes = 0
        
    End If

End Sub


Private Sub txtSend_Click()

    Check4SelectedText

End Sub

Private Sub txtSend_DblClick()

    Check4SelectedText

End Sub


Private Sub txtSend_GotFocus()

    bln_frmChatHasFocus = True

    rtfOutput.SelLength = 0
    
    Check4SelectedText

    If Clipboard.GetText() = "" Or tcpClient.State = sckClosed Then
        mnuEditPaste.Enabled = False
    Else
        mnuEditPaste.Enabled = True
    End If


End Sub

Private Sub txtSend_KeyPress(KeyAscii As Integer)

    Dim intNumOfSubSegments As Integer, intI As Integer
    Dim intPosnFound As Integer, intPosnStart As Integer
    Dim strSegment As String, strSubSegment As String
    
    ' Checks for either:
    ' - Shift key being held down or
    ' - Not a <cr> press (ASCII 13) or
    ' - Not connected to Syndicomm
    '   Then exit
    
    If ShiftKeyHeld Or _
        KeyAscii <> 13 Or _
        tcpClient.State <> sckConnected Then
        Exit Sub
    End If
    
    ' Parses the chat window (txtSend.Text) into Syndicomm's maximum 400 bytes at a time.
    ' strSegment is a <cr> delimited chunk of txtSend.Text
    ' strSubSegment is a mamimum 399 byte section (plus <cr> = 400) of strSegment to send.
    
    intPosnStart = 1                                'Start <cr> search at beginning
    
    Do
        'Find position of embedded <cr> in txtSend.Text
        intPosnFound = InStr(intPosnStart, txtSend.Text, vbCr, vbBinaryCompare)
        
        'If no embedded <cr> found then put pointer to end of txtSend.Text
        If intPosnFound = 0 Then
            intPosnFound = Len(txtSend.Text) + 1
        End If
        
        'Extract into strSegment
        strSegment = Mid(txtSend.Text, intPosnStart, intPosnFound - intPosnStart)
        
        'Figure out how many max 399 byte segments this will be
        intNumOfSubSegments = (Len(strSegment) \ 399) + 1
        
        'Do for each segment number
        For intI = 1 To intNumOfSubSegments
            'Extract maximum 399 byte subsegment
            strSubSegment = Mid(strSegment, 1 + (399 * (intI - 1)), 399)
            'Send subsegment
            tcpClient.SendData strSubSegment & vbCr
            'Give the system time to process the data sent
            DoEvents
            'Report for debuging purposes
            WriteDebug "Sending " & Len(strSubSegment) + 1 & " bytes. (Segment " & _
                        intI & " of " & intNumOfSubSegments & ")"
        'Next segment number
        Next
        
        'Move search start pointer to position after last embedded <cr> found.
        intPosnStart = intPosnFound + 1
    
    'Loop until the search start pointer is beyond the end of txtSend
    Loop Until intPosnStart > Len(txtSend.Text)

    'Clear chat box and clear the <cr> entered.
    txtSend.Text = Empty
    KeyAscii = 0
    
End Sub

Public Sub Close_Connection()

    StatusMsg "Closing...."
    tcpClient.Close
    
    staStatus.Panels("Status").Text = "Disconnect"
    tmrConnectTime.Enabled = False
    mnuFileConnect.Enabled = True
    mnuFileDisconnect.Enabled = False
    txtSend.Locked = True
    mnuHelpAreYouThere.Enabled = False
    staStatus.Panels("AntiIdle").Enabled = False
    blnLoggedIn = False

End Sub

Public Sub TCPConnect()
    
    StatusMsg "Connecting...."
    tcpClient.Connect
    
    blnFirstData = True
    
    staStatus.Panels("Status").Text = "Connect"
    dtmConnectTime = Now
    tmrConnectTime.Enabled = True
    mnuFileDisconnect.Enabled = True
    mnuFileConnect.Enabled = False
    staStatus.Panels("AntiIdle").Enabled = True
    txtSend.Locked = False
    mnuHelpAreYouThere.Enabled = True
    
    If gstrUserID <> "" Then
        staStatus.Panels("UserID").Text = gstrUserID
        staStatus.Panels("UserID").Visible = True
    Else
        staStatus.Panels("UserID").Visible = False
    End If

End Sub

Public Sub Check4SelectedText()

    If txtSend.SelText = "" Then
        mnuEditCut = False
    Else
        mnuEditCut = True
    End If

    If rtfOutput.SelText = "" And txtSend.SelText = "" Then
        mnuEditCopy.Enabled = False
        mnuEditCopytoChat.Enabled = False
    
    Else
        mnuEditCopy.Enabled = True
        
        If tcpClient.State = sckConnected And rtfOutput.SelText <> "" Then
            mnuEditCopytoChat.Enabled = True
        
        Else
            mnuEditCopytoChat.Enabled = False
        
        End If
    
    End If

End Sub

Public Sub CopyText()
    
    Clipboard.Clear
    
    If bln_frmChatHasFocus = True Then
        Clipboard.SetText txtSend.SelText
        
    Else
        Clipboard.SetText rtfOutput.SelText
    
    End If

End Sub

Public Sub PasteText()

    txtSend.SelText = Clipboard.GetText()
    txtSend.SetFocus

End Sub

Public Function Process_IAC(ByVal intPosition, ByVal strIACData) As Integer

    Dim strMessage As String
    
    strMessage = "Server : IAC "

    If Asc(Mid(strIACData, 1, 1)) = 251 Then                                ' WILL
        
        strMessage = strMessage & "WILL "
        
        If Asc(Mid(strIACData, 2, 1)) = 1 Then                              ' Echo
            strMessage = strMessage & "ECHO (RFC 857)"
            WriteDebug strMessage
            tcpClient.SendData Chr(255) & Chr(253) & Chr(1)                 'IAC DO
            WriteDebug "Client : IAC DO ECHO (RFC 857)"
        
        ElseIf Asc(Mid(strIACData, 2, 1)) = 3 Then                          ' Suppress Go Ahead
            strMessage = strMessage & "SUPPRESS-GO-AHEAD (RFC 858)"
            WriteDebug strMessage
            tcpClient.SendData Chr(255) & Chr(253) & Chr(3)                 'IAC DO
            WriteDebug "Client : IAC DO SUPPRESS-GO-AHEAD (RFC 858)"
        
        ElseIf Asc(Mid(strIACData, 2, 1)) = 5 Then                          ' Status
            strMessage = strMessage & "STATUS (RFC 651)"
            WriteDebug strMessage
            tcpClient.SendData Chr(255) & Chr(253) & Chr(3)                 'IAC DO
            WriteDebug "Client : IAC DO STATUS (RFC 651)"
        
        Else
        
           WriteDebug strMessage & Val(Asc(Mid(strIACData, 2, 1)))
        
        End If
        
        Process_IAC = intPosition + 2
    
    ElseIf Asc(Mid(strIACData, 1, 1)) = 253 Then                            ' DO
    
        strMessage = strMessage & "DO "
        
        If Asc(Mid(strIACData, 2, 1)) = 1 Then                              ' Echo
            strMessage = strMessage & "ECHO (RFC 857)"
            WriteDebug strMessage
            tcpClient.SendData Chr(255) & Chr(251) & Chr(1)                 'IAC WILL
            WriteDebug "Client : IAC WILL ECHO (RFC 857)"
        
        ElseIf Asc(Mid(strIACData, 2, 1)) = 24 Then                         ' Terminal Type
            strMessage = strMessage & "TERMINAL-TYPE (RFC 930)"
            WriteDebug strMessage
            tcpClient.SendData Chr(255) & Chr(252) & Chr(24)                'IAC WON'T
            WriteDebug "Client : IAC WON'T TERMINAL-TYPE (RFC 930)"
        
        ElseIf Asc(Mid(strIACData, 2, 1)) = 31 Then                         ' NAWS
            strMessage = strMessage & "NAWS (RFC 1073)"
            WriteDebug strMessage
            tcpClient.SendData Chr(255) & Chr(251) & Chr(31)                'IAC WILL
            WriteDebug "Client : IAC WILL NAWS (RFC 1073)"
            Set_Window
        Else
            WriteDebug strMessage & Val(Asc(Mid(strIACData, 2, 1)))
            tcpClient.SendData Chr(255) & Chr(252) & Mid(strIACData, 2, 1)  ' No to everything else too
            WriteDebug "Client : IAC WON'T " & Val(Asc(Mid(strIACData, 2, 1)))
        
        End If
    
        Process_IAC = intPosition + 2
    
    ElseIf Asc(Mid(strIACData, 1, 1)) = 254 Then                            ' DON'T
    
        strMessage = strMessage & "DON'T "
        
        WriteDebug strMessage & Val(Asc(Mid(strIACData, 2, 1)))
        tcpClient.SendData Chr(255) & Chr(252) & Mid(strIACData, 2, 1)       'WON'T
        WriteDebug "Client : IAC WON'T " & Val(Asc(Mid(strIACData, 2, 1)))
    
        Process_IAC = intPosition + 2
    
    Else
        WriteDebug strMessage & Val(Asc(Mid(strIACData, 1, 1)))
        
        Process_IAC = intPosition
    
    End If

End Function

Public Sub WriteDebug(ByVal strData As String)

'   Write to debug file.

    If blnDebug = True Then
        Print #intDebugFile, Format(Now, "yyyy-mm-dd hh:mm:ss") & " - " & strData
    End If

End Sub

Public Sub Set_Window()

    Dim intLineSize As Integer, bytLineSize1 As Byte, bytLineSize2 As Byte
    
    If bytFontSize <> 0 Then
        rtfOutput.Font.Size = bytFontSize
        txtSend.FontSize = bytFontSize
    End If
    
    rtfOutput.SelStart = Len(rtfOutput.Text)
    rtfOutput.SelLength = 0
    
    If tcpClient.State <> sckConnected Then
        Exit Sub
    End If
    
    intLineSize = ((rtfOutput.Width * 0.067) \ bytFontWidth(bytFontSize)) - 5
    
    If intLineSize < 20 Then
        intLineSize = 20
    End If
    
    staStatus.Panels("Screen").Text = intLineSize & " char."        ' Display line size in status bar.
    
    bytLineSize1 = intLineSize \ 256
    bytLineSize2 = intLineSize Mod 256
    
    tcpClient.SendData Chr(255) & Chr(250) & Chr(31) & _
             Chr(bytLineSize1) & Chr(bytLineSize2) & Chr(0) & Chr(24) & _
             Chr(255) & Chr(240)                                            'IAC SB NAWS xx xx 0 24 IAC SE
             
    WriteDebug "Client : IAC SB NAWS " & bytLineSize1 & " " & bytLineSize2 & " 0 24 IAC SE (RFC 1073)"

End Sub

Private Sub txtSend_KeyUp(KeyCode As Integer, Shift As Integer)

    Check4SelectedText

End Sub



Public Sub ShowSkbSize()

    Dim lngScrollBack As Long, sngScrollback As Single, strUnit As String

    lngScrollBack = Len(rtfOutput.Text)
       
    Select Case lngScrollBack
    
        Case 0 To 1023
            sngScrollback = lngScrollBack
            strUnit = "B"
            
        Case 1024 To 1048575
            sngScrollback = Format(lngScrollBack / 1024, "0.0")
            strUnit = "KB"
            
        Case Else
            sngScrollback = Format(lngScrollBack / 1048576, "0.00")
            strUnit = "MB"
    
    End Select
    
    staStatus.Panels("Skb").Text = sngScrollback & " " & strUnit

End Sub

Public Sub OpenCloseDebug()
    
    If mnuOptionsDebug.Checked = True Then
        blnDebug = True
        intDebugFile = FreeFile
        Open strAppPath & "\Debug.log" For Output As #intDebugFile
        WriteDebug "*** Starting debug...."
    Else
        blnDebug = False
        If intDebugFile <> 0 Then                   ' Not 0 when debug has been previously opened
            WriteDebug "*** Closing debug ..."
            Close #intDebugFile
            intDebugFile = 0                        ' Zero file handle for a flag
        End If
    End If

End Sub

Public Sub SaveINI()

    Dim intFileHandle As Integer, intKey As Integer, intI As Integer
    Dim strPassword As String
    
    If ShiftKeyHeld Then                    ' Hold shift key to bypass saving INI file
        Exit Sub
    End If
    
    If Not gblnOptOut And Len(gstrUserID) <> 0 Then  ' encrypt if he is using auto-login
        intKey = Asc(Right(Trim(gstrUserID), 1))
        strPassword = ""
    
        For intI = Len(gstrPassword) To 1 Step -1
            strPassword = strPassword & Format(Asc(Mid(gstrPassword, intI, 1)) Xor intKey + intI, "000")
        Next
    End If
    
    intFileHandle = FreeFile
    Open strAppPath & "\SyndiChat.ini" For Output As #intFileHandle
    
    Print #intFileHandle, "***** SyndiChat INI ***** "
    Print #intFileHandle,
    Print #intFileHandle, "*** " & Format(Now, "yyyy-mm-dd hh:mm:ss")
    Print #intFileHandle,
    Print #intFileHandle, "*** Screen Settings ***"
    Print #intFileHandle, "frmChat.Width: " & frmChat.Width
    Print #intFileHandle, "frmChat.Height: " & frmChat.Height
    Print #intFileHandle, "frmChat.Top: " & frmChat.Top
    Print #intFileHandle, "frmChat.Left: " & frmChat.Left
    Print #intFileHandle, "rtfOutput.Font.Size: " & rtfOutput.Font.Size
    Print #intFileHandle, "txtSend.FontSize: " & txtSend.FontSize
    Print #intFileHandle,
    Print #intFileHandle, "*** ID and Password ***"
    If Not gblnOptOut Then
        Print #intFileHandle, "gstrUserID: " & gstrUserID
        Print #intFileHandle, "gstrPassword-e: " & strPassword
    End If
    Print #intFileHandle, "gblnOptOut: " & gblnOptOut
    Print #intFileHandle,
    Print #intFileHandle, "*** Misc Settings ***"
    Print #intFileHandle, "intAntiIdle: " & intAntiIdle
    Close #intFileHandle
    
End Sub

Public Sub LoadINI()

    Dim intFileHandle As Integer, intDelimiter As Integer, intKey As Integer, intI As Integer
    Dim intKeyOffset As Integer
    Dim strLine As String, strObject As String, strSetting As String, strPassword As String
    
    If ShiftKeyHeld Then                ' Shift key to bypass loading INI file
        Exit Sub
    End If
    
    strPassword = ""
    
'   Check for INI file
    
    If FileExists(strAppPath & "\SyndiChat.ini") = True Then

        intFileHandle = FreeFile
    
        Open strAppPath & "\SyndiChat.ini" For Input As #intFileHandle
    
        Do While Not EOF(intFileHandle)
    
            Line Input #intFileHandle, strLine
            
            If Left(strLine, 3) <> "***" And _
               Len(strLine) <> 0 Then           ' *** or blank = Comment line
            
                intDelimiter = InStr(1, strLine, ":")   ' Find delimiter position
                
                If intDelimiter <> 0 Then               ' 0 = didn't find a delimiter
                
                    strObject = Trim(Left(strLine, intDelimiter - 1))
                    strSetting = Trim(Right(strLine, Len(strLine) - intDelimiter))
                
                    Select Case strObject
                    
                        Case "frmChat.Width"
                            frmChat.Width = Val(strSetting)
                        Case "frmChat.Height"
                            frmChat.Height = Val(strSetting)
                        Case "frmChat.Top"
                            frmChat.Top = Val(strSetting)
                        Case "frmChat.Left"
                            frmChat.Left = Val(strSetting)
                        Case "rtfOutput.Font.Size"
                            MnuOptionsFontSizeNum_Click (Round(Val(strSetting), 0))
                        Case "gstrUserID"
                            gstrUserID = strSetting
                        Case "gstrPassword"             ' Obsolete but here for backward compatability
                            gstrPassword = strSetting
                        Case "gstrPassword-e"
                            strPassword = strSetting
                        Case "gblnOptOut"
                            If strSetting = "True" Then
                                gblnOptOut = True
                            Else
                                gblnOptOut = False
                            End If
                        Case "intAntiIdle"
                            mnuOptionsAntiIdleNum_Click (Round(Val(strSetting), 0))
                        Case Else
                    
                    End Select
                
                End If
            
            End If
            
        Loop
    
        Close #intFileHandle
        
        If Len(strPassword) > 0 Then        'Decrypt password
        
            intKey = Asc(Right(Trim(gstrUserID), 1))
            gstrPassword = ""
            
            For intI = 1 To Len(strPassword) \ 3
            
                intKeyOffset = intKey + Abs(intI - (Len(strPassword) \ 3) - 1)
                gstrPassword = Chr(Val(Mid(strPassword, (intI * 3) - 2, 3)) Xor intKeyOffset) & gstrPassword
            
            Next intI
            
        End If
        
    End If

End Sub
