VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmScrollback 
   Caption         =   "SyndiChat - Scrollback Window"
   ClientHeight    =   6255
   ClientLeft      =   1845
   ClientTop       =   2685
   ClientWidth     =   9360
   Icon            =   "frmScrollback.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6255
   ScaleWidth      =   9360
   Begin MSComDlg.CommonDialog dlgCommon 
      Left            =   8760
      Top             =   5640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox rtfScrollback 
      Height          =   3615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   6376
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      TextRTF         =   $"frmScrollback.frx":0E42
   End
   Begin VB.Menu mnuFIle 
      Caption         =   "&File"
      Begin VB.Menu mnuFileSaveScrollback 
         Caption         =   "&Save Scrollback"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFIleHyphen1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Close Window"
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
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
         Enabled         =   0   'False
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEditChat 
         Caption         =   "Copy to Chat &Window"
         Shortcut        =   ^W
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
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuEditHyphen2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditSelectAll 
         Caption         =   "Select &All"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuRefreshAll 
         Caption         =   "Refresh &All"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuRefreshUpdated 
         Caption         =   "Refresh &Updated"
         Shortcut        =   +{F5}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About SyndiChat..."
      End
   End
End
Attribute VB_Name = "frmScrollback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()

    gblnScrollbackForm = True
    
    mnuEditCopy.Enabled = False
    mnuEditChat.Enabled = False
    mnuEditFindNext.Enabled = False
    
    rtfScrollback.Text = frmChat.rtfOutput.Text

    frmScrollback.Top = frmChat.Top + 400
    frmScrollback.Left = frmChat.Left + 400
    frmScrollback.Width = frmChat.Width
    frmScrollback.Height = frmChat.rtfOutput.Height + 800
    
    rtfScrollback.Font.Name = frmChat.rtfOutput.Font.Name
    rtfScrollback.Font.Size = frmChat.rtfOutput.Font.Size
    
    rtfScrollback.SelStart = Len(rtfScrollback.Text)
    rtfScrollback.SelLength = 0
    
End Sub


Private Sub Form_Resize()

    rtfScrollback.Height = frmScrollback.ScaleHeight
    rtfScrollback.Width = frmScrollback.ScaleWidth

End Sub


Private Sub Form_Unload(Cancel As Integer)

    frmChat.mnuFileScrollback.Enabled = True
    gblnScrollbackForm = False
    rtfScrollback.Text = ""             ' Just to make sure it isn't taking up memory

End Sub

Private Sub mnuEditChat_Click()

    mnuEditCopy_Click
    
    With frmChat
        .txtSend.SelText = Clipboard.GetText()
        .txtSend.SetFocus
    End With

End Sub

Private Sub mnuEditCopy_Click()

    Clipboard.Clear
    
    Clipboard.SetText rtfScrollback.SelText

End Sub

Private Sub mnuEditFind_Click()

    frmFind2.Show vbModal
    Set frmFind2 = Nothing

End Sub

Private Sub mnuEditFindNext_Click()

    Dim lngPosFound As Long, lngBegSearch As Long
    Dim intRetVal As Integer
    
    lngBegSearch = frmScrollback.rtfScrollback.SelStart + 2
    
    lngPosFound = InStr(lngBegSearch, frmScrollback.rtfScrollback.Text, gstrFind, vbTextCompare)

    If lngPosFound = 0 Then
        intRetVal = MsgBox("End of scrollback.", vbOKOnly + vbInformation, "Find Next")
        frmScrollback.rtfScrollback.SelStart = 0
        frmScrollback.rtfScrollback.SelLength = 0
    Else
        frmScrollback.rtfScrollback.SelStart = lngPosFound - 1
        frmScrollback.rtfScrollback.SelLength = Len(gstrFind)
    End If

End Sub

Private Sub mnuEditSelectAll_Click()

    With rtfScrollback
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
        rtfScrollback_Click
    End With

End Sub

Private Sub mnuFileClose_Click()

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
    Print #intFileHandle, rtfScrollback.Text
    Close #intFileHandle
    
    Exit Sub
    
mnuFileSaveScrollback_Click_Err:

    'Cancel is a valid error
    If Err.Number <> cdlCancel Then
        MsgBox Err.Number & " - " & Err.Description
    End If

End Sub

Private Sub mnuHelpAbout_Click()

    frmAbout.Show vbModal, Me
    Set frmAbout = Nothing

End Sub

Private Sub mnuRefreshAll_Click()

    rtfScrollback.Text = frmChat.rtfOutput.Text
    rtfScrollback.SelStart = Len(rtfScrollback.Text)
    rtfScrollback.SelLength = 0
    gstrNewText = Empty

End Sub

Private Sub mnuRefreshUpdated_Click()

    rtfScrollback.Text = gstrNewText
    rtfScrollback.SelStart = Len(rtfScrollback.Text)
    rtfScrollback.SelLength = 0
    gstrNewText = Empty

End Sub


Private Sub rtfScrollback_Click()

    If rtfScrollback.SelLength = 0 Then      'No selected text, disable copy and copy to chat.
        mnuEditCopy.Enabled = False
        mnuEditChat.Enabled = False
        Exit Sub                            'Exit sub since nothing else matters.
    End If

    mnuEditCopy.Enabled = True              'Selected text, allow copy
    
    If frmChat.tcpClient.State = sckConnected Then
        mnuEditChat.Enabled = True          'Copy to chat only if we're connected.
    Else
        mnuEditChat.Enabled = False
    End If

End Sub
