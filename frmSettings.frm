VERSION 5.00
Begin VB.Form frmSettings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Syndicom ID Settings"
   ClientHeight    =   2685
   ClientLeft      =   4755
   ClientTop       =   4365
   ClientWidth     =   3405
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   3405
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkOptOut 
      Caption         =   "I do not wish to use this feature."
      Height          =   195
      Left            =   360
      TabIndex        =   6
      Top             =   2280
      Width           =   2655
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   1800
      TabIndex        =   5
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Default         =   -1  'True
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox txtUserID 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label lblPassword 
      Caption         =   "Password :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label lblUserId 
      Caption         =   "User ID :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub chkOptOut_Click()

    If chkOptOut.Value = vbChecked Then
        txtUserID.Enabled = False
        txtPassword.Enabled = False
    Else
        txtUserID.Enabled = True
        txtPassword.Enabled = True
    End If

End Sub

Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub cmdSave_Click()

    
    If chkOptOut.Value = vbChecked Then
        gstrUserID = ""
        gstrPassword = ""
        gblnOptOut = True
    Else
        gstrUserID = txtUserID
        gstrPassword = txtPassword
        gblnOptOut = False
    End If

    Unload Me

End Sub



Private Sub Form_Activate()

    If chkOptOut.Value = vbChecked Then
        chkOptOut.SetFocus
    ElseIf txtUserID = "" Then
        txtUserID.SetFocus
    ElseIf txtPassword = "" Then
        txtPassword.SetFocus
    Else
        cmdSave.SetFocus
    End If

End Sub

Private Sub Form_Load()

'   Center over frmChat
    
    Me.Left = frmChat.Left + ((frmChat.Width - Me.Width) \ 2)
    Me.Top = frmChat.Top + ((frmChat.Height - Me.Height) \ 2)
    
    txtUserID = gstrUserID
    txtPassword = gstrPassword
    
    If gblnOptOut = True Then
        chkOptOut.Value = vbChecked
        txtUserID.Enabled = False
        txtPassword.Enabled = False
    End If
    
End Sub


Private Sub txtPassword_GotFocus()

    With txtPassword
        .SelStart = 0
        .SelLength = Len(txtPassword)
    End With

End Sub


Private Sub txtUserID_GotFocus()

    With txtUserID
        .SelStart = 0
        .SelLength = Len(txtUserID)
    End With

End Sub


