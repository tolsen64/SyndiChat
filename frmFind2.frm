VERSION 5.00
Begin VB.Form frmFind2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find"
   ClientHeight    =   1710
   ClientLeft      =   2925
   ClientTop       =   4920
   ClientWidth     =   6735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1710
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtFind 
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   240
      Width           =   3735
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find"
      Default         =   -1  'True
      Height          =   495
      Left            =   5280
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   5280
      TabIndex        =   3
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label lblFind 
      Caption         =   "Fi&nd what:"
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
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   975
   End
End
Attribute VB_Name = "frmFind2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub cmdFind_Click()

    Dim lngPosFound As Long, intRetVal As Integer
    
    lngPosFound = InStr(1, frmScrollback.rtfScrollback.Text, gstrFind, vbTextCompare)

    If lngPosFound = 0 Then
        intRetVal = MsgBox(gstrFind & " not found.", vbOKOnly + vbInformation, "Find")
        cmdFind.SetFocus
        txtFind.SetFocus
    Else
        frmScrollback.rtfScrollback.SelStart = lngPosFound - 1
        frmScrollback.rtfScrollback.SelLength = Len(gstrFind)
        Unload Me
    End If

End Sub

Private Sub Form_Load()

'   Center over frmChat
    
    Me.Left = frmScrollback.Left + ((frmScrollback.Width - Me.Width) \ 2)
    Me.Top = frmScrollback.Top + ((frmScrollback.Height - Me.Height) \ 2)

    If gstrFind = "" Then
        cmdFind.Enabled = False
    Else
        cmdFind.Enabled = True
    End If
    
    txtFind.Text = gstrFind

End Sub


Private Sub txtFind_Change()
    
    If Len(txtFind.Text) = 0 Then
        cmdFind.Enabled = False
        frmScrollback.mnuEditFindNext.Enabled = False
    Else
        cmdFind.Enabled = True
        frmScrollback.mnuEditFindNext.Enabled = True
    End If
    
    gstrFind = txtFind.Text

End Sub


Private Sub txtFind_GotFocus()
    
    txtFind.SelStart = 0
    txtFind.SelLength = Len(txtFind.Text)

End Sub


