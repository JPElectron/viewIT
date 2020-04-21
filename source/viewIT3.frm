VERSION 5.00
Begin VB.Form Password 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Password"
   ClientHeight    =   1215
   ClientLeft      =   8250
   ClientTop       =   7230
   ClientWidth     =   3150
   Icon            =   "viewIT3.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1215
   ScaleWidth      =   3150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   120
      Top             =   600
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   600
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   600
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   240
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   240
      Width           =   2655
   End
End
Attribute VB_Name = "Password"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
If Text1.Text = pass Then
    Dim frmWork As Form
    For Each frmWork In Forms
        Unload frmWork
        Set frmWork = Nothing
    Next frmWork
Else
    Unload Password
End If
End Sub

Private Sub Command2_Click()
resp = ""
waiting = False
Unload Password
End Sub

Private Sub Form_Load()
Text1.Enabled = False
waiting = True

Dim lR As Long

     wFlags% = &H2 Or &H1 Or &H40 Or &H10
     Res% = SetWindowPos(Password.hWnd, -1, 0, 0, 0, 0, wFlags%)
     Me.Show
Text1.Text = ""

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
Command1_Click
End If
End Sub

Private Sub Timer1_Timer()
Timer1.Enabled = False
Text1.Enabled = True
Text1.SetFocus
If pass = "" Then
    Command1_Click
End If
End Sub
