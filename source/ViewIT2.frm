VERSION 5.00
Begin VB.Form Setup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "viewIT Setup"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9240
   Icon            =   "ViewIT2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   9240
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   375
      Left            =   7680
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   4575
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   600
      Width           =   9015
   End
End
Attribute VB_Name = "Setup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Open App.Path & "\viewit.ini" For Output As #1
Print #1, Text1.Text

Close
End Sub

Private Sub Form_Load()
On Error Resume Next

Open App.Path & "\viewit.ini" For Input As #1

Do While Not EOF(1)
    Line Input #1, fdata 'get line from file
    Text1 = Text1 & fdata & vbCrLf
Loop

Close #1
End Sub

