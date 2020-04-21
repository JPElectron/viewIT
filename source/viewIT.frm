VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form Window 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "JPElectron.com - ViewIt"
   ClientHeight    =   5730
   ClientLeft      =   6885
   ClientTop       =   5175
   ClientWidth     =   7110
   Icon            =   "viewIT.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5730
   ScaleWidth      =   7110
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Timer Mousemonitor 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   6240
      Top             =   2520
   End
   Begin VB.TextBox txtStroke 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   19800
      Top             =   14640
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      CausesValidation=   0   'False
      Height          =   15375
      Left            =   -3
      TabIndex        =   0
      Top             =   -3
      Width           =   20805
      ExtentX         =   36698
      ExtentY         =   27120
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   1
      RegisterAsDropTarget=   0
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "Window"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer 'URL count
Dim k As Integer 'currently displayed URL
Dim url() As String 'array of URLS
Dim Interval() 'array of intervals, in seconds
Dim ticktock As Integer 'number of second current page has been up
Dim KbdHook
Dim clickpause
Dim random
Dim oldk
Dim InINI As Boolean
Dim exturl() As String
Dim extinterval()

Private Declare Function SetCursorPos Lib "user32.dll" ( _
    ByVal x As Long, ByVal y As Long) As Long

Private Sub Form_Load()
If App.PrevInstance = True Then End

Window.WindowState = 2
random = False
i = 0 'initialize array count
Randomize
'load URLs and timing data from file
Open App.Path & "\viewit.ini" For Input As #1

ShowCursor (False)
Do While Not EOF(1)
    Line Input #1, fdata 'get line from file
    If Left(fdata, 1) <> ";" And fdata Like "*:*" = True Then
        ReDim Preserve url(0 To i) 'Extend the URL array size
        ReDim Preserve Interval(0 To i) 'Extend the interval array size
        
        url(i) = Left(fdata, InStr(fdata, ",") - 1)
        Interval(i) = Right(fdata, Len(fdata) - InStr(fdata, ",") - 1)
        
        i = i + 1
    End If
    
    If fdata Like "Password=*" Then
        pass = Left(Right(fdata, Len(fdata) - InStr(fdata, "=")), Len(fdata) - InStr(fdata, "="))
    End If
    
    If fdata Like "Cursor=*" Then
        ShowCursor (True)
    End If
    
    If fdata Like "ClickPause=*" Then
        Mousemonitor.Enabled = True
    End If
    
    If fdata Like "Randomize=*" Then
        random = True
    End If
        
Loop

Close #1

k = 0 'current url = #1
WebBrowser1.Navigate url(0) 'show first page
 
'start the display timer
Timer1.Enabled = True

Dim lR As Long
     wFlags% = &H2 Or &H1 Or &H40 Or &H10
     Res% = SetWindowPos(Window.hWnd, -1, 0, 0, 0, 0, wFlags%)
     Me.Show
     
Call SetCursorPos(Screen.Width, Screen.Height)

End Sub

Private Sub Form_Resize()
WebBrowser1.Top = -50
WebBrowser1.Left = -50
WebBrowser1.Height = Window.ScaleHeight + 75
WebBrowser1.Width = Window.ScaleWidth + 315
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call UnhookWindowsHookEx(KbdHook)
Call EndIt
ShowCursor (True)
End Sub

Private Sub Mousemonitor_Timer()
If CBool(GetAsyncKeyState(vbLeftButton)) = True _
Or CBool(GetAsyncKeyState(vbMiddleButton)) = True _
Or CBool(GetAsyncKeyState(vbRightButton)) = True Then
Timer1.Enabled = False
End If
End Sub

Private Sub Timer1_Timer()
'On Error Resume Next
Window.WindowState = 2
If Int(Interval(k)) <> 0 Then ticktock = ticktock + 1 'one more second has passed

If ticktock >= Int(Interval(k)) And Int(Interval(k)) <> 0 Then 'if the interval is up
    oldk = k
    k = k + 1 'rotate to next url
    If k = i Then k = 0 'if its the last url, start over
    If random = True Then
GUESS:
        k = Int((i + 1) * Rnd) - 1
        If k = oldk Then GoTo GUESS
    End If
    ticktock = 0 'reset timer
    WebBrowser1.Navigate url(k)
End If

End Sub

Private Sub txtStroke_Change()
If txtStroke.Text = "{RightArrow}" Then
If Interval(k) = 0 Then
    k = k + 1
    If k > UBound(Interval) Then k = 0
    ticktock = 0
    WebBrowser1.Navigate url(k)
Else
    ticktock = Int(Interval(k) - 1)
    Timer1_Timer
End If

End If

If txtStroke.Text = "{LeftArrow}" Then
k = k - 2
If k = -1 Then
k = i - 1
End If
If k = -2 Then
k = i - 2
End If
If UBound(Interval) = 0 Then k = 0
ticktock = Int(Interval(k) - 1)
Timer1_Timer
End If

If txtStroke.Text = "s" Or txtStroke = "p" Then
    Timer1.Enabled = Not Timer1.Enabled
End If

If txtStroke.Text = "{Esc}" Or txtStroke = "q" Then
        Load Password
        Password.Show
End If

txtStroke = ""

End Sub

Function FormCount(ByVal frmName As String) As Long
    Dim frm As Form
    For Each frm In Forms
        If StrComp(frm.Name, frmName, vbTextCompare) = 0 Then
            FormCount = FormCount + 1
        End If
    Next
End Function
