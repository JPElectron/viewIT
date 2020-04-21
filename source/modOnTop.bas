Attribute VB_Name = "modOnTop"
'----------------------------------------------------------------------------------------------------------
' Use KbdHook = SetWindowsHookEx(WH_KEYBOARD_LL, AddressOf LowLevelKeyboardProc, App.hInstance, 0&) to hook keyboard
' Use Call UnhookWindowsHookEx(KbdHook) to unhook
' Failing to unhook when program or form ends will cause VB to crash
'-----------------------------------------------------------------------------------------------------------
'Option Explicit

Public Const VK_CAPITAL = 20
Public Const VK_NUMLOCK = 140
Public Const VK_SHIFT = 16
Public Const VK_CONTROL = 17
Public Const VK_ALT = 18
Public Const WH_KEYBOARD_LL = 13&
Public Const HC_ACTION = 0&

Public Type KBHookStruct
  vkCode As Long
  scanCode As Long
  Flags As Long
  time As Long
  dwExtraInfo As Long
End Type

Public Declare Function GetKeyState Lib "user32.dll" _
    (ByVal nVirtKey As Long) As Integer

Public Declare Function SetWindowsHookEx Lib "user32" _
   Alias "SetWindowsHookExA" _
  (ByVal idHook As Long, _
   ByVal lpfn As Long, _
   ByVal hmod As Long, _
   ByVal dwThreadId As Long) As Long
   
Public Declare Function UnhookWindowsHookEx Lib "user32" _
  (ByVal hHook As Long) As Long

Public Declare Function CallNextHookEx Lib "user32" _
  (ByVal hHook As Long, _
   ByVal nCode As Long, _
   ByVal wParam As Long, _
   ByVal lParam As Long) As Long
   
Public Declare Sub CopyMemory Lib "kernel32" _
   Alias "RtlMoveMemory" _
  (pDest As Any, _
   pSource As Any, _
   ByVal cb As Long)



Public KbdHook As Long

Public iKeyCount As Integer
Public bCopyOK As Boolean
Dim bCapsLock As Boolean

Type OsVersionInfo
    dwVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatform As Long
    szCSDVersion As String * 128
End Type

Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = 1
Public Const Flags = SWP_NOMOVE Or SWP_NOSIZE
Public Const HWND_TOPMOST = -1
Public Const HKEY_CURRENT_USER = &H80000001

'Registry Read permissions:
Private Const KEY_QUERY_VALUE = &H1&
Private Const KEY_ENUMERATE_SUB_KEYS = &H8&
Private Const KEY_NOTIFY = &H10&
Private Const READ_CONTROL = &H20000
Private Const STANDARD_RIGHTS_READ = READ_CONTROL
Private Const Key_Read = STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY

Private Const REG_DWORD = 4&       '32-bit number
Public Const SPI_SCREENSAVERRUNNING = 97&
Public tmplng&
Public nMouseMoves%
Public xPixel%
Public yPixel%
Public ScrMsg As String
Public ScreenWidth%
Public ScreenHeight%
Private OsVers As OsVersionInfo
Public winOS&
Public waiting

'--------------------------------------------------------------------------
'API declarations
'--------------------------------------------------------------------------
Private Declare Function FindWindow& Lib "user32" Alias "FindWindowA" (ByVal lpClassName$, ByVal lpWindowName$)
Private Declare Function GetVersionEx& Lib "kernel32" Alias "GetVersionExA" (lpStruct As OsVersionInfo)
Private Declare Function RegCloseKey& Lib "advapi32.dll" (ByVal HKey&)
Private Declare Function RegOpenKeyExA& Lib "advapi32.dll" (ByVal HKey&, ByVal lpszSubKey$, dwOptions&, ByVal samDesired&, lpHKey&)
Private Declare Function RegQueryValueExA& Lib "advapi32.dll" (ByVal HKey&, ByVal lpszValueName$, lpdwRes&, lpdwType&, ByVal lpDataBuff$, nSize&)
Public Declare Function SetWindowPos Lib "user32" (ByVal h&, ByVal hb&, ByVal x&, ByVal y&, ByVal cx&, ByVal cy&, ByVal f&) As Integer
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
'API functions
'Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
'Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)

'Mouse Event API constants
Public Const MOUSEEVENTF_LEFTDOWN = &H2
Public Const MOUSEEVENTF_LEFTUP = &H4
Public Const MOUSEEVENTF_MIDDLEDOWN = &H20
Public Const MOUSEEVENTF_MIDDLEUP = &H40
Public Const MOUSEEVENTF_RIGHTDOWN = &H8
Public Const MOUSEEVENTF_RIGHTUP = &H10


Public pass As String
Public resp As String
      
Sub Main()

    Select Case Left$(LCase$(Command()), 2)
        Case "/s"
            Res% = SystemParametersInfo(17, 0, ByVal 0&, 0)
            Res% = ShowCursor(False)
            KbdHook = SetWindowsHookEx(WH_KEYBOARD_LL, _
        AddressOf LowLevelKeyboardProc, _
        App.hInstance, _
        0&)
            If KbdHook <> 0 Then
        'do functions like showing the form etc
Load Window
        
        bCapsLock = False
    Else
        MsgBox "Keyhook failed to start - " & Err.LastDllError
        End
    End If
            
        Case "/c"
            Load Setup
            Setup.Show
            
        Case ""
        KbdHook = SetWindowsHookEx(WH_KEYBOARD_LL, _
        AddressOf LowLevelKeyboardProc, _
        App.hInstance, _
        0&)
            If KbdHook <> 0 Then
        'do functions like showing the form etc
Load Window
        
        bCapsLock = False
    Else
        MsgBox "Keyhook failed to start - " & Err.LastDllError
        End
    End If
            
        End Select
    'set and obtain the handle to the keyboard hook
    
    'check to see if hook started properly

End Sub

Sub EndIt()
'On Error Resume Next
Call UnhookWindowsHookEx(KbdHook)
    Res% = SystemParametersInfo(17, 1, ByVal 0&, 0)
    Res% = ShowCursor(True)
End Sub


      Public Function SetTopMostWindow(hWnd As Long, Topmost As Boolean) As Long

         If Topmost = True Then 'Make the window topmost
            SetTopMostWindow = SetWindowPos(hWnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)
         Else
            SetTopMostWindow = SetWindowPos(hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, Flags)
            SetTopMostWindow = False
         End If
      End Function

Public Function ShiftPressed() As Boolean
    ShiftPressed = (GetAsyncKeyState(VK_SHIFT) < 0)
End Function

Public Function CapsLock() As Boolean
    CapsLock = GetKeyState(VK_CAPITAL) = 1
End Function

Public Function NumLock() As Boolean
    NumLock = GetKeyState(VK_NUMLOCK) = 1
End Function

Public Function LowLevelKeyboardProc(ByVal nCode As Long, _
        ByVal wParam As Long, _
        ByVal lParam As Long) As Long
    Dim Shift As Integer
    Dim iCapsLock As Integer
    Static kbhook As KBHookStruct
    Dim sChr As String

    If nCode = HC_ACTION Then

        Call CopyMemory(kbhook, ByVal lParam, Len(kbhook))

        If iKeyCount = 0 Then
        
            'check for capslock state
            If CapsLock Then
                If bCapsLock = False Then
                    iCapsLock = 1
                Else
                    iCapsLock = 0
                End If
            End If
         
            If Not ShiftPressed Then '
                'MsgBox "Shift Pressed"
                Shift = 0
            Else
                Shift = 1
            End If
            sChr = ParseText(CInt(kbhook.vkCode), Shift, iCapsLock)
            Window.txtStroke = Window.txtStroke + sChr
            iKeyCount = 1
        Else
            iKeyCount = 0
        End If
    End If
    
    'determine if captured character is a space
    If sChr = " " Then
        bCopyOK = True
    Else
        bCopyOK = False
    End If
    
    'KeyCaptured
    
    LowLevelKeyboardProc = CallNextHookEx(KbdHook, _
        nCode, _
        wParam, _
        lParam)
End Function
Public Function CheckText(KeyCode As Integer) As Boolean 'checks for all non-extended keys
    Select Case KeyCode
        Case 48 To 57
            CheckText = True
        Case 187 To 220
            CheckText = True
        Case 65 - 90
            CheckText = True
        Case Else
            CheckText = False
    End Select
End Function

Public Function ParseText(KeyCode As Integer, Shift As Integer, CapsLock As Integer) As String

    If CapsLock = 1 Then Shift = 1
    Select Case KeyCode
        'Enter key
        Case 13
            ParseText = vbNewLine
        
        'A-Z
        Case 65 To 90
            If Shift = 0 Then
                ParseText = LCase(Chr(KeyCode))
            Else
                ParseText = UCase(Chr(KeyCode))
            End If
        
        '0-9
        Case 48
            If Shift = 0 Then
                ParseText = Chr(KeyCode)
            Else
                ParseText = ")"
            End If
        Case 49
            If Shift = 0 Then
                ParseText = Chr(KeyCode)
            Else
                ParseText = "!"
            End If
        Case 50
            If Shift = 0 Then
                ParseText = Chr(KeyCode)
            Else
                ParseText = "@"
            End If
        Case 51
            If Shift = 0 Then
                ParseText = Chr(KeyCode)
            Else
                ParseText = "#"
            End If
        Case 52
            If Shift = 0 Then
                ParseText = Chr(KeyCode)
            Else
                ParseText = "$"
            End If
        Case 53
            If Shift = 0 Then
                ParseText = Chr(KeyCode)
            Else
                ParseText = "%"
            End If
        Case 54
            If Shift = 0 Then
                ParseText = Chr(KeyCode)
            Else
                ParseText = "^"
            End If
        Case 55
            If Shift = 0 Then
                ParseText = Chr(KeyCode)
            Else
                ParseText = "&"
            End If
        Case 56
            If Shift = 0 Then
                ParseText = Chr(KeyCode)
            Else
                ParseText = "*"
            End If
        Case 57
            If Shift = 0 Then
                ParseText = Chr(KeyCode)
            Else
                ParseText = "("
            End If
        
        'NumberPad 0-9
        Case 96
            ParseText = "{KeyPad0}"
        Case 97
            ParseText = "{KeyPad1}"
        Case 98
            ParseText = "{KeyPad2}"
        Case 99
            ParseText = "{KeyPad3}"
        Case 100
            ParseText = "{KeyPad4}"
        Case 101
            ParseText = "{KeyPad5}"
        Case 102
            ParseText = "{KeyPad6}"
        Case 103
            ParseText = "{KeyPad7}"
        Case 104
            ParseText = "{KeyPad8}"
        Case 105
            ParseText = "{KeyPad9}"

        'NumberPad
        Case 144
            ParseText = "{NumLock}"
        Case 111
            ParseText = "/"
        Case 106
            ParseText = "*"
        Case 109
            ParseText = "-"
        Case 107
            ParseText = "+"
        Case 110
            ParseText = "."
        
        'Misc Keys
        Case 145
            ParseText = "{ScrollLock}"
        Case 19
            ParseText = "{Pause/Break}"
        Case 45
            ParseText = "{Insert}"
        Case 36
            ParseText = "{Home}"
        Case 33
            ParseText = "{PageUp}"
        Case 46
            ParseText = "{Delete}"
        Case 35
            ParseText = "{End}"
        Case 34
            ParseText = "{PageDown}"
        
        'Arrow Keys
        Case 38
            ParseText = "{UpArrow}"
        Case 37
            ParseText = "{LeftArrow}"
        Case 39
            ParseText = "{RightArrow}"
        Case 40
            ParseText = "{DownArrow}"
        
        'Function Keys
        Case 112
            ParseText = "{F1}"
        Case 113
            ParseText = "{F2}"
        Case 114
            ParseText = "{F3}"
        Case 115
            ParseText = "{F4}"
        Case 116
            ParseText = "{F5}"
        Case 117
            ParseText = "{F6}"
        Case 118
            ParseText = "{F7}"
        Case 119
            ParseText = "{F8}"
        Case 120
            ParseText = "{F9}"
        Case 121
            ParseText = "{F10}"
        Case 122
            ParseText = "{F11}"
        Case 123
            ParseText = "{F12}"
        
        Case 27
            ParseText = "{Esc}"
        Case 9
            ParseText = "{Tab}"
        Case 20
            ParseText = "{CapsLock}"
        Case 160
            'ParseText = "{LeftShift}"
        Case 161
            'ParseText = "{RightShift}"
        Case 17
            ParseText = "{Ctrl}"
        Case 91
            ParseText = "{LeftWinKey}"
        Case 92
            ParseText = "{RightWinKey}"
        Case 164
            ParseText = "{LeftAlt}"
        Case 165
            ParseText = "{RightAlt}"
        Case 32
            ParseText = " "
        Case 8
            ParseText = "{BS}"
        
        'Special Characters
        Case 189
            If Shift = 0 Then
                ParseText = "-"
            Else
                ParseText = "_"
            End If
        Case 187
            If Shift = 0 Then
                ParseText = "="
            Else
                ParseText = "+"
            End If
        Case 219
            If Shift = 0 Then
                ParseText = "["
            Else
                ParseText = "{"
            End If
        Case 221
            If Shift = 0 Then
                ParseText = "]"
            Else
                ParseText = "}"
            End If
        Case 186
            If Shift = 0 Then
                ParseText = ";"
            Else
                ParseText = ":"
            End If
        Case 222
            If Shift = 0 Then
                ParseText = "'"
            Else
                ParseText = """"
            End If
        Case 188
            If Shift = 0 Then
                ParseText = ","
            Else
                ParseText = "<"
            End If
        Case 190
            If Shift = 0 Then
                ParseText = "."
            Else
                ParseText = ">"
            End If
        Case 191
            If Shift = 0 Then
                ParseText = "/"
            Else
                ParseText = "?"
            End If
        Case 220
            If Shift = 0 Then
                ParseText = "\"
            Else
                ParseText = "|"
            End If
        Case 192
            If Shift = 0 Then
                ParseText = "`"
            Else
                ParseText = "~"
            End If
    End Select
End Function
