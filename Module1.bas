Attribute VB_Name = "Module1"

Global MyName As String
Global TheRoom As String
Global TheVictem As String
Global TheVicName As String
Global TheVicNamea As Integer
Global TheTour As String
Global TheHook As String
Global TheCount As Integer
Global TheImCount As Integer
Global MyPassy As String
Global Idiot As String
Global Mimic As Boolean
Global MimicR As Boolean
Global TheVicMime As String
Global HookTalker As String
Global TheMagic As Boolean
Global TheMagic2 As Boolean
Global RoomPart As String
Global SearchN As Boolean
Global PlayerJoin As Boolean
Global MyNameArray(0 To 1000) As String
Global Seeka As Boolean
Global bombSTRING As String
'llllll Java Flood
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Declare Function SndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Declare Function CreateSolidBrush Lib "GDI32" (ByVal crColor As Long) As Long
Declare Function Rectangle Lib "GDI32" (ByVal hdc As Long, ByVal x1 As Long, ByVal Y1 As Long, ByVal x2 As Long, ByVal Y2 As Long) As Long
Declare Function DeleteObject Lib "GDI32" (ByVal hObject As Long) As Long
Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal x1 As Long, ByVal Y1 As Long, ByVal x2 As Long, ByVal Y2 As Long) As Long
Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function RegisterWindowMessage& Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String)
Declare Function GetWindowLong& Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long)
Declare Function EnumWindows& Lib "user32" (ByVal lpenumfunc As Long, ByVal lparam As Long)
Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lparam As Long)
Declare Function GetClassName& Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long)
Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Declare Function getparent Lib "user32" Alias "GetParent" (ByVal hwnd As Long) As Long
Declare Function findwindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function FindWindowEx& Lib "user32" Alias "FindWindowExA" (ByVal hwndParent As Long, ByVal hWndChildAfter As Long, ByVal lpClassName As String, ByVal lpWindowName As String)
Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lparam As String) As Long
Declare Function sendmessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lparam As Long) As Long
Declare Function CreatePopupMenu Lib "user32" () As Long
Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetWindowtextlength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Declare Function GettopWindow Lib "user32" Alias "GetTopWindow" (ByVal hwnd As Long) As Long
Declare Function setfocusapi Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function getwindow Lib "user32" Alias "GetWindow" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function DestroyMenu Lib "user32" (ByVal hMenu%) As Integer
Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Declare Function InvertRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Declare Function SetWindowPos Lib "user32" (ByVal h%, ByVal hb%, ByVal x%, ByVal y%, ByVal cx%, ByVal cy%, ByVal f%) As Integer
Public Const base64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
   Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As Any) As Long
       '***Part of the bonus code********************************


   Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
       '*********************************************************
'Global Const MF_BITMAP = 4
Public Const MF_BITMAP = &H4

Public Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long



Global Const SWP_NOMOVE = 2
Global Const SWP_NOSIZE = 1
Global Const Flags = SWP_NOMOVE Or SWP_NOSIZE
Global Const HWND_TOPMOST = -1
Global Const HWND_NOTOPMOST = -2
Global Const KEY_SNAPSHOT = &H2C
Global chatsendbutton%
Global gesturebutton%
Global chattextbox%
Global User$



Public Const WM_CHAR = &H102
Public Const WM_SETTEXT = &HC
Public Const WM_USER = &H400
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_CLOSE = &H10
Public Const WM_COMMAND = &H111
Public Const WM_CLEAR = &H303
Public Const WM_DESTROY = &H2
Public Const wm_gettext = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205

Public Const BM_GETCHECK = &HF0
Public Const BM_GETSTATE = &HF2
Public Const BM_SETCHECK = &HF1
Public Const BM_SETSTATE = &HF3

Public Const LB_GETITEMDATA = &H199
Public Const LB_GETCOUNT = &H18B
Public Const LB_ADDSTRING = &H180
Public Const LB_DELETESTRING = &H182
Public Const LB_FINDSTRING = &H18F
Public Const LB_FINDSTRINGEXACT = &H1A2
Public Const LB_GETCURSEL = &H188
Public Const LB_GETTEXT = &H189
Public Const LB_GETTEXTLEN = &H18A
Public Const LB_SELECTSTRING = &H18C
Public Const LB_SETCOUNT = &H1A7
Public Const LB_SETCURSEL = &H186
Public Const LB_SETSEL = &H185
Public Const LB_INSERTSTRING = &H181

Public Const VK_HOME = &H24
Public Const VK_RIGHT = &H27
Public Const VK_CONTROL = &H11
Public Const VK_DELETE = &H2E
Public Const VK_DOWN = &H28
Public Const VK_LEFT = &H25
Public Const VK_RETURN = &HD
Public Const VK_SPACE = &H20
Public Const VK_TAB = &H9

Public Const GW_CHILD = 5
Public Const Gw_hwndFirst = 0
Public Const gw_hwndlast = 1
Public Const GW_HWNDNEXT = 2
Public Const GW_HWNDPREV = 3
Public Const GW_MAX = 5
Public Const GW_OWNER = 4
Public Const SW_MAXIMIZE = 3
Public Const SW_MINIMIZE = 6
Public Const SW_HIDE = 0
Public Const SW_RESTORE = 9
Public Const SW_SHOW = 5
Public Const SW_SHOWDEFAULT = 10
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMINNOACTIVE = 7
Public Const SW_SHOWNOACTIVATE = 4
Public Const SW_SHOWNORMAL = 1

Public Const MF_APPEND = &H100&
Public Const MF_DELETE = &H200&
Public Const MF_CHANGE = &H80&
Public Const MF_ENABLED = &H0&
Public Const MF_DISABLED = &H2&
Public Const MF_REMOVE = &H1000&
Public Const MF_POPUP = &H10&
Public Const MF_STRING = &H0&
Public Const MF_UNCHECKED = &H0&
Public Const MF_CHECKED = &H8&
Public Const MF_GRAYED = &H1&
Public Const MF_BYPOSITION = &H400&
Public Const MF_BYCOMMAND = &H0&
Public Const MOUSE_MOVED = &H1

Public Const GWW_HINSTANCE = (-6)
Public Const GWW_ID = (-12)
Public Const GWL_STYLE = (-16)


Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Type POINTAPI
   x As Long
   y As Long
End Type

Function StringToHex(TheString)
Dim TheHex, Final As String
If Len(TheString) <> 4 Then Exit Function
For i = 1 To Len(TheString)
TheHex = Hex(Asc(Mid(TheString, i, 1)))
If Len(TheHex) = 1 Then TheHex = "0" & TheHex
Final = Final & TheHex
Next i
StringToHex = Final
End Function

Function HexToString(TheHex)
Dim TheString, Final As String
Dim TheLast As Integer
If Len(TheHex) <> 8 Then Exit Function
TheLast = 1
For i = 1 To 4
TheString = Chr(CByte("&H" & Mid(TheHex, TheLast, 2)))
Final = Final & TheString
TheLast = TheLast + 2
Next i
HexToString = Final
End Function

Function IPToString(IP)
Dim TheSection, Final As String
IP = IP & "."
A = InStr(1, IP, ".")
For i = 1 To 4
TheSection = Mid(IP, 1, A - 1)
Final = Final & Chr(TheSection)
IP = Right(IP, Len(IP) - A)
A = InStr(1, IP, ".")
Next i
IPToString = Final
End Function

Function Login()
Login = Chr(0) & Chr(0) & Chr(4)
End Function

Function RealLen(TheNum)
Dim TheLen As String
p = Hex(TheNum)
Select Case Len(p)
Case 1
TheLen = Chr(0) & Chr(CByte("&H" & p))
Case 2
TheLen = Chr(0) & Chr(CByte("&H" & p))
Case 3
TheLen = Chr(CByte("&H" & Left(p, 1))) & Chr(CByte("&H" & Right(p, 2)))
Case 4
TheLen = Chr(CByte("&H" & Left(p, 2))) & Chr(CByte("&H" & Right(p, 2)))
End Select
RealLen = TheLen
End Function



'writtEn By: -I-MoUsE-I-
'mAn wAs this A sonnA BitCh
Function DEC(Strin)
     If Strin = "00" Then A = 0
     If Strin = "01" Then A = 1
     If Strin = "02" Then A = 2
     If Strin = "03" Then A = 3
     If Strin = "04" Then A = 4
     If Strin = "05" Then A = 5
     If Strin = "06" Then A = 6
     If Strin = "07" Then A = 7
     If Strin = "08" Then A = 8
     If Strin = "09" Then A = 9
     If Strin = "0A" Then A = 10
     If Strin = "0B" Then A = 11
     If Strin = "0C" Then A = 12
     If Strin = "0D" Then A = 13
     If Strin = "0E" Then A = 14
     If Strin = "0F" Then A = 15
     If Strin = "10" Then A = 16
     If Strin = "11" Then A = 17
     If Strin = "12" Then A = 18
     If Strin = "13" Then A = 19
     If Strin = "14" Then A = 20
     If Strin = "15" Then A = 21
     If Strin = "16" Then A = 22
     If Strin = "17" Then A = 23
     If Strin = "18" Then A = 24
     If Strin = "19" Then A = 25
     If Strin = "1A" Then A = 26
     If Strin = "1B" Then A = 27
     If Strin = "1C" Then A = 28
     If Strin = "1D" Then A = 29
     If Strin = "1E" Then A = 30
     If Strin = "1F" Then A = 31
     If Strin = "20" Then A = 32
     If Strin = "21" Then A = 33
     If Strin = "22" Then A = 34
     If Strin = "23" Then A = 35
     If Strin = "24" Then A = 36
     If Strin = "25" Then A = 37
     If Strin = "26" Then A = 38
     If Strin = "27" Then A = 39
     If Strin = "28" Then A = 40
     If Strin = "29" Then A = 41
     If Strin = "2A" Then A = 42
     If Strin = "2B" Then A = 43
     If Strin = "2C" Then A = 44
     If Strin = "2D" Then A = 45
     If Strin = "2E" Then A = 46
     If Strin = "2F" Then A = 47
     If Strin = "30" Then A = 48
     If Strin = "31" Then A = 49
     If Strin = "32" Then A = 50
     If Strin = "33" Then A = 51
     If Strin = "34" Then A = 52
     If Strin = "35" Then A = 53
     If Strin = "36" Then A = 54
     If Strin = "37" Then A = 55
     If Strin = "38" Then A = 56
     If Strin = "39" Then A = 57
     If Strin = "3A" Then A = 58
     If Strin = "3B" Then A = 59
     If Strin = "3C" Then A = 60
     If Strin = "3D" Then A = 61
     If Strin = "3E" Then A = 62
     If Strin = "3F" Then A = 63
     If Strin = "40" Then A = 64
     If Strin = "41" Then A = 65
     If Strin = "42" Then A = 66
     If Strin = "43" Then A = 67
     If Strin = "44" Then A = 68
     If Strin = "45" Then A = 69
     If Strin = "46" Then A = 70
     If Strin = "47" Then A = 71
     If Strin = "48" Then A = 72
     If Strin = "49" Then A = 73
     If Strin = "4A" Then A = 74
     If Strin = "4B" Then A = 75
     If Strin = "4C" Then A = 76
     If Strin = "4D" Then A = 77
     If Strin = "4E" Then A = 78
     If Strin = "4F" Then A = 79
     If Strin = "50" Then A = 80
     If Strin = "51" Then A = 81
     If Strin = "52" Then A = 82
     If Strin = "53" Then A = 83
     If Strin = "54" Then A = 84
     If Strin = "55" Then A = 85
     If Strin = "56" Then A = 86
     If Strin = "57" Then A = 87
     If Strin = "58" Then A = 88
     If Strin = "59" Then A = 89
     If Strin = "5A" Then A = 90
     If Strin = "5B" Then A = 91
     If Strin = "5C" Then A = 92
     If Strin = "5D" Then A = 93
     If Strin = "5E" Then A = 94
     If Strin = "5F" Then A = 95
     If Strin = "60" Then A = 96
     If Strin = "61" Then A = 97
     If Strin = "62" Then A = 98
     If Strin = "63" Then A = 99
     If Strin = "64" Then A = 100
     If Strin = "65" Then A = 101
     If Strin = "66" Then A = 102
     If Strin = "67" Then A = 103
     If Strin = "68" Then A = 104
     If Strin = "69" Then A = 105
     If Strin = "6A" Then A = 106
     If Strin = "6B" Then A = 107
     If Strin = "6C" Then A = 108
     If Strin = "6D" Then A = 109
     If Strin = "6E" Then A = 110
     If Strin = "6F" Then A = 111
     If Strin = "70" Then A = 112
     If Strin = "71" Then A = 113
     If Strin = "72" Then A = 114
     If Strin = "73" Then A = 115
     If Strin = "74" Then A = 116
     If Strin = "75" Then A = 117
     If Strin = "76" Then A = 118
     If Strin = "77" Then A = 119
     If Strin = "78" Then A = 120
     If Strin = "79" Then A = 121
     If Strin = "7A" Then A = 122
     If Strin = "7B" Then A = 123
     If Strin = "7C" Then A = 124
     If Strin = "7D" Then A = 125
     If Strin = "7E" Then A = 126
     If Strin = "7F" Then A = 127
     If Strin = "80" Then A = 128
     If Strin = "81" Then A = 129
     If Strin = "82" Then A = 130
     If Strin = "83" Then A = 131
     If Strin = "84" Then A = 132
     If Strin = "85" Then A = 133
     If Strin = "86" Then A = 134
     If Strin = "87" Then A = 135
     If Strin = "88" Then A = 136
     If Strin = "89" Then A = 137
     If Strin = "8A" Then A = 138
     If Strin = "8B" Then A = 139
     If Strin = "8C" Then A = 140
     If Strin = "8D" Then A = 141
     If Strin = "8E" Then A = 142
     If Strin = "8F" Then A = 143
     If Strin = "90" Then A = 144
     If Strin = "91" Then A = 145
     If Strin = "92" Then A = 146
     If Strin = "93" Then A = 147
     If Strin = "94" Then A = 148
     If Strin = "95" Then A = 149
     If Strin = "96" Then A = 150
     If Strin = "97" Then A = 151
     If Strin = "98" Then A = 152
     If Strin = "99" Then A = 153
     If Strin = "9A" Then A = 154
     If Strin = "9B" Then A = 155
     If Strin = "9C" Then A = 156
     If Strin = "9D" Then A = 157
     If Strin = "9E" Then A = 158
     If Strin = "9F" Then A = 159
     If Strin = "A0" Then A = 160
     If Strin = "A1" Then A = 161
     If Strin = "A2" Then A = 162
     If Strin = "A3" Then A = 163
     If Strin = "A4" Then A = 164
     If Strin = "A5" Then A = 165
     If Strin = "A6" Then A = 166
     If Strin = "A7" Then A = 167
     If Strin = "A8" Then A = 168
     If Strin = "A9" Then A = 169
     If Strin = "AA" Then A = 170
     If Strin = "AB" Then A = 171
     If Strin = "AC" Then A = 172
     If Strin = "AD" Then A = 173
     If Strin = "AE" Then A = 174
     If Strin = "AF" Then A = 175
     If Strin = "B0" Then A = 176
     If Strin = "B1" Then A = 177
     If Strin = "B2" Then A = 178
     If Strin = "B3" Then A = 179
     If Strin = "B4" Then A = 180
     If Strin = "B5" Then A = 181
     If Strin = "B6" Then A = 182
     If Strin = "B7" Then A = 183
     If Strin = "B8" Then A = 184
     If Strin = "B9" Then A = 185
     If Strin = "BA" Then A = 186
     If Strin = "BB" Then A = 187
     If Strin = "BC" Then A = 188
     If Strin = "BD" Then A = 189
     If Strin = "BE" Then A = 190
     If Strin = "BF" Then A = 191
     If Strin = "C0" Then A = 192
     If Strin = "C1" Then A = 193
     If Strin = "C2" Then A = 194
     If Strin = "C3" Then A = 195
     If Strin = "C4" Then A = 196
     If Strin = "C5" Then A = 197
     If Strin = "C6" Then A = 198
     If Strin = "C7" Then A = 199
     If Strin = "C8" Then A = 200
     If Strin = "C9" Then A = 201
     If Strin = "CA" Then A = 202
     If Strin = "CB" Then A = 203
     If Strin = "CC" Then A = 204
     If Strin = "CD" Then A = 205
     If Strin = "CE" Then A = 206
     If Strin = "CF" Then A = 207
     If Strin = "D0" Then A = 208
     If Strin = "D1" Then A = 209
     If Strin = "D2" Then A = 210
     If Strin = "D3" Then A = 211
     If Strin = "D4" Then A = 212
     If Strin = "D5" Then A = 213
     If Strin = "D6" Then A = 214
     If Strin = "D7" Then A = 215
     If Strin = "D8" Then A = 216
     If Strin = "D9" Then A = 217
     If Strin = "DA" Then A = 218
     If Strin = "DB" Then A = 219
     If Strin = "DC" Then A = 220
     If Strin = "DD" Then A = 221
     If Strin = "DE" Then A = 222
     If Strin = "DF" Then A = 223
     If Strin = "E0" Then A = 224
     If Strin = "E1" Then A = 225
     If Strin = "E2" Then A = 226
     If Strin = "E3" Then A = 227
     If Strin = "E4" Then A = 228
     If Strin = "E5" Then A = 229
     If Strin = "E6" Then A = 230
     If Strin = "E7" Then A = 231
     If Strin = "E8" Then A = 232
     If Strin = "E9" Then A = 233
     If Strin = "EA" Then A = 234
     If Strin = "EB" Then A = 235
     If Strin = "EC" Then A = 236
     If Strin = "ED" Then A = 237
     If Strin = "EE" Then A = 238
     If Strin = "EF" Then A = 239
     If Strin = "F0" Then A = 240
     If Strin = "F1" Then A = 241
     If Strin = "F2" Then A = 242
     If Strin = "F3" Then A = 243
     If Strin = "F4" Then A = 244
     If Strin = "F5" Then A = 245
     If Strin = "F6" Then A = 246
     If Strin = "F7" Then A = 247
     If Strin = "F8" Then A = 248
     If Strin = "F9" Then A = 249
     If Strin = "FA" Then A = 250
     If Strin = "FB" Then A = 251
     If Strin = "FC" Then A = 252
     If Strin = "FD" Then A = 253
     If Strin = "FE" Then A = 254
     If Strin = "FF" Then A = 255
     DEC = A
End Function
Public Function DBL_Mod(ByVal N1 As Double, ByVal N2 As Double) As Double
    DBL_Mod = CDbl(N1 - (DBL_Divide(N1, N2)) * N2)
End Function

Public Function DBL_Divide(ByVal N1 As Double, ByVal N2 As Double) As Double
    DBL_Divide = Int(N1 / N2)
End Function

Public Function DEC_HEX(ByVal Number As Double) As String
    Dim i As Long, j As String, s As String
    Do
        j = Trim(CStr(DBL_Mod(Val(CStr(Number)), 16)))
        
        If j > 9 Then
            j = Chr((Val(j)) + 55)
        End If
        
        Number = DBL_Divide(Number, 16)
        s = Trim(j) & s
    Loop Until Number = 0
    
    DEC_HEX = CStr(s)
    
End Function

Function AsciiToHex(Strin)
'this was written By: -I-MoUsE-I-!
    Dim NewSTrin As String
    
    Do Until Strin = ""
        x = Hex(AscB(Left(Strin, 1)))
        
        If Len(TrimSpaces(x)) = 2 Then
            NewSTrin = NewSTrin & x
        Else
            NewSTrin = NewSTrin & "0" & x
        End If
        
        Strin = Right(Strin, Len(Strin) - 1)
    Loop
    
    AsciiToHex = NewSTrin
    
End Function

Function AsciiToHex2(Strin As String)
'this was written By: -I-MoUsE-I-!
    Dim NewSTrin As String
    
    Do Until Strin = ""
        x = Hex(AscB(Left(Strin, 1)))
        
        If Len(TrimSpaces(x)) = 2 Then
            NewSTrin = NewSTrin & x & " "
        Else
            NewSTrin = NewSTrin & "0" & x & " "
        End If
        
        Strin = Right(Strin, Len(Strin) - 1)
    Loop
    
    AsciiToHex2 = NewSTrin
    
End Function

Function Hex_Dec(Hex_val As String) As Variant


    Dim lstr
    Dim x, y
    Dim ch_in
    Dim conv_temp As Variant
    lstr = Len(Hex_val)
    For x = 0 To lstr - 1
        y = lstr - x
        ch_in = Mid$(Hex_val, y, 1)
        If Asc(ch_in) >= 48 And Asc(ch_in) <= 57 Then
            ch_in = ch_in
        ElseIf Asc(ch_in) >= 65 And Asc(ch_in) <= 70 Then
            ch_in = Asc(ch_in) - 55
        ElseIf Asc(ch_in) >= 97 And Asc(ch_in) <= 102 Then
            ch_in = Asc(ch_in) - 87
        End If


        Hex_Dec = Hex_Dec + 16 ^ x * ch_in
    Next x


End Function

Function Hex_Dec2(Hex_val As String) As Variant


    Dim lstr
    Dim x, y
    Dim ch_in
    Dim conv_temp As Variant
    lstr = Len(Hex_val)
    For x = 0 To lstr - 1
        y = lstr - x
        ch_in = Mid$(Hex_val, y, 1)
        If Asc(ch_in) >= 48 And Asc(ch_in) <= 57 Then
            ch_in = ch_in
        ElseIf Asc(ch_in) >= 65 And Asc(ch_in) <= 70 Then
            ch_in = Asc(ch_in) - 55
        ElseIf Asc(ch_in) >= 97 And Asc(ch_in) <= 102 Then
            ch_in = Asc(ch_in) - 87
        End If


        Hex_Dec2 = Hex_Dec2 + 16 ^ x * ch_in
        Hex_Dec2 = Hex_Dec2 & " "
    Next x


End Function

Function TrimSpaces(text)


    If InStr(text, " ") = 0 Then
        TrimSpaces = text
        Exit Function
    End If

    For trimspace = 1 To Len(text)
        TheChar$ = Mid(text, trimspace, 1)
        thechars$ = thechars$ & TheChar$

        If TheChar$ = " " Then
            thechars$ = Mid(thechars$, 1, Len(thechars$) - 1)
        End If
    Next trimspace

    TrimSpaces = thechars$
End Function

Function GetCaption(hwnd)
hwndlength% = GetWindowtextlength(hwnd)
hwndtitle$ = String$(hwndlength%, 0)
A% = GetWindowText(hwnd, hwndtitle$, (hwndlength% + 1))

GetCaption = hwndtitle$
End Function

Function FindChildByTitle(parentw, childhand)
    firs% = getwindow(parentw, 5)
    If UCase(GetCaption(firs%)) Like UCase(childhand) Then
        GoTo bone
    End If
    firs% = getwindow(parentw, GW_CHILD)

    While firs%
        firss% = getwindow(parentw, 5)
        If UCase(GetCaption(firss%)) Like UCase(childhand) & "*" Then
            GoTo bone
        End If
        firs% = getwindow(firs%, 2)
        If UCase(GetCaption(firs%)) Like UCase(childhand) & "*" Then
            GoTo bone
        End If
        Wend
    FindChildByTitle = 0
bone:
    room% = firs%
    FindChildByTitle = room%
End Function

Function VPGetText(child)
'Get the text of a window
    gettrim = SendMessageByNum(child, 14, 0&, 0&)
    trimspace$ = Space$(gettrim)
    getstrin = SendMessageByString(child, 13, gettrim + 1, trimspace$)

    VPGetText = trimspace$
End Function

Sub StayOnTop(frm As Form)
    On Error GoTo don
    Dim success%
    success% = SetWindowPos(frm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)
don:
End Sub

Sub Pause(interval)
    Current = Timer
    Do While Timer - Current < Val(interval)
        DoEvents
    Loop
End Sub

Function FileExist(Fname As String) As Boolean
    On Local Error Resume Next
   FileExist = (Dir(Fname) <> "")
End Function

Function VPWindow()
    vp% = findwindow("VPFrame", vbNullString)
    VPWindow = vp%
End Function

Function base64_encode(DecryptedText As String) As String
Dim C1, C2, c3 As Integer
Dim w1 As Integer
Dim w2 As Integer
Dim w3 As Integer
Dim w4 As Integer
Dim n As Integer
Dim retry As String
   For n = 1 To Len(DecryptedText) Step 3
      C1 = Asc(Mid$(DecryptedText, n, 1))
      C2 = Asc(Mid$(DecryptedText, n + 1, 1) + Chr$(0))
      c3 = Asc(Mid$(DecryptedText, n + 2, 1) + Chr$(0))
      w1 = Int(C1 / 4)
      w2 = (C1 And 3) * 16 + Int(C2 / 16)
      If Len(DecryptedText) >= n + 1 Then w3 = (C2 And 15) * 4 + Int(c3 / 64) Else w3 = -1
      If Len(DecryptedText) >= n + 2 Then w4 = c3 And 63 Else w4 = -1
      retry = retry + mimeencode(w1) + mimeencode(w2) + mimeencode(w3) + mimeencode(w4)
   Next
   base64_encode = retry
End Function

Function base64_decode(A As String) As String
Dim w1 As Integer
Dim w2 As Integer
Dim w3 As Integer
Dim w4 As Integer
Dim n As Integer
Dim retry As String

   For n = 1 To Len(A) Step 4
      w1 = mimedecode(Mid$(A, n, 1))
      w2 = mimedecode(Mid$(A, n + 1, 1))
      w3 = mimedecode(Mid$(A, n + 2, 1))
      w4 = mimedecode(Mid$(A, n + 3, 1))
      If w2 >= 0 Then retry = retry + Chr$(((w1 * 4 + Int(w2 / 16)) And 255))
      If w3 >= 0 Then retry = retry + Chr$(((w2 * 16 + Int(w3 / 4)) And 255))
      If w4 >= 0 Then retry = retry + Chr$(((w3 * 64 + w4) And 255))
   Next
   base64_decode = retry
End Function

Private Function mimeencode(w As Integer) As String
   If w >= 0 Then mimeencode = Mid$(base64, w + 1, 1) Else mimeencode = ""
End Function

Private Function mimedecode(A As String) As Integer
   If Len(A) = 0 Then mimedecode = -1: Exit Function
   mimedecode = InStr(base64, A) - 1
End Function

Public Sub PlaySound(strFileName As String)
    SndPlaySound strFileName, 1
End Sub

Function FileExista(Fname As String) As Boolean
    On Local Error Resume Next
   FileExista = (Dir(Fname) <> "")
End Function

Function Wave_Lenght(Dateiname)
    Dim i As Long, RS As String, cb As Long
    RS = Space$(128)
    i = mciSendString("stop sound", RS, 128, cb)
    i = mciSendString("close sound", RS, 128, cb)
    i = mciSendString("open waveaudio!" & Dateiname & " Alias sound", RS, 128, cb)
    i = mciSendString("status sound length", RS, 128, cb)
    Wave_Lenght = RS
    i = mciSendString("stop sound", RS, 128, cb)
   
    i = mciSendString("close sound", RS, 128, cb)
End Function

Public Sub PlayMouseSound(MouseSoundPath As String)
    Dim i As Long, RS As String, cb As Long
    RS = Space$(128)
    i = mciSendString("open waveaudio!" & MouseSoundPath & " Alias sound", RS, 128, cb)
    i = mciSendString("play sound", RS, 128, cb)
End Sub

Function findchildbyclass(parentw, childhand)
firs% = getwindow(parentw, 5)
If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo bone
firs% = getwindow(parentw, GW_CHILD)
If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo bone

While firs%
firss% = getwindow(parentw, 5)
If UCase(Mid(GetClass(firss%), 1, Len(childhand))) Like UCase(childhand) Then GoTo bone
firs% = getwindow(firs%, 2)
If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo bone
Wend
findchildbyclass = 0

bone:
room% = firs%
findchildbyclass = room%

End Function

Function GetClass(child)
buffer$ = String$(250, 0)
getclas% = GetClassName(child, buffer$, 250)

GetClass = buffer$
End Function

Function VPGetUser()
'Get the user name of the person using VP
hwndz% = findwindow(vbNullString, "My Identity")
If hwndz% = 0 Then
If GetCaption(VPWindow) = vbNullString Then Exit Function
AppActivate "vplaces"
SendKeys "%AE", True
hwndz% = findwindow(vbNullString, "My Identity")
End If
id1% = FindChildByTitle(hwndz%, "Basic Info")
firs% = getwindow(id1%, GW_CHILD)
VPGetUser = VPGetText(firs%)
hwndz2% = FindChildByTitle(hwndz%, "Cancel")
VPButton (hwndz2%)
VPButton (hwndz2%)
End Function

Public Sub VPButton(but%)
'Click on the button
clickicon% = sendmessage(but%, WM_KEYDOWN, VK_SPACE, 0)
clickicon% = sendmessage(but%, WM_KEYUP, VK_SPACE, 0)
End Sub

Sub OpenURL(lol)
ShellExecute hwnd, "open", lol, vbNullString, vbNullString, SW_SHOWMAXIMIZED
End Sub

Public Function text_read(filename)
'This function reads a file and spits out the text in it.

Dim f
Dim textda
Dim cha

On Error Resume Next
i = 1
f = FreeFile
textda = ""
        Open filename For Binary As #f   ' Open file.
            textda = Input(LOF(f), #f) ' I HAVE CHANGED THIS FROM 1 TO LOF(f) BECAUSE OF BIG FILES (100 KB)
            DoEvents  'I HAVE ADDED THIS FOR BIG FILES
        Close #f
text_read = textda


End Function

Function HTTP2Comp(TheString, TheType)
HTTP2Comp = ""
For i = 1 To Len(TheString)
    HTTP2Comp = HTTP2Comp & Asciier(Mid(TheString, i, 1), TheType)
    DoEvents
Next i
End Function

Function Asciier(TheChar, TheType)
Dim AS1
If TheType = 0 Then
    AS1 = Asc(TheChar) + 7
    If AS1 > 255 Then
        AS1 = (AS1 - 255) - 1
        Asciier = Chr(AS1)
    Else
        Asciier = Chr(AS1)
    End If
Else
    AS1 = Asc(TheChar) - 7
    If AS1 < 0 Then
        AS1 = (256 + AS1)
        Asciier = Chr(AS1)
    Else
        Asciier = Chr(AS1)
    End If
End If
End Function

Public Function GRInteger(LowerBound, UpperBound) As Long
GRInteger = Int((UpperBound - LowerBound + 1) * Rnd + LowerBound)
End Function

Sub ListLoader(Path, Lst As ListBox)
On Error Resume Next
Open Path For Input As #1
While Not EOF(1)
Input #1, What
Lst.AddItem What
DoEvents
Wend
Close #1
End Sub

Function FourByteLen(TheString As String) As String
FourByteLen = Chr(0) & Chr(Int(Len(TheString) / 65536)) & Chr(Int((Len(TheString) - ((Int(Len(TheString) / 65536)) * 65536)) / 256)) & Chr(Len(TheString) - (Int(Len(TheString) / 65536) * 65536) - ((Int((Len(TheString) - ((Int(Len(TheString) / 65536)) * 65536)) / 256)) * 256)) & TheString
End Function

'Give this a string and it prefixes a two byte length
Function TwoByteLen(TheString As String) As String
TwoByteLen = Chr(Int((Len(TheString) / 256))) & Chr((Len(TheString) - (Int(Len(TheString) / 256)) * 256)) & TheString
End Function

'Give this a numeric value and it returns a four byte value
Function ByteLen(PacketLength As Single) As String
ByteLen = Chr(0) & Chr(Int(PacketLength / 65536)) & Chr(Int((PacketLength - ((Int(PacketLength / 65536)) * 65536)) / 256)) & Chr(PacketLength - (Int(PacketLength / 65536) * 65536) - ((Int((PacketLength - ((Int(PacketLength / 65536)) * 65536)) / 256)) * 256))
End Function

'Give this a three byte value (four byte but first byte is assumed a value of 0) and it returns a numeric value
Function GetLength(ByteLength As String) As Single
GetLength = (Asc(Mid(ByteLength, 1, 1)) * 65536) + (Asc(Mid(ByteLength, 2, 1)) * 256) + (Asc(Mid(ByteLength, 3, 1)) * 1)
End Function

'Give this an integer and it will return a two byte base(256) string
Function IntegerToBase256(Value As Integer) As String
'Int(Value / 256) & " " & Value - (Int(Value / 256)) * 256
IntegerToBase256 = Chr(Int(Value / 256)) & Chr(Value - (Int(Value / 256)) * 256)
End Function

Function ChrA(strPoop) As String
Dim C1, C2, c3, C4
C1 = Split(strPoop, " "): ChrA = ""
For i = 0 To UBound(C1)
    If Left(C1(i), 1) = "H" Then
        ChrA = ChrA & Chr("&" & C1(i))
    Else
        ChrA = ChrA & Chr(C1(i))
    End If
    DoEvents
Next i
End Function
