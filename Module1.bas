Attribute VB_Name = "modGeneral"
''''''''''''''''''''''''''''''''''
'' RAK-Player v1.1    ''
''                              ''
'' cODE   : TARS                ''
'' gFX    : TARS                ''
'' tESTING: DiZZie & TARS       ''
''                              ''
'' WWW: http://www.tars.dk      ''
''''''''''''''''''''''''''''''''''
Global Const HWND_TOPMOST = -1
Global Const SWP_NOACTIVATE = &H10
Global Const SWP_SHOWWINDOW = &H40

'Option Explicit
Public Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long
Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
'Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long
'Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

Public fastForwardSpeed As Long    ' seconds to seek for ff/rew
