VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   480
   ClientLeft      =   165
   ClientTop       =   -120
   ClientWidth     =   2280
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   480
   ScaleWidth      =   2280
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Q"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   9
      ToolTipText     =   "Quit"
      Top             =   240
      Width           =   255
   End
   Begin VB.CommandButton cmdPref 
      BackColor       =   &H00C0C0C0&
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   8
      ToolTipText     =   "Options"
      Top             =   240
      Width           =   135
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2040
      Top             =   480
   End
   Begin VB.CommandButton eject 
      BackColor       =   &H00C0C0C0&
      Caption         =   "^"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   5
      ToolTipText     =   "Eject CD"
      Top             =   240
      Width           =   255
   End
   Begin VB.CommandButton ff 
      BackColor       =   &H00C0C0C0&
      Caption         =   ">>"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1680
      TabIndex        =   7
      ToolTipText     =   "Forward"
      Top             =   240
      Width           =   255
   End
   Begin VB.CommandButton rew 
      BackColor       =   &H00C0C0C0&
      Caption         =   "<<"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   6
      ToolTipText     =   "Rewind"
      Top             =   240
      Width           =   255
   End
   Begin VB.CommandButton ftrack 
      BackColor       =   &H00C0C0C0&
      Caption         =   ">>|"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   4
      ToolTipText     =   "Next Track"
      Top             =   240
      Width           =   255
   End
   Begin VB.CommandButton btrack 
      BackColor       =   &H00C0C0C0&
      Caption         =   "|<<"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "Prev. Track"
      Top             =   240
      Width           =   255
   End
   Begin VB.CommandButton stopbtn 
      BackColor       =   &H00C0C0C0&
      Caption         =   "X"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   3
      ToolTipText     =   "Stop"
      Top             =   240
      Width           =   255
   End
   Begin VB.CommandButton pause 
      BackColor       =   &H00C0C0C0&
      Caption         =   "||"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   2
      ToolTipText     =   "Pause"
      Top             =   240
      Width           =   255
   End
   Begin VB.CommandButton play 
      BackColor       =   &H00C0C0C0&
      Caption         =   ">"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   1
      ToolTipText     =   "Play"
      Top             =   240
      Width           =   255
   End
   Begin VB.TextBox timeWindow 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   0
      TabIndex        =   10
      TabStop         =   0   'False
      ToolTipText     =   "Current track time"
      Top             =   0
      Width           =   975
   End
   Begin VB.Label tracktime 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   120
      Left            =   960
      TabIndex        =   12
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label totalplay 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   120
      Left            =   960
      TabIndex        =   11
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''
'' RAK-Player v1.03    ''
''                              ''
'' cODE   : TARS                ''
'' gFX    : TARS                ''
'' tESTING: DiZZie & TARS       ''
''                              ''
''''''''''''''''''''''''''''''''''

' Program Hisotry:
'
' v1.03
'
'fixed:
' - All bugs when quiting
' - Bug that made the timer start at 00:00:02
' - Bug that generated a MCI error at startup
'
'
' v1.02
'
'new:
' - Made the main windows always on-top
'
'
' v1.01
'
'fixed:
' - Program didn's remove itself from memory
'new:
' - Option to dis- enable single track cd's
'

Option Explicit

Private CDpause         As Boolean    ' is the cd paused?
Private fPlaying        As Boolean    ' true if CD is currently playing
Private fCDLoaded       As Boolean    ' true if CD is the the player
Private numTracks       As Integer    ' number of tracks on audio CD
Private trackLength()   As String     ' array containing length of each track
Private Track           As Integer    ' current track
Private Min             As Integer    ' current minute on track
Private Sec             As Integer    ' current second on track
Private cmd             As String     ' string to hold mci command strings

' Send a MCI command string
' If fShowError is true, display a message box on error
Private Function SendMCIString(cmd As String, fShowError As Boolean) As Boolean
Static rc As Long, errStr As String * 200

rc = mciSendString(cmd, 0, 0, hWnd)
If (fShowError And rc <> 0) Then mciGetErrorString rc, errStr, Len(errStr): MsgBox errStr
SendMCIString = (rc = 0)
End Function

Private Sub cmdPref_Click()
    Dialog.Show , Me
End Sub


Private Sub cmdQuit_Click()
'Unloads program
    On Error Resume Next

    'SendMCIString "stop cd wait", True
    'cmd = "seek cd to " & Track
    'SendMCIString cmd, True
    'fPlaying = False
    'Update
    'mciSendString "close all", 0, 0, 0
    If fCDLoaded = False Then End   ' Let's quit right and tight..
    stopbtn_Click
    SendMCIString "close all", False
    End

End Sub

Private Sub Form_Load()

  SendMCIString "close all", False

  ' If we're already running, or the cd is used then quit '
  If App.PrevInstance = True Or SendMCIString("open cdaudio alias cd wait shareable", True) = False Then End

  ' Initialize variables  '
  ' Timer1.Enabled = True '
  fastForwardSpeed = 5
  Dialog.Text1.Text = fastForwardSpeed
  fCDLoaded = False
  CDpause = False

  SendMCIString "set cd time format tmsf wait", False
  Timer1.Enabled = False
  Update

  ' ScaleMode = 3 '
  With Me
    .Top = 0
    .Left = Screen.Width - 4000
    SetWindowPos Me.hWnd, HWND_TOPMOST, .Left / 15, .Top / 15, .Width / 15, .Height / 15, SWP_NOACTIVATE Or SWP_SHOWWINDOW
  End With
  Min = 0
  Sec = 0

  Load Dialog
  Dialog.Hide
  App.TaskVisible = False
End Sub

Private Sub Form_Terminate()
  cmdQuit_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Close all MCI devices opened by this program
  cmdQuit_Click
End Sub

Private Sub play_Click()
' Play the CD
    'If fPlaying = False Then ' And CDpause = False Then
    '    Min = 0
    '    Sec = Sec - 2
    'End If
    'If CDpause = True Then
    '    If Sec <> 0 Then
    '        cmd = "play cd from " & Track & Sec
    '    Else
    '        cmd = "play cd from " & Track
    '    End If
    '    mciSendString cmd, 0, 0, 0
    'Else
    SendMCIString "play cd", True
    'End If
    Update
    Timer1.Enabled = True
    fPlaying = True
End Sub
Private Sub stopbtn_Click()
' Stop the CD play
    SendMCIString "stop cd wait", False
    Min = 0
    Sec = 0
    cmd = "seek cd to " & Track
    SendMCIString cmd, False
    fPlaying = False
    Timer1.Enabled = False
    Min = 0
    Sec = 0
    btrack_Click
    'Update
End Sub
Private Sub pause_Click()
' Pause the CD
    SendMCIString "pause cd", True
    Timer1.Enabled = False
    fPlaying = False
    Update
    Select Case CDpause
      Case False
        CDpause = True
      Case Else
        CDpause = False
        play_Click
    End Select
End Sub
Private Sub eject_Click()
' Eject the CD
    SendMCIString "set cd door open", True
    Update
End Sub
Private Sub ff_Click()
' Fast forward
Dim s As String * 40

  SendMCIString "set cd time format milliseconds", True
  mciSendString "status cd position wait", s, Len(s), 0
  If fPlaying Then cmd = "play cd from " & CStr(CLng(s) + fastForwardSpeed * 1000) Else cmd = "seek cd to " & CStr(CLng(s) + fastForwardSpeed * 1000)
  mciSendString cmd, 0, 0, 0
  SendMCIString "set cd time format tmsf", True
  Update
End Sub
Private Sub rew_Click()
' Rewind the CD
Dim s As String * 40

  SendMCIString "set cd time format milliseconds", True
  mciSendString "status cd position wait", s, Len(s), 0
  If fPlaying Then cmd = "play cd from " & CStr(CLng(s) - fastForwardSpeed * 1000) Else cmd = "seek cd to " & CStr(CLng(s) - fastForwardSpeed * 1000)
  mciSendString cmd, 0, 0, 0
  SendMCIString "set cd time format tmsf", True
  Update
End Sub
Private Sub ftrack_Click()
' Forward track
  Select Case Track
    Case Is < numTracks
      If fPlaying Then cmd = "play cd from " & Track + 1 Else cmd = "seek cd to " & Track + 1
      SendMCIString cmd, True
    Case Else
      SendMCIString "seek cd to 1", True
  End Select
  Update
End Sub
Private Sub btrack_Click()
' Go to previous track
Dim from As String

  If fPlaying Then cmd = "play cd from " & from Else cmd = "seek cd to " & from
  SendMCIString cmd, False
  Select Case Min And Sec
    Case 0
      If Track > 1 Then from = CStr(Track - 1) Else from = CStr(numTracks)
    Case Else
      from = CStr(Track)
  End Select
  Update
End Sub
Private Sub Update()

'On Error Resume Next
' Update the display and state variables
Static s As String * 30

' Check if CD is in the player
  mciSendString "status cd media present", s, Len(s), 0
  Select Case CBool(s)
    Case True
      ' Enable all the controls, get CD information '
      Select Case fCDLoaded
        Case False
          mciSendString "status cd number of tracks wait", s, Len(s), 0
          numTracks = CInt(Mid$(s, 1, 2))
          eject.Enabled = True
          ' If CD only has 1 track, then it's probably a data CD '
          If Dialog.opt1(0).Value = False And numTracks = 1 Then Exit Sub
          mciSendString "status cd length wait", s, Len(s), 0
          totalplay = "T: " & numTracks & " Ttime: " & s
          ReDim trackLength(1 To numTracks)
          Dim i As Long
          For i = 1 To numTracks
            cmd = "status cd length track " & i
            mciSendString cmd, s, Len(s), 0
            trackLength(i) = s
          Next
          play.Enabled = True
          pause.Enabled = True
          ff.Enabled = True
          rew.Enabled = True
          ftrack.Enabled = True
          btrack.Enabled = True
          stopbtn.Enabled = True
          fCDLoaded = True
          SendMCIString "seek cd to 1", True
      End Select
      ' Update the track time display '
      mciSendString "status cd position", s, Len(s), 0
      Track = CInt(Mid$(s, 1, 2))
      Min = CInt(Mid$(s, 4, 2))
      Sec = CInt(Mid$(s, 7, 2))
      timeWindow.Text = "[" & Format(Track, "00") & "] " & Format(Min, "00") & ":" & Format(Sec, "00")
      ' If Track Not 0 is selected, then show Track-time... '
      If Track > 0 Then tracktime = "Track: " & trackLength(Track)
      ' Check if CD is playing '
      mciSendString "status cd mode", s, Len(s), 0
      fPlaying = (Mid$(s, 1, 7) = "playing")
    Case Else
      eject.Enabled = False
      ' Disable all the controls, clear the display '
      If fCDLoaded Then
        play.Enabled = False
        pause.Enabled = False
        ff.Enabled = False
        rew.Enabled = False
        ftrack.Enabled = False
        btrack.Enabled = False
        stopbtn.Enabled = False
        fCDLoaded = False
        fPlaying = False
        totalplay = vbNullString
        tracktime = vbNullString
        timeWindow.Text = vbNullString
      End If
  End Select
End Sub

Private Sub Timer1_Timer()
  Update
End Sub
