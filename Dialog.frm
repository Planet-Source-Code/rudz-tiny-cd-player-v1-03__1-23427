VERSION 5.00
Begin VB.Form Dialog 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   270
   ClientLeft      =   2715
   ClientTop       =   3420
   ClientWidth     =   2535
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00FFFFFF&
   Icon            =   "Dialog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   270
   ScaleWidth      =   2535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton opt1 
      BackColor       =   &H00000000&
      Caption         =   "No"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   135
      Index           =   1
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   120
      Value           =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton opt1 
      BackColor       =   &H00000000&
      Caption         =   "Yes"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   135
      Index           =   0
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   3.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   140
      Left            =   1200
      TabIndex        =   3
      Text            =   "5"
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   2040
      TabIndex        =   1
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   1800
      TabIndex        =   0
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Allow Single Track"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   120
      Left            =   0
      TabIndex        =   6
      Top             =   120
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Set Fast Forward Speed"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   120
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   1155
   End
End
Attribute VB_Name = "Dialog"
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

Private Sub CancelButton_Click()
'LoadSettings
  Me.Hide
End Sub

Private Sub Form_Load()
  Me.Icon = Form1.Icon
  SetWindowPos Dialog.hWnd, HWND_TOPMOST, Dialog.Left / 15, Dialog.Top / 15, Dialog.Width / 15, Dialog.Height / 15, SWP_NOACTIVATE Or SWP_SHOWWINDOW
End Sub

Private Sub OKButton_Click()
' Set the fast-forward speed
    'Dim s As String
    's = InputBox("Enter the new speed in seconds", "Fast Forward Speed", CStr(fastForwardSpeed))
    's = txtFFS.Text
  If IsNumeric(Text1) Then fastForwardSpeed = CLng(Text1)
    'fastForwardSpeed = CLng(Text1.tex)
  Me.Hide
End Sub
