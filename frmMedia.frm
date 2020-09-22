VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmMedia 
   Caption         =   "MCI Callback Demo"
   ClientHeight    =   5865
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5850
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   5850
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkLoop 
      Caption         =   "&Loop"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      ToolTipText     =   "Auto Repeat "
      Top             =   1080
      Width           =   975
   End
   Begin VB.Timer PosTimer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5280
      Top             =   960
   End
   Begin MSComctlLib.Slider PosSlider 
      Height          =   615
      Left            =   1200
      TabIndex        =   12
      Top             =   3360
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   1085
      _Version        =   393216
      TickStyle       =   2
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Cl&ear List"
      Height          =   255
      Left            =   4680
      TabIndex        =   11
      Top             =   4320
      Width           =   975
   End
   Begin VB.ListBox lstMsg 
      Height          =   1230
      Left            =   0
      TabIndex        =   10
      Top             =   4320
      Width           =   5895
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "&Browse"
      Height          =   375
      Left            =   4920
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   5280
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   4080
      TabIndex        =   6
      Top             =   480
      Width           =   735
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "&Stop"
      Height          =   375
      Left            =   3240
      TabIndex        =   5
      Top             =   480
      Width           =   735
   End
   Begin VB.CommandButton cmdPause 
      Caption         =   "P&ause"
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   480
      Width           =   735
   End
   Begin VB.Frame MediaFrame 
      Height          =   2415
      Left            =   1178
      TabIndex        =   8
      Top             =   960
      Width           =   3495
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "&Play"
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   480
      Width           =   855
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "&Open"
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   480
      Width           =   855
   End
   Begin VB.TextBox txtFile 
      Height          =   285
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © 2002 Behrooz Sangani"
      ForeColor       =   &H80000011&
      Height          =   255
      Left            =   120
      TabIndex        =   14
      ToolTipText     =   "Email: bs20014@yahoo.com"
      Top             =   5640
      Width           =   2775
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000011&
      X1              =   2760
      X2              =   5880
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      X1              =   2760
      X2              =   5880
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Label Label1 
      Caption         =   "File:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblStats 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   3960
      Width           =   5655
   End
End
Attribute VB_Name = "frmMedia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================================================================
'  Main Form
'  Media Demo Form
'=========================================================================================
'  Created By: Behrooz Sangani <bs20014@yahoo.com>
'  Published Date: 26/09/2002
'  WebSite: http://www.geocities.com/bs20014
'  Legal Copyright: Behrooz Sangani © 26/09/2002
'=========================================================================================
'Feel free to use in your applications
'and give credit if you like
'If you want to redistribute this project on any
'kind of media or on the net you must distribute it
'as is and unmodified and you must not earn money from
'this sample.

'Comments are appreciated. Please send feedback to the email above

'This application demonstrates MCI callback messages
'I wrote this example as a response to a question on
'PSC (www.planet-source-code.com) forum because I was
'unable to find any example of this feature included
'in MCI


'IMPORTANT: Do not use stop button to unload the form

Dim IsPaused As Boolean

Private Sub chkLoop_Click()
    bLoop = CBool(chkLoop.Value)    'Loop boolean
End Sub

'Open dialog. MCI can do wonders in opening media files
'Check the filter
Private Sub cmdBrowse_Click()
    On Error GoTo error
    With CD1
        .DialogTitle = "Browse for media files..."
        .CancelError = True
        .flags = &H1000 Or &H80000 Or &H4
        'Do you like this filter? :)
        .Filter = "All Media files |*.mpg;*.mpeg;*.dat;*.avi;*.mpa;*.mpv;*.m1v;*.mp2;*.mp3" & _
            ";*.mpe;*.mpm;*.snd;*.enc;*.au;*.aif;*.aiff;*.aifc;*.rmi;*.wav;*.wmv;*.wma;*.midi;*.mid" & _
            ";*.qt;*.mov"
        .ShowOpen
        If .FileName <> "" Then txtFile.Text = .FileName
    End With
error:      'Canceled
End Sub

Private Sub cmdClear_Click()
    'convenience
    lstMsg.Clear
End Sub

Private Sub cmdClose_Click()
    MediaClose MyAlias
    PosTimer.Enabled = False
End Sub

Private Sub cmdOpen_Click()
    If txtFile.Text = "" Then Exit Sub  'No file
    MediaOpen txtFile.Text, MyAlias
    DoEvents    'To avoid mixing got messages
    MediaPut MyAlias, 0, 0
    PosSlider.Max = MediaTotalFrames(MyAlias)
End Sub

Private Sub cmdPause_Click()
    If IsPaused Then
        MediaPlay MyAlias, MediaCurrentPosition(MyAlias)
        IsPaused = Not IsPaused
    Else
        MediaPause MyAlias
        IsPaused = Not IsPaused
    End If
End Sub

Private Sub cmdPlay_Click()
    If IsPaused Then
        'MediaResume MyAlias    '**
        'We want the play end notify on loop playing so we do not use resume
        'to supersede our play notification
        MediaPlay MyAlias, MediaCurrentPosition(MyAlias)
        IsPaused = Not IsPaused
    Else
        MediaPlay MyAlias
        PosTimer.Enabled = True
    End If
End Sub

Private Sub cmdStop_Click()
    MediaStop MyAlias
End Sub

Private Sub Form_Load()
    'Hook
    ghWnd = MediaFrame.hWnd
    OldWinProc = SetWindowLong(ghWnd, GWL_WNDPROC, AddressOf NewWindowProc)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MediaClose MyAlias
    'Unhook
    RemoveOldProc ghWnd
End Sub

Private Sub PosSlider_Scroll()
    MediaCurrentPosition(MyAlias) = PosSlider.Value
End Sub

'Timer for the current position
'It has nothing to do with looping
Private Sub PosTimer_Timer()
    PosSlider.Value = MediaCurrentPosition(MyAlias)
End Sub
