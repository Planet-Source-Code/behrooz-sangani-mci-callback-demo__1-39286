Attribute VB_Name = "mMedia"
'=========================================================================================
'  MultiMedia Module
'  Basic functions to play a media file
'=========================================================================================
'  Created By: Behrooz Sangani <bs20014@yahoo.com>
'  Published Date: 26/09/2002
'  WebSite: http://www.geocities.com/bs20014
'  Legal Copyright: Behrooz Sangani © 26/09/2002
'=========================================================================================

'This module contains some basic functions to
'handle a media file. You can add more MCI functions
'I just used them to demonstrate the MCI callback part

Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long

Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

Public Const WS_CHILD = &H40000000

Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Const MM_MCINOTIFY = &H3B9  'MCI
'Flags for wParam of MM_MCINOTIFY message
Public Const MCI_NOTIFY_SUCCESSFUL = &H1        'Notification successful
Public Const MCI_NOTIFY_SUPERSEDED = &H2        'Notification superseded
Public Const MCI_NOTIFY_ABORTED = &H4           'Notification aborted
Public Const MCI_NOTIFY_FAILURE = &H8           'Notification failed

'Our instance alias.
'But setting different aliases
'you can open as many as media files as you want
Public Const MyAlias = "BSMovie1"

Public ghWnd As Long            'Our window to be hooked
Public gCurFunction As String   'Current processing function
Public IsPlaying As Boolean     'True during playback
Public bLoop As Boolean         'Loop playback

Dim lRet As Long                'Return Value
Dim strCommand As String        'MCI command string

'Open a media file. (A file must be opened before it can be played)
Public Sub MediaOpen(sFileName As String, InstanceAlias As String)
    'What we want notification for
    gCurFunction = "Open: "

    'Get the short path name of the file
    sFileName = ShortName(sFileName)

    'build the open command string for mci  ==NOTICE: The 'notify' flag at the end of the command
    strCommand = "Open " & sFileName & " type MPEGVideo Alias " & InstanceAlias & " parent " & ghWnd & " Style " & WS_CHILD & " notify"
    'Send string
    lRet = mciSendString(strCommand, 0&, 0&, ghWnd) 'set hwnd for the call back

    'Get the error message if there is any else return current status
    If lRet <> 0 Then
        frmMedia.lblStats.Caption = ErrText(lRet)
    Else
        frmMedia.lblStats.Caption = MediaStat(InstanceAlias)
    End If
End Sub

'Close media instance
Public Sub MediaClose(InstanceAlias As String)
    gCurFunction = "Close: "

    'Again the 'notify' flag
    strCommand = "Close " & InstanceAlias & " notify"
    lRet = mciSendString(strCommand, 0&, 0&, ghWnd)

    'Get the error message if there is any else return current status
    If lRet <> 0 Then
        frmMedia.lblStats.Caption = ErrText(lRet)
    Else
        frmMedia.lblStats.Caption = MediaStat(InstanceAlias)
        IsPlaying = False
    End If
End Sub

'Resume the paused media instance. Use this sub to resume
'after pause method not the pause method itself
Public Sub MediaResume(InstanceAlias As String)
    gCurFunction = "Resume: "

    strCommand = "Resume " & InstanceAlias & " notify"
    lRet = mciSendString(strCommand, 0&, 0&, ghWnd)

    'Get the error message if there is any else return current status
    If lRet <> 0 Then
        frmMedia.lblStats.Caption = ErrText(lRet)
    Else
        frmMedia.lblStats.Caption = MediaStat(InstanceAlias)
    End If
End Sub

'Pause the named media instance
Public Sub MediaPause(InstanceAlias As String)
    gCurFunction = "Pause: "

    strCommand = "Pause " & InstanceAlias & " notify"
    lRet = mciSendString(strCommand, 0&, 0&, ghWnd)

    'Get the error message if there is any else return current status
    If lRet <> 0 Then
        frmMedia.lblStats.Caption = ErrText(lRet)
    Else
        frmMedia.lblStats.Caption = MediaStat(InstanceAlias)
    End If
End Sub

'Play the named media instance. File must be opened
'before it can be played
Public Sub MediaPlay(InstanceAlias As String, Optional lFrom = 0, Optional lTo = "TheEnd")
    gCurFunction = "Play: "

    'If we should play to the end of the file
    If lTo = "TheEnd" Then lTo = MediaTotalFrames(InstanceAlias)

    'Play from lFrom to lTo
    strCommand = "Play " & InstanceAlias & " from " & CStr(lFrom) & " to " & CStr(lTo) & " notify"
    lRet = mciSendString(strCommand, 0&, 0&, ghWnd)

    'Get the error message if there is any else return current status
    If lRet <> 0 Then
        frmMedia.lblStats.Caption = ErrText(lRet)
    Else
        frmMedia.lblStats.Caption = MediaStat(InstanceAlias)
        IsPlaying = True
    End If
End Sub

'Stop named media instance that is playing
Public Sub MediaStop(InstanceAlias As String)
    gCurFunction = "Stop: "

    strCommand = "Stop " & InstanceAlias & " notify"
    lRet = mciSendString(strCommand, 0&, 0&, ghWnd)

    'Get the error message if there is any else return current status
    If lRet <> 0 Then
        frmMedia.lblStats.Caption = ErrText(lRet)
    Else
        frmMedia.lblStats.Caption = MediaStat(InstanceAlias)
        IsPlaying = False
    End If
End Sub

'Changes the named media instance window to lLeft, lTop, and the optional
'lHeight and lWidth (If not set parent window height and width will be used)
Public Sub MediaPut(InstanceAlias As String, lLeft As Long, lTop As Long, Optional lWidth As Long, Optional lHeight As Long)
    gCurFunction = "Put: "

    Dim ParentRect As RECT
    'Get rect of the parent window
    GetWindowRect ghWnd, ParentRect
    'If width and height is not set then set them
    'to parent window width and height
    If lWidth = 0 Then lWidth = ParentRect.Right - ParentRect.Left
    If lHeight = 0 Then lHeight = ParentRect.Bottom - ParentRect.Top

    strCommand = "Put " & InstanceAlias & " window at " & lLeft & " " & lTop & " " & lWidth & " " & lHeight & " notify"
    lRet = mciSendString(strCommand, 0&, 0&, ghWnd)

    'Get the error message if there is any else return current status
    If lRet <> 0 Then
        frmMedia.lblStats.Caption = ErrText(lRet)
    Else
        frmMedia.lblStats.Caption = MediaStat(InstanceAlias)
    End If
End Sub

'Retrieves total media frames
Public Property Get MediaTotalFrames(InstanceAlias As String) As Long

    Dim sMsg As String
    sMsg = Space(128)
    'Just in case, set format to frames
    strCommand = "Set " & InstanceAlias & " time format frames"
    lRet = mciSendString(strCommand, sMsg, Len(sMsg), 0&)
    'Get total frames
    strCommand = "Status " & InstanceAlias & " length"
    lRet = mciSendString(strCommand, sMsg, Len(sMsg), 0&)

    If lRet <> 0 Then
        MediaTotalFrames = -1
        frmMedia.lblStats.Caption = ErrText(lRet)
    Else
        MediaTotalFrames = Val(sMsg)
    End If
End Property

'Let property seeks media to the frame user requires
Public Property Let MediaCurrentPosition(InstanceAlias As String, lTo As Long)

    'Seek command to ...
    strCommand = "Seek " & InstanceAlias & " to " & lTo & " notify"
    lRet = mciSendString(strCommand, 0&, 0&, ghWnd)
    'Play again
    strCommand = "Play " & InstanceAlias & " notify"
    mciSendString strCommand, 0&, 0&, ghWnd

    If lRet <> 0 Then
        frmMedia.lblStats.Caption = ErrText(lRet)
    End If
End Property

'Get property retrieves the current media frame
Public Property Get MediaCurrentPosition(InstanceAlias As String) As Long
    Dim sMsg As String
    sMsg = Space(128)
    strCommand = "Status " & InstanceAlias & " position"
    lRet = mciSendString(strCommand, sMsg, Len(sMsg), 0&)

    If lRet <> 0 Then
        MediaCurrentPosition = -1   'Not playing
        frmMedia.lblStats.Caption = ErrText(lRet)
    Else
        MediaCurrentPosition = Val(sMsg)
    End If
End Property

'Retrieves current media instance status
Private Property Get MediaStat(InstanceAlias As String) As String

    Dim Stat As String
    Stat = Space(128)
    lRet = mciSendString("Status " & InstanceAlias & " mode", Stat, Len(Stat), 0&)
    If lRet <> 0 Then
        frmMedia.lblStats.Caption = ErrText(lRet)
    Else
        MediaStat = Left(Stat, InStr(Stat, vbNullChar) - 1)
    End If
End Property
'MCI error string from error code
Private Property Get ErrText(ErrCode As Long) As String

    Dim sErr As String
    sErr = Space(128)
    mciGetErrorString ErrCode, sErr, Len(sErr)
    ErrText = Left(sErr, InStr(sErr, vbNullChar) - 1)
End Property

'Short Path Name for a long path
Private Property Get ShortName(sPath As String) As String

    'Returns short path of the file name
    Dim sBuffer As String
    sBuffer = Space(1024)
    GetShortPathName sPath, sBuffer, Len(sBuffer)
    ShortName = Left(sBuffer, InStr(1, sBuffer, vbNullChar) - 1)
End Property

