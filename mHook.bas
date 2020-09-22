Attribute VB_Name = "mHook"
'=========================================================================================
'  Hook module
'  Window Proc Functions
'=========================================================================================
'  Created By: Behrooz Sangani <bs20014@yahoo.com>
'  Published Date: 26/09/2002
'  WebSite: http://www.geocities.com/bs20014
'  Legal Copyright: Behrooz Sangani © 26/09/2002
'=========================================================================================



Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hWnd As Long, ByVal lpString As String) As Long

Public Const GWL_WNDPROC = (-4)

Public OldWinProc As Long

Public Function NewWindowProc(ByVal lhWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    'Check if we get the notify message
    If Msg = MM_MCINOTIFY Then
        'wParam is what we want
        Select Case wParam
            Case MCI_NOTIFY_SUCCESSFUL
                frmMedia.lstMsg.AddItem gCurFunction & "Notification successful"
                'The loop playback without timer,
                'just using the MCI notification
                If gCurFunction = "Play: " And bLoop Then
                    'Play alias from start
                    MediaPlay MyAlias
                End If
            Case MCI_NOTIFY_SUPERSEDED
                If IsPlaying Then
                    frmMedia.lstMsg.AddItem "Play: Notification superseded"
                Else
                    frmMedia.lstMsg.AddItem gCurFunction & "Notification superseded"
                End If
            Case MCI_NOTIFY_ABORTED
                frmMedia.lstMsg.AddItem gCurFunction & "Notification aborted"
            Case MCI_NOTIFY_FAILURE
                frmMedia.lstMsg.AddItem gCurFunction & "Notification failed"
        End Select
        frmMedia.lstMsg.Selected(frmMedia.lstMsg.ListCount - 1) = True
        Exit Function
    End If
    NewWindowProc = CallWindowProc(OldWinProc, lhWnd, Msg, wParam, lParam)
End Function

Public Sub RemoveOldProc(lhWnd As Long)
    Dim tmpProc As Long
    tmpProc = GetProp(lhWnd, "OldWinProc")
    If tmpProc = 0 Then
        Exit Sub
    End If
    RemoveProp lhWnd, "OldWinProc"
    SetWindowLong lhWnd, GWL_WNDPROC, tmpProc
End Sub
