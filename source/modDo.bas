Attribute VB_Name = "modDo"
Public lBegin As Long
Public lEnd As Long
Public bRepeat As Boolean
Sub DoPlayPause()
    If frmMain.Real.Source <> "" Then
        frmMain.Real.DoPlayPause
    End If
End Sub
Sub DoBegin()
    If frmMain.Real.GetPlayState = 3 Or frmMain.Real.GetPlayState = 5 Then
        lBegin = frmMain.Real.GetPosition
        CheckLoopProgress
    End If
End Sub
Sub DoEnd()
    If frmMain.Real.GetPlayState = 3 Or frmMain.Real.GetPlayState = 5 Then
        lEnd = frmMain.Real.GetPosition
        CheckLoopProgress
    End If
End Sub
Sub DoRepeat()
    If frmMain.Real.GetPlayState = 3 Or frmMain.Real.GetPlayState = 5 Then
        If bRepeat = False Then
            lEnd = frmMain.Real.GetPosition
            bRepeat = True
        Else
            lBegin = lEnd
            frmMain.Real.SetPosition (lBegin)
            lEnd = frmMain.Real.GetLength
            bRepeat = False
        End If
        CheckLoopProgress
    End If
End Sub
Sub DoReset()
    If frmMain.Real.Source <> "" Then
        lBegin = 0
        lEnd = frmMain.Real.GetLength
        bRepeat = False
        CheckLoopProgress
    End If
End Sub
Sub DoOpen()
    'Set OpenFileDialog
    Dim ofn As OPENFILENAME 'read modOpneDialog
    Dim rtn As String
    
    ofn.lStructSize = Len(ofn)
    ofn.hwndOwner = frmMain.hwnd
    ofn.hInstance = App.hInstance
    ofn.lpstrFilter = "All File"
    ofn.lpstrFile = Space(254)
    ofn.nMaxFile = 255
    ofn.lpstrFileTitle = Space(254)
    ofn.nMaxFileTitle = 255
    'ofn.lpstrInitialDir = App.Path
    ofn.lpstrTitle = "Loopman - Open File"
    ofn.flags = 6148
    '''''''''''''''''''''''''
    rtn = GetOpenFileName(ofn)
    If rtn >= 1 Then
        frmMain.Real.SetSource (ofn.lpstrFile)
        frmMain.Real.DoPlay
        lBegin = 0
    End If
End Sub
Sub CheckLoopProgress()
    If frmMain.Real.Source <> "" And lBegin < lEnd Then
        frmMain.LoopArea.Width = (lEnd - lBegin) / frmMain.Real.GetLength * frmMain.ScaleWidth
        frmMain.LoopArea.Left = lBegin / frmMain.Real.GetLength * frmMain.ScaleWidth
    End If
End Sub
Sub DoGoToBegin()
    If frmMain.Real.GetPlayState = 3 Or frmMain.Real.GetPlayState = 5 Then
        frmMain.Real.SetPosition (lBegin)
    End If
End Sub
Sub DoBackward()
    If frmMain.Real.GetPlayState = 3 Or frmMain.Real.GetPlayState = 5 Then
        Dim Backward As Long
        Backward = frmMain.Real.GetPosition - 1000
        If Backward >= lBegin Then
            frmMain.Real.SetPosition (Backward)
        Else
            frmMain.Real.SetPosition (lBegin)
        End If
    End If
End Sub
Sub DoForward()
    If frmMain.Real.GetPlayState = 3 Or frmMain.Real.GetPlayState = 5 Then
        Dim Forward As Long
        Forward = frmMain.Real.GetPosition + 1000
        If Forward <= lEnd Then
            frmMain.Real.SetPosition (Forward)
        Else
            frmMain.Real.SetPosition (lEnd)
        End If
    End If
End Sub
Sub DoBackward5s()
    If frmMain.Real.GetPlayState = 3 Or frmMain.Real.GetPlayState = 5 Then
        Dim Backward5s As Long
        Backward5s = frmMain.Real.GetPosition - 5000
        If Backward5s >= lBegin Then
            frmMain.Real.SetPosition (Backward5s)
        Else
            frmMain.Real.SetPosition (lBegin)
        End If
    End If
End Sub
Sub DoForward5s()
    If frmMain.Real.GetPlayState = 3 Or frmMain.Real.GetPlayState = 5 Then
        Dim Forward5s As Long
        Forward5s = frmMain.Real.GetPosition + 5000
        If Forward5s <= lEnd Then
            frmMain.Real.SetPosition (Forward5s)
        Else
            frmMain.Real.SetPosition (lEnd)
        End If
    End If
End Sub
