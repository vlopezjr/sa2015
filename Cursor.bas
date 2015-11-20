Attribute VB_Name = "Cursor"
Option Explicit

Private m_lWaitLevels As Long


Public Sub SetWaitCursor(ByVal bFlag As Boolean)
    If bFlag Then
        If m_lWaitLevels = 0 Then
            Screen.MousePointer = vbHourglass
        End If
        m_lWaitLevels = m_lWaitLevels + 1
    Else
        If m_lWaitLevels > 0 Then
            m_lWaitLevels = m_lWaitLevels - 1
            If m_lWaitLevels = 0 Then
                Screen.MousePointer = vbDefault
            End If
        End If
    End If
End Sub


Public Sub SuppressWaitCursor(ByVal bFlag As Boolean)
    If bFlag Then
        Screen.MousePointer = vbDefault
    Else
        If m_lWaitLevels = 0 Then
            Screen.MousePointer = vbDefault
        Else
            Screen.MousePointer = vbHourglass
        End If
    End If
End Sub


Public Sub ClearWaitCursor()
    m_lWaitLevels = 0
    Screen.MousePointer = vbDefault
End Sub


