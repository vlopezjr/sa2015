Attribute VB_Name = "Help"
Option Explicit

Private m_oPopup As PopUp.CPopUp

Public Sub ShowHelp(ByVal i_sContext As String, Optional ByVal i_bModal As Boolean = False)
    On Error GoTo EH
    
    If m_oPopup Is Nothing Then
        Set m_oPopup = New PopUp.CPopUp
        m_oPopup.width = 9550
        m_oPopup.Height = 7224
    End If

    With m_oPopup
        .Caption = "SageAssistant Help"
       .PopUp g_OrderPadBaseURL & "doc/" & i_sContext & ".htm", i_bModal
    End With
    Exit Sub
EH:
    msg Err.Source & " " & Err.Number & ": " & Err.Description, vbCritical
End Sub


