Attribute VB_Name = "ErrorUI"
Option Explicit
Option Base 1

Private Const m_skSource = "modError"

'Registry value for user-specific connection string
Private Const m_skRegValShowTrace = "ShowCallTrace"


Public Enum ErrorLevel
    elError = 1
    elWarning = 2
    elInformation = 3
    elDebug = 4
End Enum

'threshold for what messages should be logged
Private m_eLogLevel As ErrorLevel

'if set, call trace information will be displayed to user
Private m_bForceCallTrace As Boolean

'default titles for msgbox
Private Const m_skErrorTitle = "System Error"
Private Const m_skWarningTitle = "System Warning"
Private Const m_skInformationTitle = "System Information"
Private Const m_skDebugTitle = "Debug Information"

'Users may only pass in values for the VB buttons they want displayed.
'This mask is used to clear any other bits that might be set
Private Const m_lkButtonMask = &H7


Public Property Get LogLevel() As ErrorLevel
    LogLevel = m_eLogLevel
End Property


Public Property Let LogLevel(ByVal i_eLogLevel As ErrorLevel)
    m_eLogLevel = i_eLogLevel
End Property


Public Property Get ForceCallTrace() As Boolean
    ForceCallTrace = m_bForceCallTrace
End Property


Public Property Let ForceCallTrace(ByVal i_bShow As Boolean)
    m_bForceCallTrace = i_bShow
End Property

   
Public Property Get ShowCallTrace() As Boolean
    Dim lShow As Long
    
    If m_bForceCallTrace Then
        ShowCallTrace = True
    Else
'        '4/3/14 LR  This dependency on g_DB is ugly
'        'it's a temporary measure during refactoring
'        lShow = GetRegNumberValue(HKEY_CURRENT_USER, _
'                                  g_DB.m_sBaseRegKey, _
'                                  m_skRegValShowTrace, _
'                                  0)
        lShow = 0
        If lShow <> 0 Then
            ShowCallTrace = True
        End If
    End If
End Property


Public Sub DisplayError(Optional ByRef i_sTitle As String)
    DisplayMsgBox GetMainError.Description, vbOKOnly, elError, i_sTitle
End Sub

Public Sub DisplayWarning(Optional ByRef i_sTitle As String)
    DisplayMsgBox GetMainError.Description, vbOKOnly, elWarning, i_sTitle
End Sub

Public Sub DisplayInfo(Optional ByRef i_sTitle As String)
    DisplayMsgBox GetMainError.Description, vbOKOnly, elInformation, i_sTitle
End Sub

Public Sub DisplayDebug(Optional ByRef i_sTitle As String)
    DisplayMsgBox GetMainError.Description, vbOKOnly, elDebug, i_sTitle
End Sub


Public Function DisplayMsgBox( _
    ByRef i_sPrompt As String, _
    ByVal i_lButtons As Long, _
    ByVal i_eSeverity As ErrorLevel, _
    Optional ByVal i_sTitle As String = "" _
) As Long
    Dim sMessage As String
    Dim lButtons As Long
    Dim lEventSeverity As Long
    Dim sDefaultTitle As String
    
    'Enforce that only the button flags may be passed by user
    i_lButtons = i_lButtons And m_lkButtonMask

    Select Case i_eSeverity
      Case elError
        lButtons = vbCritical + i_lButtons
        sDefaultTitle = m_skErrorTitle
       lEventSeverity = vbLogEventTypeError
        
      Case elWarning
        lButtons = vbExclamation + i_lButtons
        sDefaultTitle = m_skWarningTitle
        lEventSeverity = vbLogEventTypeWarning
      
      Case elInformation
        lButtons = vbInformation + i_lButtons
        sDefaultTitle = m_skInformationTitle
        lEventSeverity = vbLogEventTypeInformation

      Case elDebug
        lButtons = i_lButtons
        sDefaultTitle = m_skDebugTitle
        lEventSeverity = vbLogEventTypeInformation
      
      Case Else
        lButtons = i_lButtons
        sDefaultTitle = App.Title
    End Select

    If Len(i_sTitle) = 0 Then
        i_sTitle = sDefaultTitle
    End If

    'format the message for display and optionally logging
    sMessage = FormatDisplayMessage(i_sPrompt)

    'log the message, if appropriate
    If i_eSeverity <= m_eLogLevel And lEventSeverity > 0 Then
        LogDB.LogEvent "modError", "DisplayMsgBox", sMessage
    End If

    'display the message box and return user selection
    If ShowCallTrace Then
        DisplayMsgBox = msg(sMessage, lButtons, i_sTitle)
    Else
        DisplayMsgBox = msg(i_sPrompt, lButtons, i_sTitle)
    End If
End Function


Public Function FormatDisplayMessage(ByRef i_sPrompt As String) As String
    Dim oError As ErrorInfo
    Dim oTrace As Collection
    Dim lErrCount As Long
    Dim lIndex As Long
    Dim sMsg As String
    
    sMsg = i_sPrompt & vbCrLf & vbCrLf & "Call Trace Details: "
    Set oTrace = GetErrorTrace
    lErrCount = oTrace.Count
    For lIndex = oTrace.Count To 1 Step -1
        Set oError = oTrace.Item(lIndex)
        With oError
            sMsg = sMsg & vbCrLf _
                 & "    " & .Source & " reports error " _
                 & .Number & " (0x" & Hex(.Number) & "):" _
                 & vbCrLf & "        " & .Description & vbCrLf
        End With
    Next
    FormatDisplayMessage = sMsg
End Function


Public Function GetErrorTitle(ByVal i_eErrorNumber As CPCErrorNumber) As String
    Select Case i_eErrorNumber
      Case CPCErrorNumber.ErrSystem
        GetErrorTitle = "System Error"
      Case CPCErrorNumber.ErrDatabase
        GetErrorTitle = "Database Error"
      Case CPCErrorNumber.ErrConfiguration
        GetErrorTitle = "Configuration Error"
      Case CPCErrorNumber.ErrLogic
        GetErrorTitle = "Logic Error"
      Case Else
        GetErrorTitle = "Unexpected Error"
    End Select
End Function


Public Sub FatalError(ByVal i_sSource As String, ByVal i_sMsg As String)
    msg "An unexpected condition occurred.  After clicking OK, OrderPad will exit." & vbCrLf _
      & "Please restart OrderPad.  If this happens again, please call the IT department." & vbCrLf & vbCrLf _
      & "Error Source: " & i_sSource & vbCrLf & "Error Details:" & vbCrLf _
      & i_sMsg, _
      vbCritical + vbOKOnly, "System Error"
    End
End Sub
