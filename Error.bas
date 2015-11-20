Attribute VB_Name = "modError"
'====================================================================
'
' Module: modError.bas
'
' Purpose: Utility functions for CPC error handling
'
' Dependencies: To use this module, you also need the following files:
'   ErrorInfo.cls - Error information structure
'   Tokenizer.cls - text parsing class
'
' Author: Jay Cincotta
'
' Theory of Operation:
'   The error handling strategy for the CPC business system calls for
'   an ability to build a trace of messages as errors propogate upward
'   through multiple layers of software.  Such error traces provide
'   detailed information on where and why errors occur but the amount
'   of information may be overwhelming for most end-users.  Consequently,
'   it is desirable to allow for the possibility that the most useful
'   message to an end-user may be neither the highest level error (which
'   may be so generic as to be useless) or the lowest level error (which
'   may be so technical in its description that it is meaningless.
'
'   The strategy we use to decide on the most useful end-user message
'   is to display the highest-level error description that is assigned
'   a specific CPC error number.  The list of numbers is defined by the
'   CPCErrorNumber enumeration which also contains the following magic values:
'
'       zErrMin         'all specific CPC error numbers are > zErrMin
'       zErrMax         'all specific CPC error numbers are < zErrMax
'       zErrUnexpected  'special value indicating an unanticipated error
'                       'This value is returned by the UnexpectedError
'                       'procedure which is usually called as the default
'                       'action in an error handler after dealing with
'                       'all error conditions that are expected at that level.
'       zErrPassThrough 'This value is returned by the PassError procedure.
'                       'It indicates that a more useful error exists
'                       'at a lower level.
'====================================================================

Option Explicit

'This is a magic value that is stored in the Err.Source
'field to indicate that the Err.Description field contains
'an error trace.
Private Const m_skErrorSubsystem = "Error Subsystem"

'constants for parsing error information
Private Const m_skErrorRowDelimiter = "@" & vbCrLf
Private Const m_skErrorColDelimiter = "|"

'Error codes used by the CPC business system
Public Enum CPCErrorNumber
    ErrSystem = 1000
    ErrDatabase = 2000
    ErrLogic = 3000
    ErrConfiguration = 4000

    zErrMin = 999
    zErrMax = 9999
    zErrUnexpected = -1
    zErrPassThrough = -2
End Enum


'=============================================================================
' Public Sub UnexpectedError
'
' Purpose: Add a message to the error trace indicating an unexpected error
'=============================================================================
Public Sub UnexpectedError( _
    ByRef i_sSource As String, _
    ByRef i_sProcedure As String, _
    Optional ByRef i_sDescription As String _
)
    RaiseError zErrUnexpected, i_sSource, i_sProcedure, i_sDescription
End Sub


'=============================================================================
' Public Sub PassError
'
' Purpose: Add a pass-through error to the error trace
'=============================================================================
Public Sub PassError( _
    ByRef i_sSource As String, _
    ByRef i_sProcedure As String, _
    Optional ByRef i_sDescription As String _
)
    RaiseError zErrPassThrough, i_sSource, i_sProcedure, i_sDescription
End Sub



'=============================================================================
' Public Sub RaiseError
'
' Purpose: Raise an error using the error trace mechanism
'=============================================================================
Public Sub RaiseError( _
    ByVal i_eNumber As CPCErrorNumber, _
    ByRef i_sSource As String, _
    ByRef i_sProcedure As String, _
    ByRef i_sDescription As String _
)
    Dim sError As String

    sError = FormatError(i_eNumber, i_sSource, i_sProcedure, i_sDescription)

    With Err

        If .number = 0 Then
            'This is the first level of the error trace
            .Description = sError
            .number = 1
        ElseIf .Source <> m_skErrorSubsystem Then
            'An error exists, but this is the first level where it's
            'being incorporated into an error trace.  So, retain the
            'existing message and supplement it with this higher level
            'error information
            .Description = sError & m_skErrorRowDelimiter _
                         & FormatError(.number, .Source, "", .Description)
            .number = 2
        Else
            'Add another level to the error trace
            .Description = sError & m_skErrorRowDelimiter _
                         & .Description
            .number = .number + 1
       End If

        .Raise .number, m_skErrorSubsystem, .Description
    End With
End Sub


'=============================================================================
' Public Function GetErrorTrace As Collection
'
' Purpose: Return a collection of ErrorInfo objects representing the error trace
'=============================================================================
Public Function GetErrorTrace() As Collection
    Dim oError As ErrorInfo
    Dim oRow As Tokenizer
    Dim oCol As Tokenizer

    Set GetErrorTrace = New Collection
    If Err.Source = m_skErrorSubsystem Then
        Set oRow = New Tokenizer
        Set oCol = New Tokenizer
        oRow.Delimiter = m_skErrorRowDelimiter
        oCol.Delimiter = m_skErrorColDelimiter

        oRow.ParseString = Err.Description
        While Not oRow.Done
            Set oError = New ErrorInfo
            oCol.ParseString = oRow.GetNextToken
            oError.number = oCol.GetNextToken
            oError.Source = oCol.GetNextToken
            oError.Description = oCol.GetNextToken
            GetErrorTrace.Add oError
        Wend
    Else
        Set oError = New ErrorInfo
        oError.number = Err.number
        oError.Source = Err.Source
        oError.Description = Err.Description
        GetErrorTrace.Add oError
    End If
End Function


'=============================================================================
' Public Function GetTopError As ErrorInfo
'
' Purpose: Return the highest level error in the error trace
'=============================================================================
Public Function GetTopError() As ErrorInfo
    Dim oError As ErrorInfo
    Dim oRow As Tokenizer
    Dim oCol As Tokenizer

    Set oError = New ErrorInfo
    If Err.Source = m_skErrorSubsystem Then
        Set oRow = New Tokenizer
        Set oCol = New Tokenizer
        oRow.Delimiter = m_skErrorRowDelimiter
        oCol.Delimiter = m_skErrorColDelimiter
        oRow.ParseString = Err.Description
        oCol.ParseString = oRow.GetNextToken
        oError.number = oCol.GetNextToken
        oError.Source = oCol.GetNextToken
        oError.Description = oCol.GetNextToken
    Else
        oError.number = Err.number
        oError.Source = Err.Source
        oError.Description = Err.Description
    End If
    Set GetTopError = oError
End Function


'=============================================================================
' Public Function GetMainError As ErrorInfo
'
' Purpose: Return the most useful error from the error trace
'=============================================================================
Public Function GetMainError() As ErrorInfo
    Dim oError As ErrorInfo
    Dim oRow As Tokenizer
    Dim oCol As Tokenizer
    
    Set oError = New ErrorInfo
    If Err.Source = m_skErrorSubsystem Then
        Set oRow = New Tokenizer
        Set oCol = New Tokenizer
        oRow.Delimiter = m_skErrorRowDelimiter
        oCol.Delimiter = m_skErrorColDelimiter
        oRow.ParseString = Err.Description
        Do Until oRow.Done Or IsCPCError(oError)
            oCol.ParseString = oRow.GetNextToken
            oError.number = oCol.GetNextToken
            oError.Source = oCol.GetNextToken
            oError.Description = oCol.GetNextToken
        Loop
    Else
        oError.number = Err.number
        oError.Source = Err.Source
        oError.Description = Err.Description
    End If

    Set GetMainError = oError
End Function


'=============================================================================
' Private Function FormatError As String
'
' Purpose: Format error information as a string to include in an error trace
'=============================================================================
Private Function FormatError( _
    ByVal i_lNumber As Long, _
    ByRef i_sSource As String, _
    ByRef i_sProcedure As String, _
    ByRef i_sDescription As String _
) As String
    FormatError = i_lNumber & m_skErrorColDelimiter _
                & i_sSource & ":" & i_sProcedure & m_skErrorColDelimiter _
                & i_sDescription
End Function


'=============================================================================
' Private Function IsCPCError As Boolean
'
' Purpose: Determine if an error is a defined CPC error or not
'=============================================================================
Private Function IsCPCError(ByRef oError As ErrorInfo) As Boolean
    Dim lErrorNumber As Long

    lErrorNumber = oError.number
    
    If lErrorNumber > CPCErrorNumber.zErrMin _
    And lErrorNumber <> CPCErrorNumber.zErrMin Then
        IsCPCError = True
    End If
End Function


'=============================================================================
' Public Sub RaiseError
'
' Purpose: Raise error with all pertinent tracing info
'=============================================================================
Public Sub ThrowError(ByVal i_sInModule As String, _
                        ByVal i_sInProcedure As String, _
                        ByVal i_sDescription As String)
    Const ERROR_LITERAL = " encountered error "
    If Len(i_sDescription) > 0 Then
        If InStr(Err.Description, ERROR_LITERAL) = 0 Then
            Err.Description = i_sDescription & ERROR_LITERAL & Err.number & " (" & Err.Description & ")"
        End If
    End If
    
    Err.Raise Err.number, SetErrSource(i_sInModule, i_sInProcedure), Err.Description
End Sub

'=============================================================================
' Private Function SetErrSource As String
'
' Purpose: Form error source info, to include app, user, computer, version info
'=============================================================================
Public Function SetErrSource(ByVal i_sInModule As String, ByVal i_sInProcedure As String, Optional ByVal i_bCreateStack As Boolean = True) As String
    SetErrSource = App.EXEName & "." & i_sInModule & "." & i_sInProcedure & "@" & Environ("COMPUTERNAME") & ":" & Environ("USERNAME") & "[version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    If i_bCreateStack Then
        If InStr(Err.Source, ".") > 0 Then
            SetErrSource = Err.Source & "->" & vbCrLf & SetErrSource
        End If
    End If
End Function

