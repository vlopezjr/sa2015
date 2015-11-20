Attribute VB_Name = "Registry"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"37937E5E029A"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   File: Registry.bas
'   Created By: Jay Cincotta
'   Date: April 5, 1999
'
'   Purpose: Utility functions for manipulating the Windows NT Registry.
'
'   Description Of Operation:
'       This module provides a number of utility functions the provide a
'       more convenient interface to the Win32 API for working with the
'       Windows NT Registry.
'
'   Summary:
'       ReadRegStringValue - Read an existing string value from the registry.
'       ReadRegNumberValue - Read an existing numeric value from the registry.
'       GetRegStringValue - Read a string value from the registry,
'                           creating it if needed.
'       GetRegNumberValue - Read a numeric value from the registry,
'                           creating it if needed.
'       PutRegStringValue - Write a string value to the registry.
'       PutRegNumberValue - Write a numeric value to the registry.
'       ReadRegStringDefault - Read the default string value for a registry key
'       WriteRegStringDefault - Write the default string value for a registry key
'   Maintenance Notes:
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Set environment flags
Option Base 1       ' Force array subscripts to begin with 1 rather than 0
Option Explicit     ' Force explicit declaration of variables
Option Compare Text ' All case-insensitive text comparisons

Private Const m_skSource = "Registry"

' pre-defined handles to the standard registry hives
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003

' constants for various Win32 API magic numbers
Private Const ERROR_NONE = 0
Private Const ERROR_BADDB = 1
Private Const ERROR_BADKEY = 2
Private Const ERROR_CANTOPEN = 3
Private Const ERROR_CANTREAD = 4
Private Const ERROR_CANTWRITE = 5
Private Const ERROR_OUTOFMEMORY = 6
Private Const ERROR_ARENA_TRASHED = 7
Private Const ERROR_ACCESS_DENIED = 8
Private Const ERROR_INVALID_PARAMETERS = 87
Private Const ERROR_NO_MORE_ITEMS = 259
Private Const KEY_ALL_ACCESS = &H3F
Private Const REG_OPTION_NON_VOLATILE = 0
Private Const REG_SZ As Long = 1
Private Const REG_DWORD As Long = 4

' Declarations for Win32 API functions
Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long
Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long
Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long
Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegQueryValue Lib "advapi32.dll" Alias "RegQueryValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal lpValue As String, lpcbValue As Long) As Long
Public Declare Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long

Private Const kErrorInternal = vbError + 100 'arbitrary error number


'************************************************************
' Public Function: ReadRegStringValue
'
' Purpose: Read an existing string value from the registry.
'
' Description:
'       If the string value is successfully read, pass zero
'       as the function return value and return the string.
'       Non-zero function return indicates that the string
'       value could not be read from the registry.
'
' Parameters:
'       i_lhKey - predefined handle of registry hive
'       i_sSubKey - registry key
'       i_sValueName - name of registry value
'       o_sValue - string value read from registry
'
' Returns:
'       Zero on success.  Non-zero indicates string not read.
'************************************************************
Public Function ReadRegStringValue(ByVal i_lhKey As Long, ByVal i_sSubKey As String, ByVal i_sValueName As String, ByRef o_sValue As String) As Long
    Dim lhKey As Long         'handle to the key
    Dim lRetVal As Long
    Dim cch As Long
    Dim lType As Long

    'In the case of any error, log it and get out
    On Error GoTo ErrorHandler

    lRetVal = RegOpenKeyEx(i_lhKey, _
                           i_sSubKey, _
                           0, _
                           KEY_ALL_ACCESS, _
                           lhKey)

    If lRetVal = ERROR_NONE Then
        ' Determine the size and type of data to be read
        lRetVal = RegQueryValueExNULL(lhKey, _
                                      i_sValueName, _
                                      0&, _
                                      lType, _
                                      0&, _
                                      cch)
        If lRetVal = ERROR_NONE Then
            If lType = REG_SZ Then
                ' Allocate a zero-filled buffer of the proper length
                o_sValue = String(cch, 0)
                lRetVal = RegQueryValueExString(lhKey, _
                                                i_sValueName, _
                                                0&, _
                                                lType, _
                                                o_sValue, _
                                                cch)
                If lRetVal = ERROR_NONE Then
                    ' If the function succeeded, return the string value
                    o_sValue = Left$(o_sValue, cch - 1)
                End If
            Else
                LogDB.LogEvent m_skSource, "ReadRegStringValue", "Warning: Converted registry value to string: " & i_sSubKey & "\" & i_sValueName

                lRetVal = RegDeleteValue(lhKey, i_sValueName)
                lRetVal = ERROR_BADKEY 'force lretval <> ERROR_NONE
            End If
        End If
    End If

NormalExit:
    RegCloseKey (lhKey)
    ReadRegStringValue = lRetVal
    Exit Function

ErrorHandler:

    LogDB.LogError m_skSource, "ReadRegStringValue", "", Err.Source, Err.Number, Err.Description

    GoTo NormalExit
End Function


'************************************************************
' Public Function: ReadRegNumberValue
'
' Purpose: Read an existing DWORD value from the registry.
'
' Description:
'       If the value is successfully read, pass zero
'       as the function return value and return the number.
'       Non-zero function return indicates that the string
'       value could not be read from the registry.
'
' Parameters:
'       i_lhKey - predefined handle of registry hive
'       i_sSubKey - registry key
'       i_sValueName - name of registry value
'       o_lValue - long value read from registry
'
' Returns:
'       Zero on success.  Non-zero indicates value not read.
'************************************************************
Public Function ReadRegNumberValue(ByVal i_lhKey As Long, ByVal i_sSubKey As String, ByVal i_sValueName As String, ByRef o_lValue As Long) As Long
    Dim lhKey As Long         'handle to the key
    Dim lRetVal As Long
    Dim cch As Long
    Dim lType As Long

    'In the case of any error, log it and get out
    On Error GoTo ErrorHandler

    lRetVal = RegOpenKeyEx(i_lhKey, _
                           i_sSubKey, _
                           0, _
                           KEY_ALL_ACCESS, _
                           lhKey)

    If lRetVal = ERROR_NONE Then
        ' Determine the size and type of data to be read
        lRetVal = RegQueryValueExNULL(lhKey, _
                                      i_sValueName, _
                                      0&, _
                                      lType, _
                                      0&, _
                                      cch)
        If lRetVal = ERROR_NONE Then
            If lType = REG_DWORD Then
                lRetVal = RegQueryValueExLong(lhKey, _
                                              i_sValueName, _
                                              0&, _
                                              lType, _
                                              o_lValue, _
                                              cch)
            Else
                LogDB.LogEvent m_skSource, "ReadRegNumberValue", "Warning: Converted registry value to DWORD: " & i_sSubKey & "\" & i_sValueName

                lRetVal = RegDeleteValue(lhKey, i_sValueName)
                lRetVal = ERROR_BADKEY 'force lretval <> ERROR_NONE
            End If
        End If
    End If

NormalExit:
    RegCloseKey (lhKey)
    ReadRegNumberValue = lRetVal
    Exit Function

ErrorHandler:

    LogDB.LogError m_skSource, "ReadRegNumberValue", "", Err.Source, Err.Number, Err.Description

    GoTo NormalExit
End Function


'************************************************************
' Public Function: GetRegStringValue
'
' Purpose: Read a string value from the registry, creating it if needed.
'
' Description:
'       This function is determined to do its job.  If the
'       requested value is not already in the registry, it
'       will create the value and key as necessary.  In this
'       case, it will assign the default value passed by the
'       caller.  If the requested value exists but is of the
'       wrong type, this function will overwrite the existing
'       value with the default string value passed by the caller.
'
'       Errors will be logged, but this function will return
'       with no error code.
'
' Parameters:
'       i_lhKey - predefined handle of registry hive
'       i_sSubKey - registry key
'       i_sValueName - name of registry value
'       i_sDefaultValue - default string value to store in registry
'                         if value not already present
' Returns:
'       String value from registry or default value on error
'************************************************************
Public Function GetRegStringValue(ByVal i_lhKey As Long, ByVal i_sSubKey As String, ByVal i_sValueName As String, ByVal i_sDefaultValue As String) As String
    Dim lhKey As Long         'handle to the key
    Dim lRetVal As Long         'result of the RegCreateKeyEx function
    Dim cch As Long
    Dim lType As Long
    Dim sValue As String
    Dim sMessage As String

    'This function should return a string value no matter what
    'so begin by initializing the return value with the
    'default value passed in by the caller.
    GetRegStringValue = i_sDefaultValue

    'In the case of any error, log it and get out
    On Error GoTo ErrorHandler

    'The RegCreateKeyEx API function will open the key if it
    'exists or create it if it does not.
    lRetVal = RegCreateKeyEx(i_lhKey, _
                             i_sSubKey, _
                             0&, _
                             vbNullString, _
                             REG_OPTION_NON_VOLATILE, _
                             KEY_ALL_ACCESS, _
                             0&, _
                             lhKey, _
                             lRetVal)

    If lRetVal = ERROR_NONE Then
        ' Determine the size and type of data to be read
        lRetVal = RegQueryValueExNULL(lhKey, _
                                      i_sValueName, _
                                      0&, _
                                      lType, _
                                      0&, _
                                      cch)
        If lRetVal = ERROR_NONE Then
            If lType = REG_SZ Then
                ' Allocate a zero-filled buffer of the proper length
                sValue = String(cch, 0)
                lRetVal = RegQueryValueExString(lhKey, _
                                                i_sValueName, _
                                                0&, _
                                                lType, _
                                                sValue, _
                                                cch)
                If lRetVal = ERROR_NONE Then
                    ' If the function succeeded, return the string value
                    GetRegStringValue = Left$(sValue, cch - 1)
                End If
            Else
                LogDB.LogEvent m_skSource, "GetRegStringValue", "Warning: Converted registry value to string: " & i_sSubKey & "\" & i_sValueName

                lRetVal = RegDeleteValue(lhKey, i_sValueName)
                lRetVal = -1 'force lretval <> ERROR_NONE
            End If
        End If
    End If

    If lRetVal <> ERROR_NONE Then
        ' Upon error, try to create the value
        sValue = i_sDefaultValue & chr$(0)
        lRetVal = RegSetValueExString(lhKey, _
                                      i_sValueName, _
                                      0&, _
                                      REG_SZ, _
                                      sValue, _
                                      Len(sValue))
        If lRetVal <> ERROR_NONE Then
            Err.Raise kErrorInternal, "GetRegStringValue", _
                      "Could not read or create registry string value: " _
                      & i_sSubKey & "\" & i_sValueName
        End If
    End If

NormalExit:
    RegCloseKey (lhKey)
    Exit Function

ErrorHandler:

    sMessage = "SubKey=" & i_sSubKey & ", ValueName=" & i_sValueName & ", DefaultValue=" & i_sDefaultValue

    LogDB.LogError m_skSource, "GetRegStringValue", sMessage, Err.Source, Err.Number, Err.Description

    GoTo NormalExit
End Function


'************************************************************
' Public Function: GetRegNumberValue
'
' Purpose: Read a DWORD value from the registry, creating it if needed.
'
' Description:
'       This function is determined to do its job.  If the
'       requested value is not already in the registry, it
'       will create the value and key as necessary.  In this
'       case, it will assign the default value passed by the
'       caller.  If the requested value exists but is of the
'       wrong type, this function will overwrite the existing
'       value with the default string value passed by the caller.
'
'       Errors will be logged, but this function will return
'       with no error code.
'
' Parameters:
'       i_lhKey - predefined handle of registry hive
'       i_sSubKey - registry key
'       i_sValueName - name of registry value
'       i_lDefaultValue - default value to store in registry
'                         if value not already present
' Returns:
'       Value from registry or default value on error
'************************************************************
Public Function GetRegNumberValue(ByVal i_lhKey As Long, ByVal i_sSubKey As String, ByVal i_sValueName As String, ByVal i_lDefaultValue As Long) As Long
    Dim lhKey As Long         'handle to the key
    Dim lRetVal As Long         'result of the RegCreateKeyEx function
    Dim cch As Long
    Dim lType As Long
    Dim lValue As Long

    'This function should return a value no matter what
    'so begin by initializing the return value with the
    'default value passed in by the caller.
    GetRegNumberValue = i_lDefaultValue

    'In the case of any error, log it and get out
    On Error GoTo ErrorHandler

    'The RegCreateKeyEx API function will open the key if it
    'exists or create it if it does not.
    lRetVal = RegCreateKeyEx(i_lhKey, _
                             i_sSubKey, _
                             0&, _
                             vbNullString, _
                             REG_OPTION_NON_VOLATILE, _
                             KEY_ALL_ACCESS, _
                             0&, _
                             lhKey, _
                             lRetVal)

    If lRetVal = ERROR_NONE Then
        ' Determine the size and type of data to be read
        lRetVal = RegQueryValueExNULL(lhKey, _
                                      i_sValueName, _
                                      0&, _
                                      lType, _
                                      0&, _
                                      cch)
        If lRetVal = ERROR_NONE Then
            If lType = REG_DWORD Then
                lRetVal = RegQueryValueExLong(lhKey, _
                                              i_sValueName, _
                                              0&, _
                                              lType, _
                                              lValue, _
                                              cch)
            Else
                LogDB.LogEvent m_skSource, "GetRegNumberValue", "Warning: Converted registry value to DWORD: " & i_sSubKey & "\" & i_sValueName

                lRetVal = RegDeleteValue(lhKey, i_sValueName)
                lRetVal = -1 'force lretval <> ERROR_NONE
            End If
        End If
    End If

    If lRetVal <> ERROR_NONE Then
        'smr - 12/2004 - default value to be set in the reg
        lValue = i_lDefaultValue

        ' Upon error, try to create the value
        lRetVal = RegSetValueExLong(lhKey, _
                                    i_sValueName, _
                                    0&, _
                                    REG_DWORD, _
                                    lValue, _
                                    4)
        If lRetVal <> ERROR_NONE Then
            Err.Raise kErrorInternal, "GetRegNumberValue", _
                      "Could not read or create registry value: " _
                      & i_sSubKey & "\" & i_sValueName
        End If
    Else
        GetRegNumberValue = lValue
    End If

NormalExit:
    RegCloseKey (lhKey)
    Exit Function

ErrorHandler:

    LogDB.LogError m_skSource, "GetRegNumberValue", "", Err.Source, Err.Number, Err.Description

    GoTo NormalExit
End Function


'************************************************************
' Public Function: PutRegStringValue
'
' Purpose: Write a string value to the registry.
'
' Description:
'       This routine is determined to do its job and will
'       create the key if necessary or even change the type
'       of an existing registry key to store the string
'       value passed by the caller into the registry.
'
' Parameters:
'       i_lhKey - predefined handle of registry hive
'       i_sSubKey - registry key
'       i_sValueName - name of registry value
'       i_sValue - string value to store in registry
'************************************************************
Public Sub PutRegStringValue(ByVal i_lhKey As Long, ByVal i_sSubKey As String, ByVal i_sValueName As String, ByVal i_sValue As String)
    Dim lhKey As Long         'handle to the key
    Dim lRetVal As Long       'result of the RegCreateKeyEx function
    Dim cch As Long
    Dim lType As Long
    Dim sValue As String

    'In the case of any error, log it and get out
    On Error GoTo ErrorHandler

    'The RegCreateKeyEx API function will open the key if it
    'exists or create it if it does not.
    lRetVal = RegCreateKeyEx(i_lhKey, _
                             i_sSubKey, _
                             0&, _
                             vbNullString, _
                             REG_OPTION_NON_VOLATILE, _
                             KEY_ALL_ACCESS, _
                             0&, _
                             lhKey, _
                             lRetVal)

    If lRetVal = ERROR_NONE Then
        ' Determine the size and type of data to be read
        lRetVal = RegQueryValueExNULL(lhKey, _
                                      i_sValueName, _
                                      0&, _
                                      lType, _
                                      0&, _
                                      cch)
        If lRetVal = ERROR_NONE Then
            If lType <> REG_SZ Then
                LogDB.LogEvent m_skSource, "PutRegStringValue", "Warning: Converted registry value to string: " & i_sSubKey & "\" & i_sValueName

                lRetVal = RegDeleteValue(lhKey, i_sValueName)
                lRetVal = -1 'force lretval <> ERROR_NONE
            End If
        End If
    End If

    sValue = i_sValue & chr$(0)
    lRetVal = RegSetValueExString(lhKey, _
                                  i_sValueName, _
                                  0&, _
                                  REG_SZ, _
                                  sValue, _
                                  Len(sValue))
    If lRetVal <> ERROR_NONE Then
        Err.Raise kErrorInternal, "PutRegStringValue", _
                  "Could not write to or create registry string value: " _
                  & i_sSubKey & "\" & i_sValueName
    End If

NormalExit:
    RegCloseKey (lhKey)
    Exit Sub

ErrorHandler:

    LogDB.LogError m_skSource, "PutRegStringValue", "", Err.Source, Err.Number, Err.Description

    GoTo NormalExit
End Sub


'************************************************************
' Public Function: PutRegNumberValue
'
' Purpose: Write a DWORD value to the registry.
'
' Description:
'       This routine is determined to do its job and will
'       create the key if necessary or even change the type
'       of an existing registry key to store the string
'       value passed by the caller into the registry.
'
' Parameters:
'       i_lhKey - predefined handle of registry hive
'       i_sSubKey - registry key
'       i_sValueName - name of registry value
'       i_lValue - long value to store in registry
'************************************************************
Public Sub PutRegNumberValue(ByVal i_lhKey As Long, ByVal i_sSubKey As String, ByVal i_sValueName As String, ByVal i_lValue As Long)
    Dim lhKey As Long         'handle to the key
    Dim lRetVal As Long       'result of the RegCreateKeyEx function
    Dim cch As Long
    Dim lType As Long
    Dim lValue As Long

    'In the case of any error, log it and get out
    On Error GoTo ErrorHandler

    'The RegCreateKeyEx API function will open the key if it
    'exists or create it if it does not.
    lRetVal = RegCreateKeyEx(i_lhKey, _
                             i_sSubKey, _
                             0&, _
                             vbNullString, _
                             REG_OPTION_NON_VOLATILE, _
                             KEY_ALL_ACCESS, _
                             0&, _
                             lhKey, _
                             lRetVal)

    If lRetVal = ERROR_NONE Then
        ' Determine the size and type of data to be read
        lRetVal = RegQueryValueExNULL(lhKey, _
                                      i_sValueName, _
                                      0&, _
                                      lType, _
                                      0&, _
                                      cch)
        If lRetVal = ERROR_NONE Then
            If lType <> REG_DWORD Then
                LogDB.LogEvent m_skSource, "PutRegNumberValue", "Warning: Converted registry value to DWORD: " & i_sSubKey & "\" & i_sValueName

                lRetVal = RegDeleteValue(lhKey, i_sValueName)
                lRetVal = -1 'force lretval <> ERROR_NONE
            End If
        End If
    End If

    lRetVal = RegSetValueExLong(lhKey, _
                                i_sValueName, _
                                0&, _
                                REG_DWORD, _
                                i_lValue, _
                                4)
    If lRetVal <> ERROR_NONE Then
        Err.Raise kErrorInternal, "PutRegNumberValue", _
                  "Could not write to or create registry value: " _
                  & i_sSubKey & "\" & i_sValueName
    End If

NormalExit:
    RegCloseKey (lhKey)
    Exit Sub

ErrorHandler:

    LogDB.LogError m_skSource, "PutRegNumberValue", "", Err.Source, Err.Number, Err.Description

    GoTo NormalExit
End Sub


'************************************************************
' Public Function: ReadRegDefaultString
'
' Purpose: Read the default string value from the registry.
'
' Description:
'       This function reads the default string value associated
'       with a registry key.
'
' Parameters:
'       i_lhKey - predefined handle of registry hive
'       i_sSubKey - registry key
'       i_sValueName - name of registry value
'       o_sValue - string value read from registry
'
' Returns:
'       Zero on success.  Non-zero indicates string not read.
'************************************************************
Public Function ReadRegDefaultString(ByVal i_lhKey As Long, ByVal i_sSubKey As String, ByRef o_sValue As String) As Long
    Dim lhKey As Long         'handle to the key
    Dim lRetVal As Long
    Dim cch As Long
    Dim lType As Long

    'In the case of any error, log it and get out
    On Error GoTo ErrorHandler

    'By passing a null string, we get the buffer size
    lRetVal = RegQueryValue(i_lhKey, _
                            i_sSubKey, _
                            vbNullString, _
                            cch)

    If lRetVal = ERROR_NONE Then
        o_sValue = String(cch, 0)
        lRetVal = RegQueryValue(i_lhKey, _
                                i_sSubKey, _
                                o_sValue, _
                                cch)
        If lRetVal = ERROR_NONE Then
            o_sValue = Left$(o_sValue, cch - 1)
        End If
    End If

NormalExit:
    ReadRegDefaultString = lRetVal
    Exit Function

ErrorHandler:

    LogDB.LogError m_skSource, "ReadRegDefaultString", "", Err.Source, Err.Number, Err.Description

    GoTo NormalExit
End Function


'************************************************************
' Public Function: WriteRegDefaultString
'
' Purpose: Write the default string value for a registry key.
'
' Description:
'       This routine is determined to do its job and will
'       create the key if necessary or even change the type
'       of an existing registry key to store the string
'       value passed by the caller into the registry.
'
' Parameters:
'       i_lhKey - predefined handle of registry hive
'       i_sSubKey - registry key
'       i_sValueName - name of registry value
'       i_sValue - string value to store in registry
'************************************************************
Public Sub WriteRegDefaultString(ByVal i_lhKey As Long, ByVal i_sSubKey As String, ByVal i_sValue As String)
    Dim sValue As String
    Dim lRetVal As Long
    Dim cch As Long


    'In the case of any error, log it and get out
    On Error GoTo ErrorHandler

    sValue = i_sValue & chr$(0)
    lRetVal = RegSetValue(i_lhKey, _
                          i_sSubKey, _
                          REG_SZ, _
                          sValue, _
                          Len(sValue))
    If lRetVal <> ERROR_NONE Then
        Err.Raise kErrorInternal, "WriteRegDefaultString", _
                  "Could not write to or create registry default string value: " _
                  & i_sSubKey
    End If

NormalExit:
    Exit Sub

ErrorHandler:

    LogDB.LogError m_skSource, "WriteRegDefaultString", "", Err.Source, Err.Number, Err.Description

    GoTo NormalExit
End Sub

