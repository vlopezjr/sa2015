Attribute VB_Name = "Helpers"
Option Explicit
Option Compare Text

'Public Sub LoadCombo(cboCombo As ComboBox, rst As ADODB.Recordset, sDisplayField As String, Optional vKeyField As Variant, Optional vDfltKeyValue As Variant, Optional vWithNone As Variant)
'Public Sub SetComboByKey(cboControl As ComboBox, vKeyField As Variant, Optional vWithNone As Variant)
'Public Sub SetComboByText(cboControl As ComboBox, sText As String, Optional vWithNone As Variant)
'Public Sub SetListByText(lstControl As ListBox, sText As String, Optional vWithNone As Variant)
'Public Sub SetCheckbox(chkControl As CheckBox, IsChecked As Boolean)
'Public Sub DisableCheckbox(chkControl As CheckBox)
'Public Sub LoadImageList(ByRef i_oImageList As ImageList, ByRef o_oGridEX As GridEX)
'Public Sub LoadList(lstList As ListBox, rst As ADODB.Recordset, sDisplayField As String, Optional vKeyField As Variant, Optional vDfltKeyValue As Variant, Optional vWithNone As Variant)
'Public Function FormatInches( _
'        ByVal i_dblValue As Double, _
'        Optional ByVal i_strInchLabel As String = "", _
'        Optional ByVal i_lDenominator As Long = k_lDefaultFraction _
') As String
'Public Function CompAddr(i_sAddrName As String, _
'Public Function FormatPhoneNumber(ByVal sPhoneNbr As String, Optional sPhoneExt As Variant) As String
'Public Function FormatZipCode(ByVal sPostalText As String) As String
'Public Function FormatCaption(i_sInput As String) As String
'Public Sub TryToSetFocus(ctrl As Control)
'Public Function HandleNullValue(sValue As Variant) As String
'Public Sub ListRemoveItemByText(o_lstControl As ListBox, i_sText As String)
'Public Function ScrubText(sText As String) As String
'Public Function StripLeadingZeros(ByVal s As String) As String
'Public Function FormatBatchID(ByVal sInput As String, ByVal sMask As String) As String
'Public Sub ClearCollection(ByRef col As Collection)
'Public Function ParseString(SubStrs() As String, ByVal SrcStr As String, ByVal Delimiter As String) As Integer
'Public Function ImportString(sXML As String) As XMLNode

Private Const k_lDefaultFraction = 16

Public Sub LoadCombo(cboCombo As ComboBox, rst As ADODB.Recordset, sDisplayField As String, Optional vKeyField As Variant, Optional vDfltKeyValue As Variant, Optional vWithNone As Variant)
    Dim lIdx As Long
    
    If IsMissing(vWithNone) Then
        vWithNone = False
    End If

    cboCombo.Clear
    
    If vWithNone Then
        cboCombo.AddItem "<none>"
        cboCombo.ItemData(cboCombo.NewIndex) = 0
    End If
    
    On Error Resume Next
    rst.MoveFirst
    On Error GoTo 0
    lIdx = 0
    If Not rst.EOF Then
        With rst
            Do While Not .EOF
                cboCombo.AddItem Trim(.Fields(sDisplayField).value)
    
                If Not IsMissing(vKeyField) Then
                    cboCombo.ItemData(cboCombo.NewIndex) = .Fields(vKeyField).value
                End If
    
                If Not IsMissing(vDfltKeyValue) Then
                    If Not IsMissing(vKeyField) Then
                        If vDfltKeyValue = .Fields(vKeyField).value Then
                            lIdx = cboCombo.NewIndex
                        End If
                    Else
                        If vDfltKeyValue = Trim(.Fields(sDisplayField).value) Then
                            lIdx = cboCombo.NewIndex
                        End If
                    End If
                End If
    
                .MoveNext
            Loop
        End With
    End If
    
    If cboCombo.ListCount > 0 Then
        cboCombo.ListIndex = lIdx
    End If
End Sub


Public Sub SetComboByKey(cboControl As ComboBox, vKeyField As Variant, Optional vWithNone As Variant)
    Dim i As Long
    Dim iHoldIdx As Long

    If IsMissing(vWithNone) Then
        vWithNone = False
    End If
    
    With cboControl
        iHoldIdx = .ListIndex
    
        For i = 0 To .ListCount - 1
            If .ItemData(i) = vKeyField Then
                .ListIndex = i
                Exit Sub
            End If
        Next
    
        If vWithNone Then
            .ListIndex = 0
        Else
            .ListIndex = iHoldIdx
        End If
        
    End With
End Sub


Public Sub SetComboByText(cboControl As ComboBox, sText As String, Optional vWithNone As Variant)
    Dim i As Long
    Dim iHoldIdx As Long

    If IsMissing(vWithNone) Then
        vWithNone = False
    End If
    
    With cboControl
        iHoldIdx = .ListIndex
    
        For i = 0 To .ListCount - 1
            If .List(i) = Trim(sText) Then
                .ListIndex = i
                Exit Sub
            End If
        Next
        
        If vWithNone Then
            .ListIndex = 0
        Else
            .ListIndex = iHoldIdx
        End If
    End With
End Sub


Public Sub SetListByText(lstControl As ListBox, sText As String, Optional vWithNone As Variant)
    Dim i As Long
    Dim iHoldIdx As Long

    If IsMissing(vWithNone) Then
        vWithNone = False
    End If
    
    With lstControl
        iHoldIdx = .ListIndex
    
        For i = 0 To .ListCount - 1
            If .List(i) = Trim(sText) Then
                .ListIndex = i
                Exit Sub
            End If
        Next
        
        If vWithNone Then
            .ListIndex = 0
        Else
            .ListIndex = iHoldIdx
        End If
    End With
End Sub


Public Sub SetCheckbox(chkControl As CheckBox, IsChecked As Boolean)
    If IsChecked Then
        chkControl.value = vbChecked
    Else
        chkControl.value = vbUnchecked
    End If
End Sub


Public Sub DisableCheckbox(chkControl As CheckBox)
    With chkControl
        .Enabled = True
        .value = vbUnchecked
        .Enabled = False
    End With
End Sub


Public Sub LoadImageList(ByRef i_oImageList As ImageList, ByRef o_oGridEX As GridEX)
    Dim i As Long
    
    o_oGridEX.GridImages.Clear
    For i = 1 To i_oImageList.ListImages.Count
        o_oGridEX.GridImages.Add i_oImageList.ListImages(i).Picture
    Next
End Sub

'This has limited use
'Called by:
'   FUser.cmdRefreshUsers_Click()
'   FChooseBin.LoadNewBin()

Public Sub LoadList(lstList As ListBox, rst As ADODB.Recordset, sDisplayField As String, Optional vKeyField As Variant, Optional vDfltKeyValue As Variant, Optional vWithNone As Variant)
    Dim lIdx As Long
    
    If IsMissing(vWithNone) Then
        vWithNone = False
    End If

    lstList.Clear
    
    If vWithNone Then
        lstList.AddItem "<none>"
        lstList.ItemData(lstList.NewIndex) = 0
    End If

    With rst
        If Not .EOF Then  'guard added 7/24/02 LR
            .MoveFirst
            lIdx = 0
            Do While Not .EOF
                lstList.AddItem Trim(.Fields(sDisplayField).value)
    
                If Not IsMissing(vKeyField) Then
                    lstList.ItemData(lstList.NewIndex) = .Fields(vKeyField).value
                End If
    
                If Not IsMissing(vDfltKeyValue) Then
                    If Not IsMissing(vKeyField) Then
                        If vDfltKeyValue = .Fields(vKeyField).value Then
                            lIdx = lstList.NewIndex
                            lstList.Selected(lIdx) = True
                        End If
                    Else
                        If vDfltKeyValue = Trim(.Fields(sDisplayField).value) Then
                            lIdx = lstList.NewIndex
                        End If
                    End If
                End If
    
                .MoveNext
            Loop
        End If
    End With

    If lstList.ListCount > 0 Then
        lstList.ListIndex = lIdx
    End If
End Sub


Public Function FormatInches( _
        ByVal i_dblValue As Double, _
        Optional ByVal i_strInchLabel As String = "", _
        Optional ByVal i_lDenominator As Long = k_lDefaultFraction _
) As String
    Dim strWhole As String
    Dim strFraction As String
    
    strWhole = Fix(i_dblValue)
    strFraction = Fraction(i_dblValue, i_lDenominator)

    If Len(strWhole) > 0 And Len(strFraction) > 0 Then
        FormatInches = strWhole & "-" & strFraction
    Else
        FormatInches = strWhole & strFraction 'one of these peices is an empty string
    End If
    FormatInches = FormatInches & i_strInchLabel
End Function


Private Function Fraction(ByVal i_dblValue As Double, ByVal i_lDenominator As Long) As String
    Dim lNumerator As Long
    Dim lDenominator As Long

    lNumerator = (i_dblValue * i_lDenominator) - (Fix(i_dblValue) * i_lDenominator)
    lDenominator = i_lDenominator

    'Reduce fraction
    While lNumerator > 0 And IsEven(lNumerator) And IsEven(lDenominator)
        lNumerator = lNumerator / 2
        lDenominator = lDenominator / 2
    Wend

    If lNumerator > 0 Then
        Fraction = lNumerator & "/" & lDenominator
    End If
End Function


Private Function IsEven(ByVal i_dblValue As Double) As Boolean
    IsEven = (CLng(i_dblValue / 2) = i_dblValue / 2)
End Function


'Build a complete address for display (with embedded CRLF)

Public Function CompAddr(i_sAddrName As String, _
    i_sAddr1 As String, _
    i_sAddr2 As String, _
    i_sCity As String, _
    i_sState As String, _
    i_sZip As String, _
    i_sCountry As String) As String
    
    Dim sTemp As String
    
    sTemp = sTemp & FormatCaption(Trim(i_sAddrName))
    
    If Trim(i_sAddr1) <> "" Then
        sTemp = sTemp & vbCrLf & FormatCaption(Trim(i_sAddr1))
    End If
    
    If Trim(i_sAddr2) <> "" Then
        sTemp = sTemp & vbCrLf & FormatCaption(Trim(i_sAddr2))
    End If
    
    If Trim(i_sCity) <> "" Then
        sTemp = sTemp & vbCrLf & Trim(i_sCity)
    End If
    
    If Trim(i_sState) <> "" Then
        sTemp = sTemp & ", " & Trim(i_sState)
    End If
    
    If Trim(i_sZip) <> "" Then
        sTemp = sTemp & " " & FormatZipCode(i_sZip)
    End If
    
    If Trim(i_sCountry) <> "USA" And Trim(i_sCountry) <> "" Then
        sTemp = sTemp & vbCrLf & Trim(i_sCountry)
    End If
    
    CompAddr = sTemp
End Function

'TODO: This doesn't properly handle the case where sPhoneNbr contains too many characters.

Public Function FormatPhoneNumber(ByVal sPhoneNbr As String, Optional sPhoneExt As Variant) As String
    Dim sTemp As String
    
    sPhoneNbr = Trim$(sPhoneNbr)
    
    '10/19/05 LR
    'if the string already has punctuation characters, remove them first
    sPhoneNbr = Replace(sPhoneNbr, "(", "")
    sPhoneNbr = Replace(sPhoneNbr, ")", "")
    sPhoneNbr = Replace(sPhoneNbr, "-", "")
    
    Select Case Len(sPhoneNbr)
    Case 10
        sTemp = "(" + Mid$(sPhoneNbr, 1, 3) + ") " + Mid$(sPhoneNbr, 4, 3) + "-" + Mid$(sPhoneNbr, 7)
    Case 7
        sTemp = Mid$(sPhoneNbr, 1, 3) + "-" + Mid$(sPhoneNbr, 4)
    Case Else
        sTemp = sPhoneNbr
    End Select

    If Not IsMissing(sPhoneExt) Then
        sPhoneExt = Trim$(sPhoneExt)
        If Len(sPhoneExt) > 0 Then
            sTemp = sTemp + " x" + sPhoneExt
        End If
    End If
    
    FormatPhoneNumber = Trim(sTemp)
End Function


Public Function FormatZipCode(ByVal sPostalText As String) As String
    Select Case Len(Trim(sPostalText))
    Case 5:
        FormatZipCode = Trim$(sPostalText)
    Case 6:
        FormatZipCode = Left$(Trim(sPostalText), 3) & " " & Right$(Trim(sPostalText), 3)
    Case 9:
        FormatZipCode = Left$(Trim(sPostalText), 5) & "-" & Right$(Trim(sPostalText), 4)
    Case Else:
        FormatZipCode = Trim$(sPostalText)
    End Select
End Function


Public Function FormatCaption(i_sInput As String) As String
    FormatCaption = Replace(i_sInput, "&", "&&")
End Function


Public Sub TryToSetFocus(ctrl As Control)
    On Error Resume Next 'ignore errors if can't set focus now
    ctrl.SetFocus
End Sub


Public Function HandleNullValue(sValue As Variant) As String
    If IsNull(sValue) Then
        HandleNullValue = ""
    Else
        HandleNullValue = CStr(sValue)
    End If
End Function


Public Sub ListRemoveItemByText(o_lstControl As ListBox, i_sText As String)
    Dim i As Long
    With o_lstControl
        For i = 0 To .ListCount
            If Trim(.List(i)) = Trim(i_sText) Then
                .RemoveItem i
                Exit Sub
            End If
        Next
    End With
End Sub


Public Function ScrubText(sText As String) As String
    Dim sTemp As String
    Dim bClean As Boolean
    Dim sarBadChar(1 To 11) As String * 1
    Dim X As Integer
    
    sTemp = sText
    sTemp = UCase(sTemp)
    
    sarBadChar(1) = " "
    sarBadChar(2) = "-"
    sarBadChar(3) = "/"
    sarBadChar(4) = "\"
    sarBadChar(5) = """"
    sarBadChar(6) = "'"
    sarBadChar(7) = vbTab
    sarBadChar(8) = "*"
    sarBadChar(9) = "?"
    sarBadChar(10) = "%"
    sarBadChar(11) = "<none>"
    
    bClean = False
    Do Until bClean = True
        bClean = True
        For X = 1 To 11
            If InStr(1, sTemp, sarBadChar(X), vbTextCompare) > 0 Then
                sTemp = Replace(sTemp, sarBadChar(X), "")
                bClean = False
            End If
        Next X
    Loop
    
    ScrubText = sTemp
End Function


Public Function StripLeadingZeros(ByVal s As String) As String
   While Left$(s, 1) < "1"
      s = Right$(s, Len(s) - 1)
   Wend
   StripLeadingZeros = s
End Function


'moved here from FBilling

Public Function FormatBatchID(ByVal sInput As String, ByVal sMask As String) As String
    sInput = PrepSQLText(sInput)
    FormatBatchID = Left$(sMask, Len(sMask) - Len(sInput)) & sInput
End Function

'*** 6/18/09
' Called by Notification.cls for use with Contact Manager.
Public Sub ClearCollection(ByRef col As Collection)
    Dim i As Integer
    Dim sErrMsg As String
    
    On Error GoTo EH
    
    For i = 1 To col.Count
        col.Remove 1
    Next i
    Exit Sub
'NOTE: don't throw a msgbox in DLL
EH:
    MsgBox Err.number & ": " & Err.Description & vbCrLf _
        & sErrMsg & vbCrLf _
        & "i=" & i & " col.count=" & col.Count
End Sub


Public Function ParseString(SubStrs() As String, ByVal SrcStr As String, ByVal Delimiter As String) As Integer

    ' Dimension variables:
    ReDim SubStrs(0) As String
    Dim CurPos As Long
    Dim NextPos As Long
    Dim DelLen As Integer
    Dim nCount As Integer
    Dim TStr As String

    ' Add delimiters to start and end of string to make loop simpler:
    SrcStr = Delimiter & SrcStr & Delimiter
    ' Calculate the delimiter length only once:
    DelLen = Len(Delimiter)
    ' Initialize the count and position:
    nCount = 0
    CurPos = 1
    NextPos = InStr(CurPos + DelLen, SrcStr, Delimiter)

    ' Loop searching for delimiters:
    Do Until NextPos = 0
       ' Extract a sub-string:
       TStr = Mid$(SrcStr, CurPos + DelLen, NextPos - CurPos - DelLen)
       ' Increment the sub string counter:
       nCount = nCount + 1
       ' Add room for the new sub-string in the array:
       ReDim Preserve SubStrs(nCount) As String
       ' Put the sub-string in the array:
       SubStrs(nCount) = Trim$(TStr)
       ' Position to the last found delimiter:
       CurPos = NextPos
       ' Find the next delimiter:
       NextPos = InStr(CurPos + DelLen, SrcStr, Delimiter)
    Loop

    ' Return the number of sub-strings found:
    ParseString = nCount

End Function



Public Function ImportString(sXML As String) As JDMPDXML.XMLNode
    On Error GoTo ErrorHandler
    
    Dim oXML As JDMPDXML.XMLNode
    
    Set oXML = New JDMPDXML.XMLNode
    oXML.ImportString sXML
    Set ImportString = oXML
    Exit Function
    
ErrorHandler:
    Err.Raise Err.number, "ImportString", "Error in importing XMLNode"
End Function



'Are BranchID & WhseID different?

Public Function GetWhseIDFromWhseKey(WhseKey As Long) As String
    g_rstWhses.Filter = "WhseKey = " & WhseKey
    GetWhseIDFromWhseKey = g_rstWhses.Fields("WhseID").value
    g_rstWhses.Filter = adFilterNone
End Function


'synonomus
Public Function WhseKeyToID(WhseKey As Long) As String
    WhseKeyToID = GetWhseIDFromWhseKey(WhseKey)
End Function


Public Function GetWhseKeyFromUserKey(UserKey As Long)
    GetWhseKeyFromUserKey = GetWhseKeyFromBranchID(GetBranchIDFromUserKey(UserKey))
End Function


Public Function GetWhseKeyFromBranchID(BranchID As String) As Long
    g_rstWhses.Filter = "WhseID = '" & BranchID & "'"
    GetWhseKeyFromBranchID = g_rstWhses.Fields("WhseKey").value
    g_rstWhses.Filter = adFilterNone
End Function


Public Function GetBranchIDFromUserKey(UserKey As Long) As String
    g_rstUsers.Filter = "Userkey=" & UserKey
    GetBranchIDFromUserKey = g_rstUsers.Fields("BranchID").value
    g_rstUsers.Filter = adFilterNone
End Function


Public Function GetWhseDescriptionFromWhseKey(WhseKey As Long) As String
    g_rstWhses.Filter = "WhseKey = " & WhseKey
    GetWhseDescriptionFromWhseKey = g_rstWhses.Fields("Description").value
    g_rstWhses.Filter = adFilterNone
End Function


Public Function GetShipMethKeyFromBranchID(BranchID As String) As Long
    g_rstShipVia.Filter = "ShipMethID = '" & BranchID & "'"
    If Not g_rstShipVia.EOF Then
        GetShipMethKeyFromBranchID = Trim(g_rstShipVia.Fields("ShipMethKey").value)
    End If
    g_rstShipVia.Filter = adFilterNone
End Function


Public Function CreateMISC_CustID() As String
    With g_rstUsers
        .Filter = "Userkey=" & GetUserKey
        CreateMISC_CustID = .Fields("BranchID").value & "-MISC"
        .Filter = adFilterNone
    End With
End Function

Public Function ShipMethIDtoKey(ByVal i_sShipMethID As String) As Long
    With g_rstShipVia
        .Filter = "ShipMethID='" & i_sShipMethID & "'"
        If Not .EOF Then
            ShipMethIDtoKey = .Fields("ShipMethKey").value
        End If
        .Filter = adFilterNone
    End With
End Function


Public Function ShipMethKeytoID(ByVal i_sShipMethKey As Long) As String
    With g_rstShipVia
        .Filter = "ShipMethKey=" & i_sShipMethKey
        If Not .EOF Then
            ShipMethKeytoID = Trim(.Fields("ShipMethID").value)
        End If
        .Filter = adFilterNone
    End With
End Function

'return WillCall ShipMethKey using User's BranchID

Public Function GetWillCallShipMethodKey() As Long
    Dim BranchID As String
    BranchID = GetBranchIDFromUserKey(GetUserKey)
    GetWillCallShipMethodKey = GetShipMethKeyFromBranchID(BranchID & "-Will Call")
End Function


Public Function GetShelfVendKeyFromWhseKey(WhseKey As Long) As Long
    With g_rstWhses
            .Filter = "WhseKey = " & WhseKey
            GetShelfVendKeyFromWhseKey = .Fields("WireShelfVendKey").value
            .Filter = adFilterNone
    End With
End Function


Public Function VendKeyToID(ByVal i_lVendKey As Long) As String
    With g_rstVendors
        .Filter = "VendKey=" & i_lVendKey
        If Not .EOF Then
            VendKeyToID = Trim(.Fields("VendID").value)
        End If
        .Filter = adFilterNone
    End With
End Function


