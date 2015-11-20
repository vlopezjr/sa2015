Attribute VB_Name = "Search"
Option Explicit

' Public Function FindCustomer(ByVal i_sText As String, ByVal i_lSearchType As Long, ByRef i_oCustomer As Customer, Optional bFindOnly As Boolean = False) As Long
'   Called By:
'       FWarehouse
'       FBilling
'       FUPS
'       FAcctRcv
'       FOrder

'   Who sets the optional flag?

'    Returns CustKey
'    A Customer object is passed in by Reference

' Public Function GetSearchType(ByVal i_sText As String) As Long
'   Called By:
'       FOrder.txtCustSearch_KeyUp
'       FAcctRcv.txtCustSearch_KeyUp



'These constants are used for the customer search dropdown list
'NOTE: The cboCustSearch list is defined at design time and must
'      *EXACTLY* match these constants
Private Const k_lCustSearchByID = 0
Private Const k_lCustSearchByName = 1
Private Const k_lCustSearchByZip = 2
Private Const k_lCustSearchByPhone = 3
'Private Const k_lCustSearchByVaxAcct = 4
Private Const k_lCustSearchByNatAcct = 4


' i_lSearchType has one of these constant values which are passed in as magic numbers or list indexes

Public Function FindCustomer(ByVal i_sText As String, ByVal i_lSearchType As Long, ByRef i_oCustomer As Customer, Optional bFindOnly As Boolean = False) As Long
    Dim lCustKey As Long
    Dim oFrm As FCustSearch

    'a search string is required (don't throw an error)
    If Len(i_sText) = 0 Then
        Exit Function
    End If

    ClearWaitCursor
    
    Set oFrm = New FCustSearch

    'Invoke the appropriate type of search
    Select Case i_lSearchType
        Case k_lCustSearchByID
            If Left(i_sText, 1) = "=" Then
                i_sText = Mid(i_sText, 2)
            End If
            lCustKey = oFrm.Find(k_sCustID, i_sText, i_oCustomer, bFindOnly)
            
        Case k_lCustSearchByName
            lCustKey = oFrm.Find(k_sCustNameOrID, i_sText, i_oCustomer, bFindOnly)
            
        Case k_lCustSearchByZip
            If Left(i_sText, 1) = "Z" Then
                i_sText = Mid(i_sText, 2)
            End If
            lCustKey = oFrm.Find(k_sCustZip, i_sText, i_oCustomer, bFindOnly)
            
        Case k_lCustSearchByPhone
            If Left(i_sText, 1) = "P" Then
                i_sText = Mid(i_sText, 2)
            End If
            lCustKey = oFrm.Find(k_sCustPhone, i_sText, i_oCustomer, bFindOnly)
            
'        Case k_lCustSearchByVaxAcct
'            If Left(i_sText, 1) = "Z" Or Left(i_sText, 1) = "P" Then
'                i_sText = Mid(i_sText, 2)
'            End If
'            lCustKey = oFrm.Find(k_sCustVaxAcct, i_sText, i_oCustomer, bFindOnly)
            
        Case k_lCustSearchByNatAcct
            If Left(i_sText, 1) = "=" Then
                i_sText = Mid(i_sText, 2)
            End If
            lCustKey = oFrm.Find(k_sNationalAccount, i_sText, i_oCustomer, bFindOnly)
        
    End Select
    
    FindCustomer = lCustKey
End Function


'Called By
'   FOrder.txtCustSearch_KeyUp
'   FAcctRcv.txtCustSearch_KeyUp

Public Function GetSearchType(ByVal i_sText As String) As Long
    If Len(i_sText) = 0 Then
        GetSearchType = k_lCustSearchByID
    ElseIf IsCustID(i_sText) Then
        GetSearchType = k_lCustSearchByID
    ElseIf IsPhone(i_sText) Then
        GetSearchType = k_lCustSearchByPhone
    ElseIf IsZip(i_sText) Then
        GetSearchType = k_lCustSearchByZip
'    ElseIf IsVaxAcct(i_sText) Then
'        GetSearchType = k_lCustSearchByVaxAcct
    Else
        GetSearchType = k_lCustSearchByID
    End If
End Function


Private Function IsCustID(ByVal i_sText As String) As Boolean
    If Left(i_sText, 1) = "=" Then
        IsCustID = True
    End If
End Function


Private Function IsPhone(ByVal i_sText As String) As Boolean
    If Left(i_sText, 1) = "P" Then 'we can assume our caller is forcing uppercase
        If Len(i_sText) = 1 Then
            IsPhone = True
        ElseIf IsNumeric(Mid(i_sText, 2)) Then
            IsPhone = True
        End If
    End If
End Function


Private Function IsZip(ByVal i_sText As String) As Boolean
    If Left(i_sText, 1) = "Z" Then 'we can assume our caller is forcing uppercase
        If Len(i_sText) = 1 Then
            IsZip = True
        ElseIf IsNumeric(Mid(i_sText, 2)) Then
            IsZip = True
        End If
    End If
End Function


'Private Function IsVaxAcct(ByVal i_sText As String) As Boolean
'    If IsNumeric(i_sText) Then
'        IsVaxAcct = True
'    End If
'End Function

'NOTE: be aware that this 'overrides' the VB intrinsic function

Private Function IsNumeric(ByVal i_sText As String) As Boolean
    Dim i As Long
    Dim chr As Long

    For i = 1 To Len(i_sText)
        chr = Asc(Mid(i_sText, i, 1))
        If chr < Asc("0") Or chr > Asc("9") Then
            IsNumeric = False
            Exit Function
        End If
    Next
    IsNumeric = True
End Function


