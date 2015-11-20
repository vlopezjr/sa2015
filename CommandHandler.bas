Attribute VB_Name = "CommandHandler"
Option Explicit


Public Sub ProcessToolbarCommand(ByVal Tool As ActiveToolBars.SSTool)
    Debug.Print "ToolID " & Tool.ID

    Select Case Tool.ID
    Case "ID_Finish":           DoFinish
    Case "ID_Save":             DoSave
    Case "ID_Cancel":           DoCancel
    Case "ID_Delete":           DoDelete
    Case "ID_NewWindow":        DoNewWindow
    Case "ID_SplitOrder":       DoSplitOrder
    Case "ID_Refresh":          MDIMain.DoRefresh
    Case "ID_Print":            DoPrint True
    Case "ID_PrintPreview":     DoPrint False
    Case "ID_Close":            DoClose
    Case "ID_Exit":             MDIMain.DoExit
    Case "ID_NextError":        DoNextError
    
    'this is a checkbox on the Development toolbar
    Case "ID_ConfirmExit":      g_bConfirmExit = (Tool.State = ssChecked)
    
    Case "ID_FilterCustSearch": g_bFilterCustSearch = (Tool.State = ssChecked)
    
    Case "ID_ATMMode":          g_bATMMode = (Tool.State = ssChecked)
                                MDIMain.DoRefreshATMMode
        
    Case "ID_IconsOnly":        MDIMain.UpdateToolBar (Tool.State = ssChecked)

    Case "ID_Cascade":          MDIMain.CascadeWindows

    Case "ID_mnuFile", "ID_mnuOptions", "ID_mnuWindowList", "ID_Management", "ID_mnuReports"
        'do nothing

    Case "ID_FOrderManager":            MsgBox "Not implemented in this release", vbInformation + vbOKOnly, "SageAssistant"
                                        'ShowTool "FOrderManager"
    Case "ID_FProvisionalShipments":    MsgBox "Not implemented in this release", vbInformation + vbOKOnly, "SageAssistant"
                                        'ShowTool "FProvisionalShipment"
    Case "ID_FPurchAssist":             ShowTool "FPurchAssist"
    Case "ID_StatusDashboard":          ShowTool "FStatusDashboard"
    Case "ID_FDocFinder":               ShowTool "FDocFinder"
    Case "ID_FCrossRef":                ShowTool "FCrossRef"
    Case "ID_FInvFinder":               ShowTool "FInvFinder"
    Case "ID_FPartsWiz":                ShowTool "FPartsWizTool"
    Case "ID_FUsers":                   MsgBox "Not implemented in this release", vbInformation + vbOKOnly, "SageAssistant"
                                        'ShowTool "FUsers"
    Case "ID_EditBins":                 ShowTool "FEditBins"
    Case "ID_FManagement":              ShowTool "FManagement"
    Case "ID_FPhoneFlagger":            ShowTool "FPhoneFlagger"
    Case "ID_FWarehouse":               ShowTool "FWarehouse"
    Case "ID_FPickConflicts":               ShowTool "FPickConflicts"
    'Case "ID_FPickConflictNotifications":   ShowTool "FPickConflictNotifications"
    Case "ID_FAcctRcv":                 ShowTool "FAcctRcv"
    Case "ID_FAcctPay":                 ShowTool "FAcctPay"
    Case "ID_FUPSAcct":                 ShowTool "FUPSAcct"
    Case "ID_FWillCallTool":            ShowTool "FWillCallTool"
    Case "ID_OnlineUsers":              DoOnlineUsers
    Case "ID_EndPoints":                ShowTool "FEndPoints"
    Case "ID_FPettyCashier":            ShowTool "FPettyCashier"
    Case "ID_FBilling":                 ShowTool "FBilling"

    Case "ID_FARCollections":   ShowTool "FARCollections"

    Case "ID_OnlineCatalog":    MsgBox "Not implemented in this release", vbInformation + vbOKOnly, "SageAssistant"
        
    Case "ID_Test"
        On Error Resume Next
        Dim frm As FTest
        
        Set frm = New FTest
        frm.Show
        On Error GoTo 0
        
    Case "ID_Bugs":             DoShowBugs
    
    Case Else
        'Reports
        If InStr(1, Tool.ID, "ID_Rpt") > 0 Then
            Dim oFrm As FViewer
            Set oFrm = New FViewer
            Call oFrm.ViewReportFromMenu(Mid$(Tool.ID, 7))
            Set oFrm = Nothing
        Else
            ' Window List Tools
            If MDIMain.ProcessingEvent Then
                Exit Sub
            Else
                MDIMain.ProcessingEvent = True
                If Tool.group = "WindowList" Then
                    DoSelectWindow Tool.ID
                ElseIf Tool.Category = "Help" Then
                    ShowHelp Mid$(Tool.ID, 4)
                Else
                    Debug.Print "No action defined for " & Tool.ID & " in category " & Tool.Category
                End If
                MDIMain.ProcessingEvent = False
            End If
        End If
    End Select
End Sub


Public Sub AutoStartTools()
    If HasRight(k_sRightAutoStartDashboard) Then
        ShowTool "FStatusDashboard"
    End If

    If HasRight(k_sRightAutoStartPartsWiz) Then
        ShowTool "FPartsWizTool"
    End If

    If HasRight(k_sRightAutoStartDocFinder) Then
        ShowTool "FDocFinder"
    End If

    If HasRight(k_sRightAutoStartInvFinder) Then
        ShowTool "FInvFinder"
    End If

    If HasRight(k_sRightShowToolOP) And HasRight(k_sRightAutoStartOP) Then
'!!! This is a special case that needs to be fixed (autostart FOrder)
        DoNewWindow
    End If

    If HasRight(k_sRightAutoStartAR) Then
        ShowTool "FAcctRcv"
    End If

    If HasRight(k_sRightAutoStartAP) Then
        ShowTool "FAcctPay"
    End If

    If HasRight(k_sRightAutoStartBilling) And HasBillingRight Then
        ShowTool "FBilling"
    End If

    If HasRight(k_sRightAutoStartPhoneFlagger) Then
        ShowTool "FPhoneFlagger"
    End If

    If HasRight(k_sRightAutoStartPurchasing) Then
        ShowTool "FPurchAssist"
    End If

End Sub


Private Sub ShowTool(ByVal i_sFrmName As String)
    Dim lWindowID As Long
    Dim oFrm As Form
    Dim i As Long

    'This enforces only one instance of all forms
    'If there is already one in the Forms collection, jump out
    'FOrder uses DoNewWIndow instead
    
    For i = 1 To Forms.Count - 1
        If Forms(i).Name = i_sFrmName Then
            Set oFrm = Forms(i)
            Exit For
        End If
    Next

    On Error GoTo ErrorHandler

    If oFrm Is Nothing Then
        Select Case i_sFrmName
            Case "FOrderManager":           Set oFrm = New FOrderManager
            Case "FStatusDashboard":        Set oFrm = New FStatusDashBoard
            Case "FProvisionalShipment":    Set oFrm = New FProvisionalShipment
            Case "FDocFinder":              Set oFrm = New FDocFinder
            Case "FInvFinder":              Set oFrm = New FInvFinder
            Case "FPartsWizTool":           Set oFrm = New FPartsWizTool
            Case "FAcctRcv":                Set oFrm = New FAcctRcv
            Case "FAcctPay":                Set oFrm = New FAcctPay
            Case "FUsers":                  Set oFrm = New FUsers
            Case "FPettyCashier":           Set oFrm = New FPettyCashier
            Case "FPurchAssist":        Set oFrm = New FPurchAssist
            Case "FManagement":         Set oFrm = New FManagement
            Case "FPhoneFlagger":       Set oFrm = New FPhoneFlagger
            Case "FBilling":            Set oFrm = New FBilling
            Case "FUPSAcct":            Set oFrm = New FUPSAcct
            Case "FCrossRef":           Set oFrm = New FCrossRef
            Case "FWarehouse":          Set oFrm = New FWarehouse
            Case "FPickConflicts":      Set oFrm = New FPickConflicts
            'Case "FPickConflictNotifications":   Set oFrm = New FPickConflictNotifications
            Case "FWillCallTool":       Set oFrm = New FWillCallTool
            Case "FARCollections":      Set oFrm = New FARCollections
            Case "FEndPoints":         Set oFrm = New FEndPoints
            
            Case Else
                Err.Raise -1, "MDIMain.ShowTool", "Need to update list of valid tools to include " & i_sFrmName
        End Select
        
        MDIMain.AddNewWindow oFrm
    End If

    With oFrm
        .Show
        .SetFocus
    End With
    
    Exit Sub
    
ErrorHandler:
    Err.Raise -1, "MDIMain.ShowTool", "Could not access WindowID for " & i_sFrmName
End Sub


Public Sub DoNewWindow()
    Dim frm As FOrder

    Set frm = New FOrder
    MDIMain.AddNewWindow frm
    frm.Show
    
    'This flushes a pending FOrder.Form_Paint event which if not flushed
    'will reactivate FOrder after all tools are displayed (see Init() above).
    DoEvents
End Sub


Public Function DoSplitOrder() As FOrder
    
    LogEvent "MDIMain", "DoSplitOrder", GetUserName & " instantiating FOrder from MDIMain.DoSplitOrder"

    'temporary reference variables
    Dim oCurrForm As FOrder
    Dim oCurrOrder As Order
    Dim oNewForm As FOrder

    Set oCurrForm = MDIMain.ActiveForm
    Set oCurrOrder = MDIMain.ActiveForm.Order
    Set oNewForm = New FOrder
    
    MDIMain.AddNewWindow oNewForm
    
    oNewForm.Show
          
    With oNewForm.Order
        .Create
        
        .Customer.Import oCurrOrder.Customer.Export

        'the new order will share the current order's customer contacts collection and selected contact object
        'overwrite the new Customer's Contacts reference (which is an unitialized collection object)
        .Customer.Contacts = oCurrOrder.Customer.Contacts
        .contact = oCurrOrder.contact
        
        CloneRemarks oCurrOrder.RemarkContext, .RemarkContext
        
        .PurchOrd = oCurrOrder.PurchOrd
        
        'changed this 2/26/15 LR
        'Set the current CSR as the split Order CSR.
        '.UserKey = GetUserKey
        
        'carry over the creating CSR
        'Note: the key also sets the id
        .UserKey = oCurrOrder.UserKey
        
        .WhseKey = oCurrOrder.WhseKey
        .ShipMethKey = oCurrOrder.ShipMethKey
        .SalesTax.Import oCurrOrder.SalesTax.Export
        
        'After splitting, the credit card information should also
        'be cloned to the new order.
            
        If oCurrOrder.PmtTerms.ID = "CrCard" Then
            'Note: the key also sets the id
            .PmtTerms.Key = oCurrOrder.PmtTerms.Key
            .CreditCard = oCurrOrder.CreditCard
        End If
        
    End With
    
    oNewForm.TransitionTabs False

'HACK: Ensure that these controls are consistent in the new form.
'This code should not be needed because the .TransitionTabs method should call
'an internal utlity function to set these variables to the proper state.
'Once that refactoring is done, the following code should be removed.

    With oNewForm
        .txtCustID.Visible = oCurrForm.txtCustID.Visible
        .lblCustID(0).Visible = oCurrForm.lblCustID(0).Visible
        .cboCustType.Visible = oCurrForm.cboCustType.Visible
        .lblCustType(0).Visible = oCurrForm.lblCustType(0).Visible
        .txtCustName.Visible = oCurrForm.txtCustName.Visible
        .lblCustName.Visible = oCurrForm.lblCustName.Visible
        .rvCustomer.Visible = oCurrForm.rvCustomer.Visible
        .txtInfo.Enabled = oCurrForm.txtInfo.Enabled
        .txtInfo.text = oCurrForm.txtInfo.text
        .cboWarehouse(0).ListIndex = oCurrForm.cboWarehouse(0).ListIndex
        .cboShipVia.ListIndex = oCurrForm.cboShipVia.ListIndex
        .chkDefaultShipMeth.Enabled = oCurrForm.chkDefaultShipMeth.Enabled
        .chkDefaultShipMeth.value = oCurrForm.chkDefaultShipMeth.value
        .chkBillRecipient.Enabled = oCurrForm.chkBillRecipient.Enabled
        .chkBillRecipient.value = oCurrForm.chkBillRecipient.value
        .txtUPSAcct.text = oCurrForm.txtUPSAcct.text
        .cmdUPSUpdate.Enabled = oCurrForm.cmdUPSUpdate.Enabled
        .chkPricePackList.value = oCurrForm.chkPricePackList.value
        '.chkDropShip.value = oCurrForm.chkDropShip.value
    End With
    
    'Log splitting events for both original and new orders
    LogOAEvent "Order", GetUserID, oCurrOrder.OPKey, , , "Split to new Order " & oNewForm.Order.OPKey
    LogOAEvent "Order", GetUserID, oNewForm.Order.OPKey, , , "This order was split from Order " & oCurrOrder.OPKey
   
    Set DoSplitOrder = oNewForm
End Function


Private Sub DoFinish()
    On Error Resume Next 'Ignore errors for unsupported method
    MDIMain.ActiveForm.CommitButton
    ForceRefresh MDIMain.ActiveForm.hwnd
End Sub


Private Sub DoCancel()
    On Error Resume Next 'Ignore errors for unsupported method
    MDIMain.ActiveForm.CancelButton
    ForceRefresh MDIMain.ActiveForm.hwnd
End Sub


Private Sub DoNextError()
    On Error Resume Next 'Ignore errors for unsupported method
    MDIMain.ActiveForm.BrokenRules.SetFocusNext
End Sub


Private Sub DoDelete()
    On Error Resume Next
    MDIMain.ActiveForm.DeleteButton
End Sub


Private Sub DoSave()
    On Error Resume Next 'Ignore errors for unsupported method
    MDIMain.ActiveForm.SaveButton
End Sub

Private Sub DoClose()
    If Not MDIMain.ActiveForm Is Nothing Then
        Unload MDIMain.ActiveForm
    End If
End Sub


Private Sub DoPrint(i_bPrintOnly As Boolean)
    On Error Resume Next 'Ignore errors for unsupported method
    MDIMain.ActiveForm.PrintButton i_bPrintOnly
End Sub


Private Sub DoOnlineUsers()
'    Dim frmMgt As New FMgt
'    frmMgt.Show
End Sub


Private Sub DoSelectWindow(sWindowID As String)
    Dim frm As Form
    Dim lWindowID As Long
    Dim lFoundID As Long

    On Error Resume Next 'this suppresses error if WindowID not defined
    lWindowID = CLng(Mid$(sWindowID, 7))
    For Each frm In Forms
        lFoundID = frm.WindowID
        If lFoundID = lWindowID Then
            With frm
                .Show
                If .WindowState = 1 Then .WindowState = 0
                .SetFocus
            End With
            MDIMain.UpdateToolbarStatus
            Exit Sub
        End If
    Next
End Sub


Public Sub DoShowBugs()
    Dim sFile As String

    With App
        sFile = .Major & "-" & .Minor & "-" & .Revision
    End With
    ShowHelp "ReleaseNotes/" & sFile

    'update registry to indicate that user has read release notes
'    PutRegNumberValue HKEY_CURRENT_USER, g_RegKeyOP, "AppMajor", App.Major
'    PutRegNumberValue HKEY_CURRENT_USER, g_RegKeyOP, "AppMinor", App.Minor
'    PutRegNumberValue HKEY_CURRENT_USER, g_RegKeyOP, "AppRevision", App.Revision
    g_UserConfig.SetKeyValue "officeassistant", "appMajor", App.Major
    g_UserConfig.SetKeyValue "officeassistant", "appMinor", App.Minor
    g_UserConfig.SetKeyValue "officeassistant", "appRevision", App.Revision
    g_UserConfig.Save
End Sub


