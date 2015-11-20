VERSION 5.00
Object = "{17259422-7BB1-421A-88DE-7EA81AC1AC6F}#4.0#0"; "IGToolBars40.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm MDIMain 
   BackColor       =   &H8000000C&
   Caption         =   "<appname>"
   ClientHeight    =   7740
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   9765
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   300
      Top             =   3300
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   300
      Top             =   2526
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483643
      ImageWidth      =   60
      ImageHeight     =   25
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   300
      Top             =   1754
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":2532
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ActiveToolBars.SSActiveToolBars Toolbar1 
      Left            =   300
      Top             =   300
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262144
      ToolBarsCount   =   6
      ToolsCount      =   49
      ActiveColors    =   -1  'True
      ShowShortcutsInToolTips=   -1  'True
      Tools           =   "MDIMain.frx":2644
      ToolBars        =   "MDIMain.frx":1A565
   End
   Begin MSComDlg.CommonDialog dlgXML 
      Left            =   300
      Top             =   1042
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   ".xml"
      Filter          =   "XML Files|*.xml"
   End
End
Attribute VB_Name = "MDIMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

'Public Methods:
'Public Sub Init()
'Public Sub GlobalKeyDownProcessing(KeyCode As Integer, Shift As Integer, Optional oRules As BrokenRules = Nothing)
'Public Sub UpdateToolBar(bCheck As Boolean)
'Public Sub DoRefreshATMMode()
'Public Sub FormUnregister(ByRef i_oFrm As Form)
'Public Sub UnloadTool(lWindowID As Long)
'Public Sub UpdateCaption(frm As Form)
'Public Sub UpdateWindowListSelection(frm As Form)
'Public Sub CascadeWindows()
'Public Sub DoRefresh()
'Public Sub AddNewWindow(ByVal frm As Form)
'Public Sub DoExit()
'Public Sub UpdateToolbarStatus()


Public ProcessingEvent As Boolean

Private m_bLoading As Boolean

Private m_lWindowID As Long

'************************************************************************************
' Form Events
'************************************************************************************

' This event handler runs before the Init method (below).

Private Sub MDIForm_Load()
    Dim i As Long
    Dim sToolbarConfig As String
    Dim lToolbarVersion As Long
    Dim liCounter As Integer
    Dim lsRptName As String
    Dim lorst As ADODB.Recordset

    Me.caption = GlobalFunctions.Version(True)

    'lToolbarVersion = GetRegNumberValue(HKEY_CURRENT_USER, g_RegKeyOP, "ToolbarVersion", -1)
    
    lToolbarVersion = g_UserConfig.GetKeyValue("toolbars", "toolbarRev")
    
    If lToolbarVersion = g_ToolbarVersion Then
        On Error Resume Next 'ignore error if file doesn't exist
        Toolbar1.LoadConfiguration ToolbarFilePath
        On Error GoTo 0
    Else
        msg "The Toolbar has been updated. Resetting. Your customizations will be lost."
    End If
   
    'Add menu and button tools to Menu toolbar
    'Note: Each new sub menu is added to top of the list,
    '      therefore, the rst must be loaded in desc order.
    Set lorst = LoadDiscRst("select RptName from tcpRpt Where RptShow = 1 order by RptSeqNo desc")
    If Not lorst.EOF Then
        With Toolbar1
            'Reports toolbar tool may not exist
            On Error Resume Next
            .Tools.Remove "ID_mnuRpts"
            On Error GoTo 0
            
            .Tools.Add "ID_mnuRpts", ssTypeMenu
            .Tools("ID_mnuRpts").Name = "Reports"
            .Toolbars("Menu Bar").Tools.Add "ID_mnuRpts", , 4
    
            For liCounter = 1 To lorst.RecordCount
                lsRptName = "ID_Rpt" & Trim(lorst.Fields("RptName"))
                
                'Button Tools - Tool ID may not exist
                On Error Resume Next
                .Tools.Remove lsRptName
                On Error GoTo 0
                
                'Debug.Print lsRptName & "|" & Trim(lorst.Fields("RptName"))
                .Tools.Add lsRptName
                .Tools(lsRptName).Name = Trim(lorst.Fields("RptName"))
                .Tools(lsRptName).DisplayStyle = ssDisplayTextOnlyAlways
                .Tools("ID_mnuRpts").Menu.Tools.Add lsRptName, , 1
                lorst.MoveNext
            Next
        End With
    End If

    g_bConfirmExit = (Toolbar1.Tools("ID_ConfirmExit").State = ssChecked)
    g_bFilterCustSearch = (Toolbar1.Tools("ID_FilterCustSearch").State = ssChecked)
    g_bATMMode = (Toolbar1.Tools("ID_ATMMode").State = ssChecked)
    
    ImplementSecurityRestrictions
End Sub


Private Sub MDIForm_Activate()
    Debug.Print "MDIMain_Activate"
    UpdateWindowListSelection ActiveForm
End Sub


Private Sub MDIForm_Unload(Cancel As Integer)
    Dim i As Long
    Dim s As String
    Dim PID As Long
    Dim SID As Long
    
    Debug.Print "Entered MDIForm_Unload"

    User.LogOffUser

    'close global recordsets
    
    Set g_rstRMAReason = Nothing
    Set g_rstRMADisposition = Nothing
    Set g_rstWhses = Nothing
    Set g_rstStates = Nothing
    Set g_rstShipVia = Nothing
    Set g_rstMakes = Nothing
    Set g_rstVendors = Nothing
    Set g_rstCustTypes = Nothing
    Set g_rstUsers = Nothing
    Set g_rstCSRs = Nothing
    Set g_rstCollectors = Nothing
    Set g_rstCountry = Nothing
    'Set g_rstPrinters = Nothing
    Set g_rstGaskMats = Nothing
    Set g_rstWarmerWire = Nothing
    Set g_rstPartIDs = Nothing

    'ensure that we save no spurious window tools
    On Error Resume Next
    For i = 1 To m_lWindowID
        Me.UnloadTool i
    Next
    
    Toolbar1.SaveConfiguration ToolbarFilePath
    'PutRegNumberValue HKEY_CURRENT_USER, g_RegKeyOP, "ToolbarVersion", g_ToolbarVersion
    g_UserConfig.SetKeyValue "toolbars", "toolbarRev", g_ToolbarVersion
    g_UserConfig.Save

    'just in case any forms remain open
    For i = Forms.Count - 1 To 0 Step -1
        Debug.Print Forms(i).Name
        s = s & vbCrLf & Forms(i).Name
        Unload Forms(i)
    Next

    LogEvent "MDIMain", "Unload", "Exited SageAssistant: " & GlobalFunctions.GetProcessInfo

    g_DB.Disconnect

End Sub


'************************************************************************************
' Public Methods
'************************************************************************************

Public Sub Init()

    'Disable Exit button on tools bar during init.
    'This will prevent user from trying to exit while MDIMain is still loading tools.
    'This fixes the error "Unable to unload within this context"
    
    Toolbar1.Tools("ID_Exit").Enabled = False
    
    CommandHandler.AutoStartTools
        
    Toolbar1.Tools("ID_Exit").Enabled = True
End Sub


'!!! Do I need to pass in BrokenRules?

Public Sub GlobalKeyDownProcessing(KeyCode As Integer, Shift As Integer, Optional oRules As BrokenRules = Nothing)

    If KeyCode = vbKeyR And Shift = (vbCtrlMask + vbAltMask) Then
        If IsAdmin(GetUserName) Then
            Dim mm As MemoMeister.RemarkContext
            
            Set mm = New MemoMeister.RemarkContext
            mm.ViewManager (1)
        End If

        'Reload internal rights cache
        RefreshRights
        DoRefresh
        ClearWaitCursor
        KeyCode = 0
        
    ElseIf KeyCode = vbKeyTab And Shift = (vbCtrlMask + vbAltMask) Then
        If Not oRules Is Nothing Then
            oRules.SetFocusNext
            KeyCode = 0
        End If
        
    ElseIf KeyCode = vbKeyF1 Then
        On Error Resume Next
        Me.ActiveForm.DoShowHelp
    End If
End Sub


'Called by
'   ToolBar1_ToolClick()
'
'Supports Icon Only mode
'NOTE: the parameter is poorly named and typed

Public Sub UpdateToolBar(bCheck As Boolean)
    Dim oTool As ActiveToolBars.SSTool
    Dim oToolBar As ActiveToolBars.SSToolBar
    Dim tempStyle As Long
    
    'For Each oTool In Toolbar1.ToolBars("Main").Tools
    
    ProcessingEvent = True
    If bCheck Then
        tempStyle = ssDisplayDefaultStyle
    Else
        tempStyle = ssDisplayImageAndText
    End If
    
    For Each oToolBar In Toolbar1.Toolbars
        For Each oTool In oToolBar.Tools
            Select Case oTool.ID
                Case "ID_NextError":    oTool.DisplayStyle = tempStyle
                Case "ID_Finish":       oTool.DisplayStyle = tempStyle
                Case "ID_Save":         oTool.DisplayStyle = tempStyle
                Case "ID_Cancel":       oTool.DisplayStyle = tempStyle
                Case "ID_NewWindow":
                    oTool.DisplayStyle = tempStyle
                Case "ID_SplitOrder":   oTool.DisplayStyle = tempStyle
                Case "ID_Refresh":      oTool.DisplayStyle = tempStyle
                Case "ID_Print":        oTool.DisplayStyle = tempStyle
                Case "ID_PrintPreview": oTool.DisplayStyle = tempStyle
                Case "ID_Exit":         oTool.DisplayStyle = tempStyle
                Case "ID_Delete":       oTool.DisplayStyle = tempStyle
            End Select
        Next
    Next
    ProcessingEvent = False
End Sub


Public Sub DumpToolbars()
    Dim oTool As ActiveToolBars.SSTool
    Dim oTools As ActiveToolBars.SSTools
    Dim oToolBar As ActiveToolBars.SSToolBar
    Dim oMenuTool As ActiveToolBars.SSTool
    
    For Each oToolBar In Toolbar1.Toolbars
        Debug.Print oToolBar.Name
        For Each oTool In oToolBar.Tools
            Debug.Print oTool.ID
            If oTool.Type = ssTypeMenu Then
                For Each oMenuTool In oTool.Menu.Tools
                    Debug.Print ">" & oMenuTool.ID
                Next
            End If
        Next
    Next
'    For Each oTool In Toolbar1.Tools
'        Debug.Print oTool.ID & vbTab & oTool.Name
'    Next
    
End Sub


Public Sub DoRefreshATMMode()
    On Error Resume Next
    Me.ActiveForm.txtPrice.ATMMode = g_bATMMode
    Me.ActiveForm.txtCost.ATMMode = g_bATMMode
    Me.ActiveForm.txtCreditLimit.ATMMode = g_bATMMode
    Me.ActiveForm.txtStdUC.ATMMode = g_bATMMode
    Me.ActiveForm.txtRplUC.ATMMode = g_bATMMode
End Sub


'NOTE: Aren't FormUnregister & UnloadTool essentially identical?
' FormUnregister overloads UnloadTool (a different parameter for the same function)
' They both simply remove the form from the Toolbar TOols collection.

Public Sub FormUnregister(ByRef i_oFrm As Form)
    Dim lWindowID As Long
    On Error GoTo IgnoreError
    lWindowID = i_oFrm.WindowID
    UnloadTool lWindowID
    Exit Sub

IgnoreError:
    'do nothing
End Sub


Public Sub UnloadTool(lWindowID As Long)
    On Error Resume Next
    Me.Toolbar1.Tools.Remove "Window" & lWindowID
End Sub


Public Sub UpdateCaption(frm As Form)
    Dim sWindowID As String
    
    On Error Resume Next 'this suppresses error if WindowID not defined
    sWindowID = "Window" & frm.WindowID
    Toolbar1.Tools("ID_mnuWindowList").Menu.Tools(sWindowID).Name = frm.caption
End Sub


Public Sub UpdateWindowListSelection(frm As Form)
    Dim sWindowID As String

    If ProcessingEvent Then Exit Sub
    
    On Error GoTo ErrorHandler
    ProcessingEvent = True
    sWindowID = "Window" & frm.WindowID
    Toolbar1.Tools("ID_mnuWindowList").Menu.Tools(sWindowID).State = ssChecked
    UpdateToolbarStatus
    ProcessingEvent = False
    Exit Sub

ErrorHandler:
    ProcessingEvent = False
    Debug.Print "Ignore error in UpdateWindowListSelection"
End Sub


Public Sub CascadeWindows()
    Me.Arrange vbCascade
End Sub


Public Sub DoRefresh()
    ForceRefresh Me.hwnd
    UpdateToolbarStatus
End Sub


Public Sub AddNewWindow(ByVal frm As Form)
    Dim sWindowID As String
    
    On Error GoTo ErrorHandler
    
    m_lWindowID = m_lWindowID + 1
    frm.WindowID = m_lWindowID
    sWindowID = "Window" & m_lWindowID
    
    ' Creates a new State Button Tool by adding it to the
    ' control-level Tools collection
    With Toolbar1
        .Tools.Add sWindowID, ssTypeStateButton
    
        ' Set properties of the new State Button Tool
        With .Tools(sWindowID)
'            .Name = frm.Caption
            .Name = sWindowID
            .group = "WindowList"
            .GroupAllowAllUp = False
            .State = ssChecked
            .PictureDown = ImageList1.ListImages(1).Picture 'assign checkmark
        End With

        ' Adds a copy of the new Tool as the last Tool on the
        ' Window List Menu Tool

        With .Tools("ID_mnuWindowList").Menu.Tools
            .Add sWindowID, , .Count + 1

        ' If there are only 2 Tools, add a Separator Tool
            If .Count = 2 Then
                .Add "separator", ssTypeSeparator, 2
            End If
        End With
    End With
    Exit Sub
    
ErrorHandler:
    ErrorUI.FatalError _
        "MDIMain:AddNewWindow", _
        "Cannot define WindowID '" & sWindowID & "' for form '" & frm.caption & "'"
End Sub


Public Sub DoExit()

    On Error Resume Next
    Dim oFrm As Form
    For Each oFrm In Forms
        If oFrm.Name = "FOrder" Then
            If Not oFrm.ExitCheck Then
                oFrm.SetFocus
                UpdateToolbarStatus
                Exit Sub        'an option to cancel out of Exit
            End If
        End If
    Next

    g_bExitNow = True

    Unload Me

End Sub


'!!! required for the Split Order button to be enabled

'Can you ever get in here with any of these conditions true?

'    If Not HasRight(k_sRightShowToolOP) Then
'        msg "Sorry, you do not have permission to split orders.", vbOKOnly + vbExclamation, "Cannot Split Order"
'        Exit Function
'    End If
'
'    If Me.ActiveForm Is Nothing Then
'        msg "You must select an Order form before you can split an order", , "Select Order to Split"
'        Exit Function
'    End If
'
'    If Me.ActiveForm.Name <> "FOrder" Then
'        msg "You must select an Order form before you can split an order", , "Select Order to Split"
'        Exit Function
'    End If


Public Sub UpdateToolbarStatus()
    Dim bEnabled As Boolean
    
    On Error Resume Next
    
    With Me.ActiveForm
        'This next line is checking to see if there are any dynamic windows in the window list
        Toolbar1.Tools("ID_Close").Enabled = (Toolbar1.Tools("ID_mnuWindowList").Menu.Tools.Count > 2)
        
        bEnabled = False    'Set to false in case the next line has error
        bEnabled = .CancelButton(False)
        Toolbar1.Tools("ID_Cancel").Enabled = bEnabled
        
        If Me.ActiveForm.Name = "FOrder" Then
            'Change the Cancel button dynamically
            If Me.ActiveForm.FindMode Then
                UpdateTool "ID_Cancel", "Reload"
                Toolbar1.Tools("ID_Cancel").ToolTipText = "Reload last order"
                UpdateWillCallToolBar
            Else
                UpdateTool "ID_Cancel", "Close"
                Toolbar1.Tools("ID_Cancel").ToolTipText = "Close this order"
               
                UpdateWillCallToolBar
                If Me.ActiveForm.RecommitSage And Not g_bWillCallUser Then
                    UpdateTool "ID_Finish", "Recommit"
                End If


            End If
        Else
            UpdateTool "ID_Cancel", "Close"
            Toolbar1.Tools("ID_Cancel").ToolTipText = "Close this order"
        End If
        
        bEnabled = False    'Set to false in case the next line has error
        bEnabled = .DeleteButton(False)
        Toolbar1.Tools("ID_Delete").Enabled = bEnabled
    
        bEnabled = False    'Set to false in case the next line has error
        bEnabled = .CommitButton(False)
        Toolbar1.Tools("ID_Finish").Enabled = bEnabled

        bEnabled = False    'Set to false in case the next line has error
        bEnabled = .SaveButton(False)
        Toolbar1.Tools("ID_Save").Enabled = bEnabled
        
        bEnabled = False    'Set to false in case the next line has error
        'If Not g_bWillCallUser Then
        bEnabled = .SplitOrderButton(False)
        'End If
        Toolbar1.Tools("ID_SplitOrder").Enabled = bEnabled
    
        bEnabled = False    'Set to false in case the next line has error
        bEnabled = .PrintButton(False, False)
'3/17/05 LR disable the Print function for now. Allow only Print Preview
        Toolbar1.Tools("ID_Print").Enabled = False
        Toolbar1.Tools("ID_PrintPreview").Enabled = bEnabled
        If g_bWillCallUser Then
            UpdateTool "ID_PrintPreview", "Receipt"
        Else
            UpdateTool "ID_PrintPreview", "Print Preview"
        End If

        bEnabled = False    'Set to false in case the next line has error
        bEnabled = .BrokenRules.Count > 0
        Toolbar1.Tools("ID_NextError").Enabled = bEnabled

        Toolbar1.Tools("ID_Refresh").Enabled = (.Name = "FOrder")
    End With
End Sub


'************************************************************************************
' Private functions
'************************************************************************************

Private Function ToolbarFilePath() As String
    ToolbarFilePath = g_ToolbarConfigPath & GetUserName & ".atb"
End Function


Private Sub ToolBar1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    CommandHandler.ProcessToolbarCommand Tool
End Sub


Private Sub ImplementSecurityRestrictions()
    On Error Resume Next
    Dim bFlag As Boolean
    Dim sRightChoice As String
    
    bFlag = HasRight(k_sRightShowToolDev)
    With Toolbar1.Toolbars.Item("Development")
        .Visible = bFlag
        .AllowCustomize = bFlag
        .AllowHiding = bFlag
    End With
      
    With Toolbar1.Tools
        'this is a temporary override
        .Item("ID_FProvisionalShipments").Visible = False
    
        .Item("ID_FAcctRcv").Visible = HasAnARRight
        '.Item("ID_FAcctPay").Visible = HasRight(k_sRightAPManageCost)
        .Item("ID_FAcctPay").Visible = HasRight(k_sRightShowToolAP)
        .Item("ID_NewWindow").Enabled = HasRight(k_sRightShowToolOP)
        .Item("ID_StatusDashboard").Enabled = HasRight(k_sRightShowToolDashboard)
        
        .Item("ID_FPurchAssist").Visible = HasRight(k_sRightShowToolPurch)
        .Item("ID_EditBins").Visible = HasRight(k_sRightShowToolBins)
        '.Item("ID_FVaxAcct").Visible = HasRight(k_sRightShowToolVaxAcct)
        .Item("ID_FUPSAcct").Visible = HasRight(k_sRightShowToolUPSAcct)
        .Item("ID_FCrossRef").Visible = HasRight(k_sRightShowToolCrossRef)
        .Item("ID_FManagement").Visible = HasRight(k_sRightShowToolManagement)
        .Item("ID_FPhoneFlagger").Visible = HasRight(k_sRightShowToolPhoneFlagger)
        .Item("ID_FBilling").Visible = HasBillingRight
        .Item("ID_Warehouse").Visible = True
        
        .Item("ID_FARCollections").Visible = HasRight(k_sRightShowToolARCollections)
        .Item("ID_FWillCallTool").Visible = HasRight(k_sRightShowToolWillCall)
    End With
    
End Sub


'Public Sub LoadCustInNewOrderPad()
'    CommandHandler.DoNewWindow
'    DoEvents
'    Me.ActiveForm.txtCustName.text = Toolbar1.Tools("ID_LoadText").Edit.text
'    Me.ActiveForm.BrokenRules.Validate
'End Sub


'Called by
'   UpdateToolbarStatus()

Private Sub UpdateWillCallToolBar()
'removed 5/7/15 LR
'    If g_bWillCallUser Then
'            UpdateTool "ID_Save", "Pay"
'            UpdateTool "ID_Finish", "Sign"
'            Toolbar1.Tools("ID_Finish").ToolTipText = "Customer is signing an account"
'            Toolbar1.Tools("ID_Save").ToolTipText = "Customer is paying now"
'    Else
            UpdateTool "ID_Finish", "Commit"
            UpdateTool "ID_Save", "Save"
            Toolbar1.Tools("ID_Finish").ToolTipText = "Commit this order to Sage"
            Toolbar1.Tools("ID_Save").ToolTipText = "Save this order in SageAssistant"
'    End If
End Sub


'Called by
'   UpdateToolbarStatus()
'   UpdateWillCallToolBar()
'
'This is used to change the label on a tool button at run time.
'Used especially to support WillCall.

Private Sub UpdateTool(sToolID As String, sToolName As String)
    Dim oTool As ActiveToolBars.SSTool
    Dim oToolBar As ActiveToolBars.SSToolBar
    Dim tempStyle As Long
    
    For Each oToolBar In MDIMain.Toolbar1.Toolbars
        For Each oTool In oToolBar.Tools
            Select Case oTool.ID
                Case sToolID: oTool.Name = sToolName
            End Select
        Next
    Next
    Toolbar1.Tools("ID_mnuFile").Menu.Tools.Item(sToolID).Name = sToolName
End Sub


Private Sub AddNewOption()
    With MDIMain.Toolbar1
        .Tools.Add "ChooseIcon", ssTypeStateButton
        With .Tools("ChooseIcon")
            .Name = "ChooseIcon"
            .group = "Development"
            .GroupAllowAllUp = False
            .State = ssUnchecked
            .PictureDown = ImageList1.ListImages(1).Picture
        End With
        
        With .Tools("ID_munOptions").Menu.Tools
            .Add "ChooseIcon", 4
        End With
    End With
End Sub


Private Sub Timer1_Timer()
    User.LogOnUser     'let's refresh the database periodically
            
    Timer1.Enabled = True
End Sub

