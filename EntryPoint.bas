Attribute VB_Name = "EntryPoint"
Option Explicit

'**********************************************************************************
' The project's Entry-point
'**********************************************************************************

Public Sub Main()
    Dim sSQL As String
    Dim TargetServer As String
    Dim UserConfigPath As String
    Dim PID As Long
    Dim sID As Long
        
    'want to classify user without going to the database (for MM group membership)
    'could go to AD, but don't know how
    'so I'll hardcode this for now
    If IsAdmin(GetUserName) Then
    
        'where do I find the user config file?
        'the first place to look in the App folder
        'if it's not there, look on the DEV web server by default
        Set g_UserConfig = New Configuration
        g_UserConfig.TargetServer = Registry.GetRegStringValue(Registry.HKEY_LOCAL_MACHINE, "SOFTWARE\Wow6432Node\Caseparts", "ConfigServer", "WEB15DEV")
        
        g_UserConfig.Load GetUserName
        
        'bring up the logon to determine where the global config file is and load it
        Set g_GlobalConfig = New Configuration
        
        If Not FLogon.Logon(g_UserConfig) Then Exit Sub
        
        g_GlobalConfig.TargetServer = FLogon.TargetServer
        g_GlobalConfig.Load "config"
        
    End If
    
    If Not IsAdmin(GetUserName) Then
        TargetServer = Registry.GetRegStringValue(Registry.HKEY_LOCAL_MACHINE, "SOFTWARE\Wow6432Node\Caseparts", "ConfigServer", "WEB15")
        Set g_UserConfig = New Configuration
        g_UserConfig.TargetServer = TargetServer
        g_UserConfig.Load GetUserName
        
        Set g_GlobalConfig = New Configuration
        g_GlobalConfig.TargetServer = TargetServer
        g_GlobalConfig.Load "config"
    End If
    
    'equivalent to the old ReadIni
    GetGlobalConfigValues
    
    Set g_DB = New DBConnect
    g_DB.GlobalConfig = g_GlobalConfig
    g_DB.userconfig = g_UserConfig
        
    On Error GoTo ErrorHandler

    'this has been unexamined since day 1
    ErrorUI.LogLevel = elWarning
    ErrorUI.ForceCallTrace = False
    
    g_bConfirmExit = True
    g_bFilterCustSearch = True
        
    With FSplash
        .Show
        
        .Progress 10, "Connecting to Sage database"
        
        If Not g_DB.Connect Then
            End
        End If
        
        'wait until here because in order to log an error, need database connection established first
        
        LogEvent "EntryPoint", "Main", "Started SageAssistant: " & GlobalFunctions.GetProcessInfo

        GlobalFunctions.KillExistingInstance "SageAssistant.exe"
        
        .Progress 15, "Loading RMA ReasonCode"
        sSQL = "SELECT * FROM tcpRMAReason WHERE deprecate=0 ORDER BY RMAReasonID"
        Set g_rstRMAReason = LoadDiscRst(sSQL)
        
        .Progress 18, "Loading RMA DispositionCode"
        sSQL = "SELECT * FROM tcpRMADisposition "
        Set g_rstRMADisposition = LoadDiscRst(sSQL)
    
'WHSE
        .Progress 30, "Loading Warehouses"
        sSQL = "SELECT timWarehouse.WhseKey, RTRIM(timWarehouse.WhseID) as WhseID, timWarehouse.ShipAddrKey, tcpBranch.BranchName, " _
            & "timWarehouse.Description, timWarehouse.Transit, timWarehouse.SalesAcctKey, " _
            & "tcpBranch.ShipMethKey , tcpBranch.SalesTerritoryKey, tcpBranch.WireShelfVendKey " _
            & "FROM timWarehouse LEFT OUTER JOIN tcpBranch ON timWarehouse.WhseKey = tcpBranch.WhseKey " _
            & "WHERE companyid = 'cpc' ORDER BY WhseID"
        Set g_rstWhses = LoadDiscRst(sSQL)

        .Progress 35, "Loading state IDs"
        sSQL = "SELECT StateID, CountryID " _
             & "FROM tsmState " _
             & "WHERE StateID <> '<>'" _
             & "ORDER BY StateID"
        Set g_rstStates = LoadDiscRst(sSQL)
             
        .Progress 40, "Loading shipping methods"
        sSQL = "SELECT ShipMethKey, ShipMethID " _
             & "FROM tciShipMethod " _
             & "WHERE CompanyID = 'CPC' " _
             & "ORDER BY ShipMethID"
        Set g_rstShipVia = LoadDiscRst(sSQL)
        
        .Progress 50, "Loading makes"
        sSQL = "SELECT tcpNewMake.* " _
             & "FROM tcpNewMake " _
             & "WHERE isobsolete = 0 " _
             & "ORDER BY tcpNewMake.MakeText"
        Set g_rstMakes = LoadDiscRst(sSQL)

        .Progress 55, "Loading vendors"
        
        'We're assuming that vendors with a DfltPurchAcctKey = 3088 (GLAcctNo = 435000)
        'are part vendors.
        sSQL = "SELECT VendKey, DfltPurchAcctKey, VendID, VendName " _
             & "FROM tapVendor " _
             & "WHERE (CompanyID = 'CPC') AND DfltPurchAcctKey = 3088 AND Status = 1 " _
             & "ORDER BY VendName"
        Set g_rstVendors = LoadDiscRst(sSQL)
        
        .Progress 60, "Loading customer class IDs"
        sSQL = "SELECT CustClassID, CustClassKey " _
             & "FROM tarCustClass " _
             & "WHERE CompanyID = 'CPC' " _
             & "ORDER BY CustClassID"
        Set g_rstCustTypes = LoadDiscRst(sSQL)
 
        .Progress 65, "Logging on NT user"
        .Progress 66, "Loading CSR Users"
        .Progress 68, "Loading Buyers"
        .Progress 69, "Loading Collectors"
        User.Init

        .Progress 70, "Loading Countries"
        sSQL = "SELECT CountryID, PhoneMask, PostalCodeMask FROM tsmCountry ORDER BY CountryID"
        Set g_rstCountry = LoadDiscRst(sSQL)
        
'        .Progress 72, "Loading Printers"
'        sSQL = "SELECT * FROM tcpPrinter "
'        Set g_rstPrinters = LoadDiscRst(sSQL)
               
        .Progress 80, "Loading gasket materials"
        Set g_rstGaskMats = LoadDiscRst("spOPGsktMatLoad")

        .Progress 82, "Loading Warmer Wire"
        sSQL = "SELECT * FROM tcpWarmerWire "
        Set g_rstWarmerWire = LoadDiscRst(sSQL)
        
        .Progress 85, "Loading special part IDs"
        sSQL = "SELECT ItemKey, ItemID " _
             & "FROM timItem " _
             & "WHERE CompanyID = 'CPC' " _
             & "AND (ItemID LIKE 'SPO-%'" _
             & "  OR ItemID LIKE 'GSK-%'" _
             & "  OR ItemID LIKE 'WWR-%'" _
             & "  OR ItemID LIKE 'SHF-%')"
        Set g_rstPartIDs = LoadDiscRst(sSQL)

        .Progress 90, "Loading main form"
        Load MDIMain

        .Progress 95, "Loading order form"
        MDIMain.Init
        
        .Progress 100, "Displaying forms"
        Unload FSplash
        MDIMain.Show
        
    End With

    Set g_OrderCommitProxy = New MSSOAPLib30.SoapClient30
    g_OrderCommitProxy.MSSoapInit g_AutoCommitUrl & "?WSDL"
    
    CheckVersion
     
    Exit Sub

ErrorHandler:
'    ErrorUI.FatalError "EntryPoint.Main", _
'                                  "Office Assistant initialization failed." & vbCrLf _
'                                & "Please try again.  If the problem" & vbCrLf _
'                                & "persists, contact the IT department immediately."
    MsgBox Err.Number & " " & Err.Description, vbOKOnly, "System Error"
    End
End Sub


'6/17/15: this is all changing

' Read the Name/Value pairs in \\WEB\OfficeAssistant\ConfigFiles\Config.ini into global variables.
' Values are conditioned on the chosen environment.
' The database connection is also conditioned on the chosen environment.
' TODO: Can the location of the config.ini file be conditioned on the environment as well?

Private Sub GetUserConfigValues()
    With g_UserConfig
    
    End With
End Sub

'the config.xml file we're reading is coming from the server pointed to by a developer's xml file
'so the target environment is already established before the xml file is read
    
Private Sub GetGlobalConfigValues()
        
    With g_GlobalConfig
    
        g_MPKWhseKey = CLng(.GetKeyValue("constants", "MPKWhseKey"))
        g_SEAWhseKey = CLng(.GetKeyValue("constants", "SEAWhseKey"))
        g_STLWhseKey = CLng(.GetKeyValue("constants", "STLWhseKey"))
        g_STLWhseID = .GetKeyValue("constants", "STLWhseID")
        g_SEAWhseID = .GetKeyValue("constants", "SEAWhseID")
        g_MPKWhseID = .GetKeyValue("constants", "MPKWhseID")
        
        g_CreatePOTempTables = .GetKeyValue("sageSqlScripts", "createPOTempTables")
        g_DropPOTempTables = .GetKeyValue("sageSqlScripts", "dropPOTempTables")
        g_CreateIMTempTables = .GetKeyValue("sageSqlScripts", "createIMTempTables")
        g_DropIMTempTables = .GetKeyValue("sageSqlScripts", "dropIMTempTables")
        
        g_SnapshotPath = .GetKeyValue("paths", "snapshot")
        g_XsltPath = .GetKeyValue("paths", "xslt")
        g_QuoteEmailPath = .GetKeyValue("paths", "quoteEmail")
        g_ReportPath = .GetKeyValue("paths", "reports")
        g_ToolbarConfigPath = .GetKeyValue("paths", "toolbarConfig")
        g_CollectionsPath = .GetKeyValue("paths", "collections")
        
        g_OrderPadBaseURL = .GetKeyValue("urls", "orderpadBase")
        g_QuoteEmailUri = .GetKeyValue("urls", "quoteEmail")
        g_AutoCommitUrl = .GetKeyValue("urls", "autoCommit")
        g_AutoPickUrl = .GetKeyValue("urls", "autoPick")
        g_PayPalUri = .GetKeyValue("urls", "payPal")
        g_MailServiceUrl = .GetKeyValue("urls", "mailService")
        g_CrCardReportURL = .GetKeyValue("urls", "creditCardReport")
        g_WASalesTaxUri = .GetKeyValue("urls", "waSalesTax")
        g_UPSOnlineURL = .GetKeyValue("urls", "upsOnline")
        g_ViewPageURL = .GetKeyValue("urls", "viewPage")
        
        
        'these are global settings which the user's registry entries are compared to
        'how do the grid revs match up against the registry entries and the database?
        
        g_ToolbarVersion = .GetKeyValue("constants", "ToolbarVersion")
        g_OrderGridRev = .GetKeyValue("constants", "OrderGridRev")
        g_CustOrderGridRev = .GetKeyValue("constants", "CustOrderGridRev")
        g_POWizGridRev = .GetKeyValue("constants", "POWizGridRev")
        g_PAVendorGridRev = .GetKeyValue("constants", "PAVendorGridRev")
    
    
        g_MPKHasCatalogs = .GetKeyValue("constants", "MPKHasCatalogs")
        g_STLHasCatalogs = .GetKeyValue("constants", "STLHasCatalogs")
        g_SEAHasCatalogs = .GetKeyValue("constants", "SEAHasCatalogs")
        g_SupportNotebooks = .GetKeyValue("constants", "SupportNotebooks")
        
        g_QueryForGift = .GetKeyValue("constants", "QueryForGift")
        g_CutOffDate = .GetKeyValue("constants", "CutOffDate")
        
        g_CryptoPassword = .GetKeyValue("constants", "cryptoPassword")
        
        g_MaxCustOrders = .GetKeyValue("constants", "MaxCustOrders")
        g_MaxCustRows = .GetKeyValue("constants", "MaxCustRows")
        g_MaxItemRows = .GetKeyValue("constants", "MaxItemRows")
    
        g_DisplayMask = .GetKeyValue("constants", "DisplayMask")
        g_MoneyMask = .GetKeyValue("constants", "MoneyMask")
    
        g_MaxHeightMagnetic = .GetKeyValue("constants", "MaxHeightMagnetic")
        g_MaxHeightCompression = .GetKeyValue("constants", "MaxHeightCompression")
    
        g_bWriteOrderXml = .GetKeyValue("constants", "WriteXmlToFile")
        
        g_FixedCostPerWire = .GetKeyValue("warmerwire", "FixedCostPerWire")
        g_CostPerFoot = .GetKeyValue("warmerwire", "CostPerFoot")
        g_WireCount = .GetKeyValue("warmerwire", "WireCount")
        g_MinSafeWPF = .GetKeyValue("warmerwire", "MinSafeWPF")
        g_MaxSafeWPF = .GetKeyValue("warmerwire", "MaxSafeWPF")
        g_BestSingleWPF = .GetKeyValue("warmerwire", "BestSingleWPF")
        g_BestDoubleWPF = .GetKeyValue("warmerwire", "BestDoubleWPF")
    
        g_sMPKDfltSchdID = .GetKeyValue("salestax", "MPKDfltSchdID")
        g_lMPKDfltSchdKey = .GetKeyValue("salestax", "MPKDfltSchdKey")
        g_dMPKDfltTaxRate = .GetKeyValue("salestax", "MPKDfltTaxRate")
    
        g_sSEADfltSchdID = .GetKeyValue("salestax", "SEADfltSchdID")
        g_lSEADfltSchdKey = .GetKeyValue("salestax", "SEADfltSchdKey")
        g_dSEADfltTaxRate = .GetKeyValue("salestax", "SEADfltTaxRate")
    
        g_sSTLDfltSchdID = .GetKeyValue("salestax", "STLDfltSchdID")
        g_lSTLDfltSchdKey = .GetKeyValue("salestax", "STLDfltSchdKey")
        g_dSTLDfltTaxRate = .GetKeyValue("salestax", "STLDfltTaxRate")
        
        g_sIStateDfltSchdID = .GetKeyValue("salestax", "IStateDfltSchdID")
        g_lIStateDfltSchdKey = .GetKeyValue("salestax", "IStateDfltSchdKey")
        g_dIStateDfltTaxRate = .GetKeyValue("salestax", "IStateDfltTaxRate")
        
        g_sIntlDfltSchdID = .GetKeyValue("salestax", "IntlDfltSchdID")
        g_lIntlDfltSchdKey = .GetKeyValue("salestax", "IntlDfltSchdKey")
        g_dIntlDfltTaxRate = .GetKeyValue("salestax", "IntlDfltTaxRate")
    
        g_sGovtDfltSchdID = .GetKeyValue("salestax", "GovtDfltSchdID")
        g_lGovtDfltSchdKey = .GetKeyValue("salestax", "GovtDfltSchdKey")
        g_dGovtDfltTaxRate = .GetKeyValue("salestax", "GovtDfltTaxRate")
    
    End With
    
End Sub


Private Sub CheckVersion()
    Dim lAppMajor As Long
    Dim lAppMinor As Long
    Dim lAppRevision As Long
    Dim bNewVersion As Boolean
    
    lAppMajor = g_UserConfig.GetKeyValue("officeassistant", "appMajor")
    lAppMinor = g_UserConfig.GetKeyValue("officeassistant", "appMinor")
    lAppRevision = g_UserConfig.GetKeyValue("officeassistant", "appRevision")

    If App.Major <> lAppMajor Then
        bNewVersion = True
    ElseIf App.Minor <> lAppMinor Then
        bNewVersion = True
    ElseIf App.Revision <> lAppRevision Then
        bNewVersion = True
    End If

    If bNewVersion Then
        If vbYes = msg("This is a new version of " & App.ProductName & " you haven't seen before." & vbCrLf _
                     & "Would you like to read the release notes?", vbYesNo, "Learn more about " & App.ProductName & " " & Version & "?") Then
            CommandHandler.DoShowBugs
        End If
    End If
    
End Sub


'Private Sub CheckVersion()
'    Dim lAppMajor As Long
'    Dim lAppMinor As Long
'    Dim lAppRevision As Long
'    Dim bNewVersion As Boolean
'
'    lAppMajor = GetRegNumberValue(HKEY_CURRENT_USER, g_RegKeyOP, "AppMajor", 0)
'    lAppMinor = GetRegNumberValue(HKEY_CURRENT_USER, g_RegKeyOP, "AppMinor", 0)
'    lAppRevision = GetRegNumberValue(HKEY_CURRENT_USER, g_RegKeyOP, "AppRevision", 0)
'
'    If App.Major > lAppMajor Then
'        bNewVersion = True
'    ElseIf App.Minor > lAppMinor Then
'        bNewVersion = True
'    ElseIf App.Revision > lAppRevision Then
'        bNewVersion = True
'    End If
'
'    If bNewVersion Then
'        If vbYes = msg("This is a new version of Office Assistant you haven't seen before." & vbCrLf _
'                     & "Would you like to read the release notes?", vbYesNo, "Learn more about Office Assistant " & Version & "?") Then
'            CommandHandler.DoShowBugs
'        End If
'    End If
'End Sub

