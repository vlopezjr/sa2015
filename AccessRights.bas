Attribute VB_Name = "AccessRights"
Option Explicit


Public Const k_sRightShowToolARCollections = "ShowTool:ARCollections"
Public Const k_sRightShowToolWillCall = "ShowTool:WillCall"
Public Const k_sRightShowToolDev = "ShowTool:Dev"
Public Const k_sRightShowToolAR = "ShowTool:A/R"
Public Const k_sRightShowToolAP = "ShowTool:A/P"
Public Const k_sRightShowToolOP = "ShowTool:OP"
Public Const k_sRightShowToolDashboard = "ShowTool:Dashboard"
Public Const k_sRightShowToolPurch = "ShowTool:Purch"
Public Const k_sRightShowToolRcv = "ShowTool:Rcv"
Public Const k_sRightShowToolBins = "ShowTool:Bins"
'Public Const k_sRightShowToolVaxAcct = "ShowTool:VaxAcct"
Public Const k_sRightShowToolUPSAcct = "ShowTool:UPSAcct"
Public Const k_sRightShowToolCrossRef = "ShowTool:CrossRef"
Public Const k_sRightShowToolManagement = "ShowTool:Management"
Public Const k_sRightShowToolPhoneFlagger = "ShowTool:PhoneFlagger"

Public Const k_sRightOPSaveOrder = "OP:SaveOrder"
Public Const k_sRightOPReleaseOrder = "OP:ReleaseOrder"

Public Const k_sRightARViewOnHold = "AR:ViewOnHold"
Public Const k_sRightARViewCustomer = "AR:ViewCustomer"
Public Const k_sRightARViewCollections = "AR:ViewCollections"
Public Const k_sRightARViewCredit = "AR:ViewCredit"
Public Const k_sRightARViewCreditCard = "AR:ViewCreditCard"
Public Const k_sRightARViewTenKey = "AR:ViewTenKey"
Public Const k_sRightARViewResearch = "AR:ViewResearch"
Public Const k_sRightARReleaseOrder = "AR:ReleaseOrder"
Public Const k_sRightAREditCollProfile = "AR:EditCollProfile"
Public Const k_sRightARUpdateStatus = "AR:UpdateStatus"
'Public Const k_sRightAPManageCost = "AP:ManageCost"
Public Const k_sRightShowToolViewPettyCashier = "AR:ViewPettyCashier"

'rights for viewing entire credit card number
Public Const k_sRightARViewCCNo = "AR:ViewCCNo"

Public Const k_sRightPurchasing = "Purchasing"
Public Const k_sRightReceiving = "Receiving"
Public Const k_sRightBillingAccount = "Billing:Account"
Public Const k_sRightBillingAssist = "Billing:Assist"
Public Const k_sRightBillingSalesTax = "Billing:SalesTax"
Public Const k_sRightBillingTemp = "Billing:Temp"
Public Const k_sRightBillingDropShip = "Billing:ViewDropShip"
Public Const k_sRightBillingRMACredMgr = "Billing:RMACredMgr"
Public Const k_sRightBillingRMAApprovalMgr = "Billing:RMAApprovalMgr"
Public Const k_sRightBillingSummary = "Billing:Summary"
Public Const k_sRightBillingWillCall = "Billing:ViewWillCall"
Public Const k_sRightBillingCreditCard = "Billing:ChargeCreditCard"

Public Const k_sRightAutoStartOP = "AutoStart:OP"
Public Const k_sRightAutoStartDashboard = "AutoStart:Dashboard"
Public Const k_sRightAutoStartPartsWiz = "AutoStart:PartsWiz"
Public Const k_sRightAutoStartDocFinder = "AutoStart:DocFinder"
Public Const k_sRightAutoStartInvFinder = "AutoStart:InvFinder"
Public Const k_sRightAutoStartAR = "AutoStart:AR"
Public Const k_sRightAutoStartAP = "AutoStart:AP"
Public Const k_sRightAutoStartPurchasing = "AutoStart:Purchasing"
Public Const k_sRightAutoStartBilling = "AutoStart:Billing"
Public Const k_sRightAutoStartPhoneFlagger = "AutoStart:PhoneFlagger"

Public Const k_sRightUpdateBillingAddr = "UpdateBillingAddr"

Private m_oAccessRight As AccessRight


Public Function HasRight(ByVal i_sRightID As String) As Boolean
    If m_oAccessRight Is Nothing Then
        Set m_oAccessRight = New AccessRight
    End If

    HasRight = m_oAccessRight.HasRight(i_sRightID)
End Function


Public Sub RefreshRights()
    Set m_oAccessRight = Nothing
End Sub


Public Function HasBillingRight() As Boolean
    HasBillingRight = HasRight(k_sRightBillingAssist) Or _
                        HasRight(k_sRightBillingAccount) Or _
                        HasRight(k_sRightBillingSalesTax) Or _
                        HasRight(k_sRightBillingTemp) Or _
                        HasRight(k_sRightBillingDropShip) Or _
                        HasRight(k_sRightBillingRMACredMgr) Or _
                        HasRight(k_sRightBillingSummary)
End Function


Public Function HasAnARRight() As Boolean
    HasAnARRight = HasRight(k_sRightARViewOnHold) Or _
        HasRight(k_sRightARViewCustomer) Or _
        HasRight(k_sRightARViewCollections) Or _
        HasRight(k_sRightARViewCredit) Or _
        HasRight(k_sRightARViewCreditCard) Or _
        HasRight(k_sRightARViewTenKey) Or _
        HasRight(k_sRightARViewResearch) Or _
        HasRight(k_sRightARReleaseOrder) Or _
        HasRight(k_sRightAREditCollProfile) Or _
        HasRight(k_sRightARUpdateStatus)
End Function


