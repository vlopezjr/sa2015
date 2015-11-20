Attribute VB_Name = "GlobalData"
Option Explicit

'*********************************************************************************************
' Global Constants

Public Const k_sCustNameOrID = "Customer Name/ID"
Public Const k_sCustName = "Customer Name"
'Public Const k_sCustVaxAcct = "Vax Account Number"
Public Const k_sCustID = "Sage Account Number"
Public Const k_sCustZip = "Customer Zip Code"
Public Const k_sCustPhone = "Customer Phone Number"
Public Const k_sNationalAccount = "National Account"
Public Const k_sPartNbr = "Part Number"
Public Const k_sPartDescr = "Part Description"

'PO Wiz LineStatus
'NOTE: The codes are used in associated stored procedures
Public Const kiRed = 1         'on SO
Public Const kiOrange = 2      'closer to zero than to Min
Public Const kiYellow = 3      'below Min
Public Const kiGreen = 4       'assembled kits
Public Const kiWhite = 5       'Added item

Public Const k_lGasketThreeSided = 1
Public Const k_lGasketDartToDart = 2
Public Const k_lGasketNoMagLHS = 4
Public Const k_lGasketInverted = 8
Public Const k_lGasketNoMagRHS = 16


'*********************************************************************************************
' Global Enums

Public Enum EnableAction
    opEnable
    opDisable
End Enum


'*********************************************************************************************
' Global Variables

Public g_UserConfig As Configuration
Public g_GlobalConfig As Configuration
    
'The Database Connection
Public g_DB As DBConnect

Public g_OrderCommitProxy As MSSOAPLib30.SoapClient30

Public g_MailProxy As MSSOAPLib30.SoapClient30

' Cached RecordSets

Public g_rstWhses As ADODB.Recordset
Public g_rstPartIDs As ADODB.Recordset
Public g_rstVendors As ADODB.Recordset
Public g_rstRMADisposition As ADODB.Recordset
'Public g_rstPrinters As ADODB.Recordset
Public g_rstGaskMats As ADODB.Recordset
Public g_rstWarmerWire As ADODB.Recordset
Public g_rstShipVia As ADODB.Recordset
Public g_rstStates As ADODB.Recordset
Public g_rstCustTypes As ADODB.Recordset
Public g_rstMakes As ADODB.Recordset
Public g_rstCountry As ADODB.Recordset
Public g_rstRMAReason As ADODB.Recordset


'Petty Cashier (why global?)
Public Const TRANSFERIN = 4
Public Const TRANSFEROUT = 5
Public Const CASHRECEIPT = 1
Public Const CASHREFUND = 2

Public g_bFilterCustSearch As Boolean
Public g_bATMMode As Boolean


'*****************************************************************************
'These flags are Teddy hacks. They're used to communicate between FOrder and
'MDIMain during Office Assistant shutdown.
'What if there's more than one FOrder instance?
'This needs study.
'
'Read by
'   FOrder.ExitCheck()
'   FOrder.Form_Unload()
'Set by
'   MDIMain.DoExit()

Public g_bConfirmExit As Boolean
Public g_bExitNow As Boolean


'*************************************************************************************
'These global variables map to Key/Value pairs in the Config.ini file.
'They are initialized at startup and are intented to be read-only.

'[Whse]
Public g_MPKWhseKey As Integer
Public g_SEAWhseKey As Integer
Public g_STLWhseKey As Integer
Public g_STLWhseID As String
Public g_SEAWhseID As String
Public g_MPKWhseID As String

'[Registry]
Public g_RegKeyOP As String
Public g_GridRegKey As String

'[Flags]
Public g_MPKHasCatalogs As Boolean
Public g_STLHasCatalogs As Boolean
Public g_SEAHasCatalogs As Boolean
Public g_SupportNotebooks As Boolean

'[SageSQLScripts]
Public g_CreatePOTempTables As String
Public g_DropPOTempTables As String
Public g_CreateIMTempTables As String
Public g_DropIMTempTables As String

'[GlobalPaths]
Public g_SnapshotPath As String
Public g_XsltPath As String
Public g_QuoteEmailPath As String
Public g_ReportPath As String
Public g_ToolbarConfigPath As String
Public g_CollectionsPath As String

'[URLs]
Public g_OrderPadBaseURL As String
Public g_QuoteEmailUri As String
Public g_CrCardReportURL As String
Public g_UPSOnlineURL As String
Public g_WASalesTaxUri As String
Public g_ViewPageURL As String
Public g_MailServiceUrl As String

Public g_AutoCommitUrl As String
Public g_PayPalUri As String
Public g_AutoPickUrl As String

'[Xmas]
Public g_QueryForGift As Boolean
Public g_CutOffDate As String

'[Crypto]
Public g_CryptoPassword As String

'[RevNbrs]
Public g_ToolbarVersion As Integer
Public g_OrderGridRev As Integer
Public g_CustOrderGridRev As Integer
Public g_POWizGridRev As Integer
Public g_PAVendorGridRev As Integer

'[SQLConstraints]
Public g_MaxCustOrders As Long   'Maximum rows returned by customer load
Public g_MaxCustRows As Long     'Maximum rows returned by customer queries
Public g_MaxItemRows As Long     'Maximum rows returned by item queries

'[Masks]
Public g_DisplayMask As String
Public g_MoneyMask As String

'[ItemWire]
Public g_FixedCostPerWire As Double ' = 2.5
Public g_CostPerFoot As Double ' = 0.09
Public g_WireCount As Integer
Public g_MinSafeWPF As Integer
Public g_MaxSafeWPF As Integer
Public g_BestSingleWPF As Integer
Public g_BestDoubleWPF As Integer

'[CustomGaskets]
Public g_MaxHeightMagnetic As Integer
Public g_MaxHeightCompression As Integer

'[SalesTax]
Public g_sMPKDfltSchdID As String
Public g_lMPKDfltSchdKey As Long
Public g_dMPKDfltTaxRate As Double

Public g_sSEADfltSchdID As String
Public g_lSEADfltSchdKey As Long
Public g_dSEADfltTaxRate As Double

Public g_sSTLDfltSchdID As String
Public g_lSTLDfltSchdKey As Long
Public g_dSTLDfltTaxRate As Double

Public g_sIStateDfltSchdID As String
Public g_lIStateDfltSchdKey As Long
Public g_dIStateDfltTaxRate As Double

Public g_sIntlDfltSchdID As String
Public g_lIntlDfltSchdKey As Long
Public g_dIntlDfltTaxRate As Double

Public g_sGovtDfltSchdID As String
Public g_lGovtDfltSchdKey As Long
Public g_dGovtDfltTaxRate As Double

'[OrderCommit]
Public g_bWriteOrderXml As Boolean

'[Collections]
'Public g_sCollectors As String

'[Profiler]
Public g_ProfilingEnabled As Boolean


