VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RMAOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public RMALineKey As Long

Public RMAInfo As String
'06/12/03 AVH PRN#5 Provide an option to sort the grid by RMA# or by CustID
Public RMAInfo_ByCustID As String
Public RMAKey As Long
Public OPKey As Long
Public SOID As Long
Public OPLineKey As Long
Public SOLineKey As Long
Public OPItemType As ItemTypeCode ' VL 08/08/14

Public Authorized As Boolean
Public Approved As Boolean
Public Credited As Boolean
Public CreditFreight As Boolean

Public Reason As Integer
Public Disposition As Integer
Public Restock As Double

Public AuthBy As String
Public AuthDate As Date

Public ItemID As String
Public Descr As String
Public CustID As String
Public CustName As String

Public Cost As Double
Public Price As Double
Public ExtPrice As Double

Public QtyAuthorized As Integer
Public QtyRcvd As Integer
'06/12/03 AVH PRN#103
Public QtyRcvdOriginal As Integer
Public QtyCred As Integer

Public QtyPreRcvd As Integer
Public QtyPreCred As Integer

Public lMaxQtyAuth As Long
Public lMaxQtyRcvd As Long
Public lMaxQtyCred As Long
Public RcvdWhseID As String


'02/19/2003     TeddyX
'This three new attributes are used for RMA Vendor Return
Public ReturnToVendor As Boolean
Public DaysNoPenalty As Long
Public VendorRMANumber As String
Public VendKey As Long


Public ReceiveDate As Variant
Public ApproveDate As Variant
Public CMNbr As String





