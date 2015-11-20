VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AcctPay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public bUpdate As Boolean
Public bColor As Boolean
Public VendID As String
Public Vendor As String
Public NewReplCost As Double
Public CurReplCost As Double
Public CPCPartNbr As String
Public VendPartNbr As String
Public Descr As String
Public CreateDate As Date
Public BatchID As String
Public OrigUserID As String
Public VendKey As Long
Public ItemKey As Long
Public MaxQty As Long
Public MinQty As Long
Public NewMax As Long
Public NewMin As Long

'The following items are used for creating Vend/Item relationship in timVendItem
Public ItemAliasKey As Long
Public BreakType As Integer
Public DiscMeth As Integer
Public ListPrice As Double
Public OrigCountry As String
Public PurchUnitMeasKey As Long
Public SubjToVendFamDisc As Boolean
Public VendFamilyKey As Long




