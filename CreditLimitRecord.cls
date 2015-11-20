VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CreditLimitRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Update As Boolean
Public NewLimit As Long
Public CurLimit As Long
Public YearRec As Long
Public ColRems As String
Public CustID As String
Public CustName As String
Public CustKey As Long
Public LastChanged As Variant
Public CurAmount As Long
Public l30DaysAmount As Long
Public l45DaysAmount As Long
Public TerritoryKey As Long
Public Terms As String
Public AgingDate As Variant

