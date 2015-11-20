VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UPSAcct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Exist As String
Public Selected As Boolean
Public AddrKey As Long
Public AddrName As String
Public AddrLine1 As String
Public AddrLine2 As String
Public City As String
Public PostalCode As String
Public StateID As String
Public UPSAcct As String
Public AddrType As String
