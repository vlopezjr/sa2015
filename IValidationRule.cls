VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IValidationRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public Property Get CtlWrapper() As ControlWrapper
End Property

Public Property Let CtlWrapper(ByVal RHS As ControlWrapper)
End Property


Public Property Get Enabled() As Boolean
End Property

Public Property Let Enabled(ByVal RHS As Boolean)
End Property


Public Property Get ClassID() As Long
End Property

Public Property Let ClassID(ByVal RHS As Long)
End Property


Public Property Get CustomMsg() As String
End Property

Public Property Let CustomMsg(ByVal RHS As String)
End Property


Public Property Get ID() As Long
End Property


Public Function Validate() As String
End Function

