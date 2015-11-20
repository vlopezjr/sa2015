VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ViewerParams"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_sParamName As String
Private m_sParamValue As Variant


Public Property Get ParamName() As String
    ParamName = m_sParamName
End Property
Public Property Let ParamName(Value As String)
    m_sParamName = Value
End Property


Public Property Get ParamValue() As Variant
    ParamValue = m_sParamValue
End Property
Public Property Let ParamValue(Value As Variant)
    m_sParamValue = Value
End Property
