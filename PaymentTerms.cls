VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PaymentTerms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Each Order has payment terms managed by the CSR (within constraints).
' The Customer has default terms managed by AR.
' Default Payment Terms are persisted in tarCustAddr and are loaded into the
' BillTo Address object.

' This class is used in a few places to simply load comboboxes (in place of g_rstTerms).

'Created 7/14/04 LR
'   This class supports determining sPmtTermsID & iDueDayOrMonth based on lPmtTermsKey

'as of 11/8/04 this is always the Default payment terms
'These are valid properties for Billing Address only
'If the key is changed, the ID & DueDayOrMonth need to change

'Updated 7/20/05 LR
'   Significantly expand the nature and scope of this class.


Private Type TermsProps
    lPmtTermsKey As Long
    sPmtTermsID As String
    iDueDayOrMonth As Integer
    m_bIsDirty As Boolean
End Type

Private m_orstTerms As ADODB.Recordset

Private m_udtTerms As TermsProps
Private m_udtTermsBackup As TermsProps


Private Sub Class_Initialize()
    Dim sSQL As String
    
    sSQL = "SELECT PmtTermsKey, PmtTermsID, DueDayOrMonth " _
         & "FROM tciPaymentTerms " _
         & "WHERE CompanyID = 'CPC' AND DiscDateOption = 1 " _
         & "ORDER BY PmtTermsID"
    Set m_orstTerms = LoadDiscRst(sSQL)
    
    m_udtTerms.m_bIsDirty = False
End Sub


Private Sub Class_Terminate()
    Set m_orstTerms = Nothing
End Sub


' PUBLIC PROPERTIES

Public Property Get Key() As Long
    Key = m_udtTerms.lPmtTermsKey
End Property


Public Property Let Key(ByVal lNewValue As Long)
    If lNewValue <> m_udtTerms.lPmtTermsKey Then
        m_udtTerms.lPmtTermsKey = lNewValue
        m_udtTerms.sPmtTermsID = SetPmtTermsID(lNewValue)
        m_udtTerms.iDueDayOrMonth = SetDueDayOrMonth(lNewValue)
        m_udtTerms.m_bIsDirty = True
    End If
End Property


Public Property Get PreviousKey() As Long
    PreviousKey = m_udtTermsBackup.lPmtTermsKey
End Property


Public Property Get PreviousID() As String
    PreviousID = m_udtTermsBackup.sPmtTermsID
End Property


Public Property Get ID() As String
    ID = m_udtTerms.sPmtTermsID
End Property


Public Property Get DueDayOrMonth() As Integer
    DueDayOrMonth = m_udtTerms.iDueDayOrMonth
End Property


Public Property Get IsDirty() As Boolean
    IsDirty = m_udtTerms.m_bIsDirty
End Property


'COD, COD-Cash
Public Property Get IsCOD() As Boolean
    If Left(UCase(m_udtTerms.sPmtTermsID), 3) <> "COD" Then
        IsCOD = False
    Else
        IsCOD = True
    End If
End Property


'COD, COD-Cash, CrCard, Net on Receipt, PrePaid, WC CrCard
Public Property Get IsCash() As Boolean
    If m_udtTerms.iDueDayOrMonth <= 1 Then
        IsCash = True
    Else
        IsCash = False
    End If
End Property



' PUBLIC METHODS

Public Sub Load(lKey As Long)
    m_udtTerms.lPmtTermsKey = lKey
    m_udtTerms.sPmtTermsID = SetPmtTermsID(lKey)
    m_udtTerms.iDueDayOrMonth = SetDueDayOrMonth(lKey)
    Backup      'for change event logic
End Sub


Public Sub Clear()
    m_udtTerms.lPmtTermsKey = 0
    m_udtTerms.sPmtTermsID = ""
    m_udtTerms.iDueDayOrMonth = 0
    m_udtTermsBackup.lPmtTermsKey = 0
    m_udtTermsBackup.sPmtTermsID = ""
    m_udtTermsBackup.iDueDayOrMonth = 0
End Sub


Public Sub Backup()
    m_udtTermsBackup = m_udtTerms
End Sub


Public Sub Restore()
    m_udtTerms = m_udtTermsBackup
End Sub


Public Sub LoadComboBox(ByRef cbo As ComboBox)
    Helpers.LoadCombo cbo, m_orstTerms, "PmtTermsID", "PmtTermsKey"
End Sub


'Used to restrict the entries in the combobox

'Called by
'   FOrder.UpdateOrderInfo

'From FOrder
'        .FilterTerms g_rstTerms       'This has the side effect of leaving g_rstTerms filtered
'        LoadCombo cboTerms, g_rstTerms, "PmtTermsID", "PmtTermsKey", .BillAddr.PmtTermsKey, False
'        If .BillAddr.PmtTermsKey = 0 Then
'            If .CustType = "EndUser" Then
'                SetComboByText cboTerms, "COD"
'            Else
'                SetComboByText cboTerms, "N30"
'            End If
'            .BillAddr.PmtTermsKey = cboTerms.ItemData(cboTerms.ListIndex)
'        Else
'            SetComboByKey cboTerms, .BillAddr.PmtTermsKey
'        End If

'Moved from Customer.cls and redefined parameters.  7/6/04 LR
'Note: This leaves the global recordset filtered. Not so good.

'Public Sub FilterTerms(ByRef rstTerms As ADODB.Recordset)

Public Sub LoadFilteredComboBox(cbo As ComboBox, ByVal PmtTermsKey As Long)
    Dim lThresholdDueDayOrMonth As Integer

    'Get the 'Days' for this customer's Billing address' selected terms
    'Filter where 'Days' are <= threshold
    If PmtTermsKey <> 0 Then
'removed this guard condition 11/2/04 LR
'        'if not Credit Card
'        If PmtTermsKey <> 41 Then
            m_orstTerms.Filter = "PmtTermsKey = " & PmtTermsKey
            lThresholdDueDayOrMonth = m_orstTerms.Fields("DueDayOrMonth").value
            m_orstTerms.Filter = "DueDayOrMonth < " & CStr(lThresholdDueDayOrMonth + 2)
'        End If
    End If
    Helpers.LoadCombo cbo, m_orstTerms, "PmtTermsID", "PmtTermsKey", PmtTermsKey, False
    m_orstTerms.Filter = adFilterNone
End Sub


' PRIVATE FUNCTIONS

Private Function SetPmtTermsID(PmtTermsKey As Long) As String
    If PmtTermsKey = 0 Then
        SetPmtTermsID = ""
    Else
        With m_orstTerms
            .Filter = "PmtTermsKey = " & PmtTermsKey
            SetPmtTermsID = Trim(.Fields("PmtTermsID").value)
            .Filter = adFilterNone
        End With
    End If
End Function


Private Function SetDueDayOrMonth(PmtTermsKey As Long) As Integer
    With m_orstTerms
        .Filter = "PmtTermsKey = " & PmtTermsKey
        SetDueDayOrMonth = .Fields("DueDayOrMonth").value
        .Filter = adFilterNone
    End With
End Function



