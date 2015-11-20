VERSION 5.00
Begin VB.Form FChoosePmtTerms 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Commit Order On Hold"
   ClientHeight    =   1290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3540
   ForeColor       =   &H00C00000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1290
   ScaleWidth      =   3540
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   300
      Left            =   2580
      TabIndex        =   2
      Top             =   780
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   300
      Left            =   1620
      TabIndex        =   1
      Top             =   780
      Width           =   855
   End
   Begin VB.ComboBox cboPmtTerms 
      Height          =   315
      Left            =   1860
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   180
      Width           =   1572
   End
   Begin VB.Label Label2 
      Caption         =   "Order Payment Terms"
      Height          =   252
      Left            =   120
      TabIndex        =   3
      Top             =   180
      Width           =   1752
   End
End
Attribute VB_Name = "FChoosePmtTerms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'This dialog box is invoked by FAcctRcv.ReleaseOrder


Private m_oOrder As Order

Private m_bLoading As Boolean


Public Sub ReleaseAROrder(ByRef i_oOrder As Order)
    Set m_oOrder = i_oOrder
    
    m_oOrder.PmtTerms.LoadComboBox cboPmtTerms
    
    m_bLoading = True
    SetComboByText cboPmtTerms, m_oOrder.PmtTerms.ID
    m_bLoading = False
    
    Show vbModal

    Set m_oOrder = Nothing
    Unload Me
End Sub


'If the user choose credit card payment terms here, the behavior will be the
'same as that in FOrder.
'First, check if there is credit card history for the customer.
'Then show a dialog box for user to edit credit card information.

Private Sub cboPmtTerms_Click()
    
    If m_bLoading = True Then Exit Sub

    If cboPmtTerms.text = "CrCard" Then
    
        If m_oOrder.Customer.HasAccount Then
            Dim oFrm As FCreditCardEditor
            
            Set oFrm = New FCreditCardEditor
            If oFrm.Init(m_oOrder.Customer, Nothing, True, m_oOrder) = vbCancel Then
                'restore the combobox selection to what's on the order
                SetComboByText cboPmtTerms, m_oOrder.PmtTerms.ID
            Else
                m_oOrder.CreditCard = oFrm.SelCC
            End If
            Unload oFrm
            Set oFrm = Nothing

        End If
       
    End If
End Sub

Private Sub cmdCancel_Click()
    Me.Hide
End Sub


Private Sub cmdOK_Click()
    Dim sSubject As String
    Dim sTo As String
    Dim sMessage As String
    
    On Error GoTo ErrorHandler
    
    SetWaitCursor True

    m_oOrder.PmtTerms.Key = cboPmtTerms.ItemData(cboPmtTerms.ListIndex)
    
    m_oOrder.Commit
    
    LogOAEvent "Order On Hold", GetUserID, m_oOrder.OPKey, , 5, _
        "Committed by AR. The payment term is " _
        & cboPmtTerms.list(cboPmtTerms.ListIndex)

    If GetUserWhseID(m_oOrder.UserKey) <> "STL" Then
        SendNotification _
            "OP " & m_oOrder.OPKey & " for " & m_oOrder.Customer.ID _
            & " Committed by AR.", GetUserName & " committed OP " & m_oOrder.OPKey _
            & " from AR Hold" & vbCrLf & "for " & m_oOrder.Customer.ID & ": " _
            & m_oOrder.Customer.Name & vbCrLf & vbCrLf & "Current database is " _
            & g_DB.database, Array(GetUserID(m_oOrder.UserKey) & "@caseparts.com")
    End If

    SetWaitCursor False
   
Cleanup:
    ClearWaitCursor
    Me.Hide
    Exit Sub
    
ErrorHandler:
    DisplayWarning "Unexpected error when committing this order to Sage."
    
    GoTo Cleanup
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Me.Hide
End Sub

