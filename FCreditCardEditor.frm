VERSION 5.00
Object = "{CAF0FDE4-8332-11CF-BC13-0020AFD6738C}#1.0#0"; "newsota.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form FCreditCardEditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Credit Card Selection"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8070
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   8070
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   1455
      Left            =   5760
      TabIndex        =   24
      Top             =   5400
      Width           =   2175
      Begin VB.CommandButton cmdtest 
         Caption         =   "Test"
         Height          =   375
         Left            =   1080
         TabIndex        =   28
         Top             =   720
         Width           =   975
      End
      Begin VB.CheckBox chkIsValid 
         Caption         =   "IsValid"
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   960
         Width           =   1095
      End
      Begin VB.CheckBox chkIsDirty 
         Caption         =   "IsDirty"
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   720
         Width           =   1095
      End
      Begin VB.CheckBox chkIsNew 
         Caption         =   "IsNew"
         Height          =   375
         Left            =   120
         TabIndex        =   27
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label lblTest 
         Caption         =   "# of Broken Rules:"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblBR 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1560
         TabIndex        =   29
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   23
      Top             =   4020
      Width           =   7815
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   6480
         TabIndex        =   13
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "&Select"
         Default         =   -1  'True
         Height          =   375
         Left            =   5400
         TabIndex        =   12
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdUndo 
         Caption         =   "&Undo"
         Height          =   375
         Left            =   3960
         TabIndex        =   11
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "&New"
         Height          =   375
         Left            =   360
         TabIndex        =   8
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdAddUpdate 
         Caption         =   "&Add"
         Height          =   375
         Left            =   1440
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   2520
         TabIndex        =   10
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame fraInfo 
      Height          =   1335
      Left            =   120
      TabIndex        =   14
      Top             =   2700
      Width           =   7815
      Begin VB.TextBox txtCardNoMask 
         Height          =   288
         Left            =   3780
         TabIndex        =   1
         Top             =   240
         Width           =   1815
      End
      Begin VB.CheckBox chkIsPrefCardNbr 
         Alignment       =   1  'Right Justify
         Caption         =   "Preferred Card"
         Height          =   255
         Left            =   90
         TabIndex        =   7
         Top             =   960
         Width           =   1428
      End
      Begin NEWSOTALib.SOTAMaskedEdit txtCardNo 
         Height          =   288
         Left            =   3780
         TabIndex        =   2
         Top             =   240
         Width           =   1812
         _Version        =   65536
         _ExtentX        =   3196
         _ExtentY        =   508
         _StockProps     =   93
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FillChar        =   "*"
      End
      Begin VB.ComboBox cboCCType 
         Height          =   288
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtCardStreetNbr 
         Height          =   288
         Left            =   6720
         MaxLength       =   30
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtCardHolderName 
         Height          =   288
         Left            =   3780
         MaxLength       =   30
         TabIndex        =   4
         Top             =   600
         Width           =   1815
      End
      Begin NEWSOTALib.SOTAMaskedEdit txtCardExp 
         Height          =   288
         Left            =   960
         TabIndex        =   3
         Top             =   600
         Width           =   1212
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   503
         _StockProps     =   93
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin NEWSOTALib.SOTAMaskedEdit txtCardZipCode 
         Height          =   285
         Left            =   6720
         TabIndex        =   6
         Top             =   600
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   503
         _StockProps     =   93
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Zip Code"
         Height          =   252
         Left            =   5760
         TabIndex        =   22
         Top             =   660
         Width           =   852
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Street Nbr"
         Height          =   252
         Left            =   5760
         TabIndex        =   21
         Top             =   300
         Width           =   852
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Card Holder Name"
         Height          =   252
         Left            =   2280
         TabIndex        =   20
         Top             =   660
         Width           =   1392
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Exp Date"
         Height          =   252
         Left            =   60
         TabIndex        =   19
         Top             =   660
         Width           =   792
      End
      Begin VB.Label lblKey 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   11040
         TabIndex        =   17
         Top             =   720
         Width           =   495
      End
      Begin VB.Label lblCCKey 
         Caption         =   "CC Key:"
         Height          =   255
         Left            =   11040
         TabIndex        =   18
         Top             =   480
         Width           =   615
      End
      Begin VB.Label lblFirstName 
         Alignment       =   1  'Right Justify
         Caption         =   "Card Type"
         Height          =   252
         Left            =   60
         TabIndex        =   16
         Top             =   300
         Width           =   792
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Card Number"
         Height          =   252
         Left            =   2340
         TabIndex        =   15
         Top             =   300
         Width           =   1332
      End
   End
   Begin GridEX20.GridEX gdxCreditCards 
      Height          =   2055
      Left            =   180
      TabIndex        =   53
      Top             =   600
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   3625
      Version         =   "2.0"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      HideSelection   =   2
      UseEvenOddColor =   -1  'True
      MethodHoldFields=   -1  'True
      AllowDelete     =   -1  'True
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      ItemCount       =   1
      DataMode        =   99
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   9
      Column(1)       =   "FCreditCardEditor.frx":0000
      Column(2)       =   "FCreditCardEditor.frx":0144
      Column(3)       =   "FCreditCardEditor.frx":0284
      Column(4)       =   "FCreditCardEditor.frx":03A0
      Column(5)       =   "FCreditCardEditor.frx":04BC
      Column(6)       =   "FCreditCardEditor.frx":05E4
      Column(7)       =   "FCreditCardEditor.frx":06F0
      Column(8)       =   "FCreditCardEditor.frx":0828
      Column(9)       =   "FCreditCardEditor.frx":0950
      FormatStylesCount=   7
      FormatStyle(1)  =   "FCreditCardEditor.frx":0A70
      FormatStyle(2)  =   "FCreditCardEditor.frx":0B50
      FormatStyle(3)  =   "FCreditCardEditor.frx":0C88
      FormatStyle(4)  =   "FCreditCardEditor.frx":0D38
      FormatStyle(5)  =   "FCreditCardEditor.frx":0DEC
      FormatStyle(6)  =   "FCreditCardEditor.frx":0EC4
      FormatStyle(7)  =   "FCreditCardEditor.frx":0F7C
      ImageCount      =   0
      PrinterProperties=   "FCreditCardEditor.frx":0F9C
   End
   Begin VB.Label Label26 
      Caption         =   "False"
      Height          =   255
      Left            =   5040
      TabIndex        =   52
      Top             =   5520
      Width           =   375
   End
   Begin VB.Label Label25 
      Caption         =   "False"
      Height          =   255
      Left            =   5040
      TabIndex        =   51
      Top             =   6000
      Width           =   375
   End
   Begin VB.Label Label10 
      Caption         =   "True"
      Height          =   255
      Left            =   5040
      TabIndex        =   50
      Top             =   6480
      Width           =   375
   End
   Begin VB.Label lblCustName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   49
      Top             =   120
      Width           =   4095
   End
   Begin VB.Label lblCustID 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   48
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label Label24 
      Caption         =   "False"
      Height          =   255
      Left            =   1200
      TabIndex        =   47
      Top             =   5520
      Width           =   375
   End
   Begin VB.Label Label23 
      Caption         =   "True"
      Height          =   255
      Left            =   2520
      TabIndex        =   46
      Top             =   6480
      Width           =   375
   End
   Begin VB.Label Label22 
      Caption         =   "True"
      Height          =   255
      Left            =   2520
      TabIndex        =   45
      Top             =   6000
      Width           =   375
   End
   Begin VB.Line Line11 
      X1              =   240
      X2              =   240
      Y1              =   5400
      Y2              =   6840
   End
   Begin VB.Line Line10 
      X1              =   5640
      X2              =   5640
      Y1              =   6840
      Y2              =   5400
   End
   Begin VB.Line Line9 
      X1              =   240
      X2              =   5640
      Y1              =   5400
      Y2              =   5400
   End
   Begin VB.Label Label21 
      Caption         =   "Ok"
      Height          =   255
      Left            =   5040
      TabIndex        =   44
      Top             =   5160
      Width           =   735
   End
   Begin VB.Label Label20 
      Caption         =   "Undo"
      Height          =   255
      Left            =   4200
      TabIndex        =   43
      Top             =   5160
      Width           =   735
   End
   Begin VB.Label Label19 
      Caption         =   "Delete"
      Height          =   255
      Left            =   3360
      TabIndex        =   42
      Top             =   5160
      Width           =   615
   End
   Begin VB.Label Label18 
      Caption         =   "New"
      Height          =   255
      Left            =   1200
      TabIndex        =   41
      Top             =   5160
      Width           =   375
   End
   Begin VB.Label Label17 
      Caption         =   "Add/Update"
      Height          =   255
      Left            =   2040
      TabIndex        =   40
      Top             =   5160
      Width           =   975
   End
   Begin VB.Label Label16 
      Caption         =   "False"
      Height          =   255
      Left            =   2520
      TabIndex        =   39
      Top             =   5520
      Width           =   375
   End
   Begin VB.Label Label15 
      Caption         =   "True"
      Height          =   255
      Left            =   4200
      TabIndex        =   38
      Top             =   6000
      Width           =   375
   End
   Begin VB.Label Label14 
      Caption         =   "False"
      Height          =   255
      Left            =   3360
      TabIndex        =   37
      Top             =   5520
      Width           =   375
   End
   Begin VB.Label Label13 
      Caption         =   "True"
      Height          =   255
      Left            =   2040
      TabIndex        =   36
      Top             =   6480
      Width           =   375
   End
   Begin VB.Label Label12 
      Caption         =   "True"
      Height          =   255
      Left            =   2040
      TabIndex        =   35
      Top             =   6000
      Width           =   375
   End
   Begin VB.Label Label11 
      Caption         =   "True"
      Height          =   255
      Left            =   2040
      TabIndex        =   34
      Top             =   5520
      Width           =   375
   End
   Begin VB.Line Line8 
      X1              =   240
      X2              =   5640
      Y1              =   6840
      Y2              =   6840
   End
   Begin VB.Line Line7 
      X1              =   4920
      X2              =   4920
      Y1              =   5400
      Y2              =   6840
   End
   Begin VB.Line Line6 
      X1              =   4080
      X2              =   4080
      Y1              =   5400
      Y2              =   6840
   End
   Begin VB.Line Line5 
      X1              =   3120
      X2              =   3120
      Y1              =   5400
      Y2              =   6840
   End
   Begin VB.Line Line4 
      X1              =   1800
      X2              =   1800
      Y1              =   5400
      Y2              =   6840
   End
   Begin VB.Line Line3 
      X1              =   240
      X2              =   5640
      Y1              =   6360
      Y2              =   6360
   End
   Begin VB.Line Line2 
      X1              =   240
      X2              =   5640
      Y1              =   5880
      Y2              =   5880
   End
   Begin VB.Line Line1 
      X1              =   1080
      X2              =   1080
      Y1              =   5400
      Y2              =   6840
   End
   Begin VB.Label Label9 
      Caption         =   "IsValid"
      Height          =   255
      Left            =   360
      TabIndex        =   33
      Top             =   6480
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "IsDirty"
      Height          =   255
      Left            =   360
      TabIndex        =   32
      Top             =   6000
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "IsNew"
      Height          =   255
      Left            =   360
      TabIndex        =   31
      Top             =   5520
      Width           =   615
   End
End
Attribute VB_Name = "FCreditCardEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const k_lClassAlways = 1

Private m_lCustKey As Long
Private m_sCustType As String
Private m_vGridData() As Variant
Private m_lGridRowCount As Long

Private m_oCustomer As Customer

Private WithEvents m_oCreditCardMgr As CreditCardMgr
Attribute m_oCreditCardMgr.VB_VarHelpID = -1
Private WithEvents m_oBrokenRules As BrokenRules
Attribute m_oBrokenRules.VB_VarHelpID = -1

Private m_bLoading As Boolean

Private m_RetCode As Integer


'*********************************************************************
' Public Properties
'*********************************************************************

Public Property Get SelCC() As CreditCard
    Set SelCC = m_oCreditCardMgr.SelCC
End Property


Public Property Get CustKey() As Long
    CustKey = m_lCustKey
End Property

Public Property Let CustKey(ByVal lNewValue As Long)
    m_lCustKey = lNewValue
End Property


Public Property Get CustType() As String
    CustType = Trim$(m_sCustType)
End Property

Public Property Let CustType(ByVal sNewValue As String)
    m_sCustType = Trim$(sNewValue)
End Property


'*********************************************************************
' Form Event Handlers
'*********************************************************************

Private Sub Form_Load()
    '*** NOTE: Before placing code in form_load, see Public Sub Init ***
End Sub


Private Sub Form_Unload(Cancel As Integer)
    m_oBrokenRules.Destroy
    Set m_oBrokenRules = Nothing
End Sub


'*********************************************************************
' Public Methods
'*********************************************************************

Public Function Init(oCust As Customer, oPrevCC As CreditCard, Optional bShow As Boolean = False, Optional oOrder As Variant) As Integer
    Set m_oCustomer = oCust
    CustType = oCust.CustType
    
    Set m_oCreditCardMgr = New CreditCardMgr
    
    If IsMissing(oOrder) Then
        m_oCreditCardMgr.LoadCreditCards oCust, oPrevCC 'oCust.Key, oCust.ID, oPrevCC
    Else
        m_oCreditCardMgr.LoadCreditCards oCust, oPrevCC, oOrder
    End If
    
    If bShow Then
        InitShow
        Me.Show vbModal
    Else
        If m_oCreditCardMgr.CreditCards.Count = 0 Then
            'sel cc will already equal nothing
            'set return code to vbCancel
            m_RetCode = vbCancel
        'End If
        
        '*** CDate was not properly coverting this specific date (MMYYYY)
            'typed as a string, to a date type for comparison***
            
            'Example Now = 052006
            
            '    Year=  2004      2006      2008
            'Mth= 02    Expired   Expired
            'Mth= 05    Expired
            'Mth= 07    Expired
                
        ElseIf (Right(m_oCreditCardMgr.SelCC.ExpireDate, 4) < Format(Now, "YYYY")) Then
            m_RetCode = vbCancel
        
        ElseIf (Right(m_oCreditCardMgr.SelCC.ExpireDate, 4) = Format(Now, "YYYY")) Then
            If (Mid(m_oCreditCardMgr.SelCC.ExpireDate, 1, 2) < Format(Now, "MM")) Then
                m_RetCode = vbCancel
                
            End If
        End If
    End If
    
    Init = m_RetCode
End Function


Private Function InitShow()
        'this assignment causes the Form_Load method to fire
        lblCustID.caption = m_oCustomer.ID
            
        'MISC customers - Note: the Customer.IsMiscellaneous property doesn't work.

        If Not m_oCustomer.HasAccount Then
            lblCustName.caption = m_oCustomer.ShipAddr.AddrName
        Else
            lblCustName.caption = m_oCustomer.name
        End If
        
        If m_oCreditCardMgr.CreditCards.Count = 0 Then
            m_oCreditCardMgr.NewCreditCard
        End If
        
        'if there's a selected card object
        If Not m_oCreditCardMgr.SelCC Is Nothing Then
            m_bLoading = True
            
            'New Validation logic - only loaded once & before BrokenRules.Validate can be called
            LoadValidationRules
            
            gdxCreditCards.ItemCount = m_oCreditCardMgr.CreditCards.Count
            GridColAutoSize gdxCreditCards
            
            'Select row in grid
            Dim i As Integer
            For i = 1 To gdxCreditCards.RowCount
                If m_oCreditCardMgr.CreditCards.Item(i).key = m_oCreditCardMgr.SelCC.key Then
                    gdxCreditCards.RowSelected(i) = True
                    gdxCreditCards.Refresh
                    Exit For
                End If
            Next
            
            DisplayCreditCard
            'Check the pref CC checkbox - if we are in the new state with no assigned CCs
            If m_oCreditCardMgr.CreditCards.Count = 0 Then
                chkIsPrefCardNbr.value = vbChecked
                m_oCreditCardMgr.SelCC.Preferred = True
            'check if this customer has at least 1 assigned credit card
            ElseIf m_oCreditCardMgr.CreditCards.Count < 2 Then
                chkIsPrefCardNbr.Enabled = False
            End If
            m_bLoading = False
        End If
End Function


'*********************************************************************
' Object Event Handlers
'*********************************************************************

Private Sub m_oCreditCardMgr_StateChange(NewState As BitMap)
    SetButtons NewState
End Sub


'*********************************************************************
' Control Event Handlers
'*********************************************************************

' BUTTONS

Private Sub cmdAddUpdate_Click()
    AddUpdate
    cboCCType.SetFocus
    gdxCreditCards.Refresh
End Sub


Private Sub cmdCancel_Click()
    If SaveValidChgs Then
        AddUpdate
    Else
        m_oCreditCardMgr.CancelChanges
    End If

    If m_oCreditCardMgr.PrevCC Is Nothing Then
        m_oCreditCardMgr.SelCC = Nothing
    Else
        m_oCreditCardMgr.SelCC = m_oCreditCardMgr.PrevCC
    End If

    m_RetCode = vbCancel
    Me.Hide
End Sub


Private Sub cmdUndo_Click()
    m_oCreditCardMgr.CancelChanges
    DisplayCreditCard
    cboCCType.SetFocus
End Sub


Private Sub cmdDelete_Click()

    'Test if deleting pref credit card
    'this should read chkIsPrefCardNbr.Value = vbChecked
    If chkIsPrefCardNbr.value = 1 And m_oCreditCardMgr.CreditCards.Count > 1 Then
        MsgBox "Please select another credit card as preferred.", vbInformation
        chkIsPrefCardNbr.SetFocus
        Exit Sub
    End If


'Shouldn't this be in the CreditCardMgr?
    'added
    If Not m_oCreditCardMgr.PrevCC Is Nothing Then
        If m_oCreditCardMgr.SelCC.key = m_oCreditCardMgr.PrevCC.key Then
            m_oCreditCardMgr.PrevCC = Nothing
        End If
    End If

'We're deleting it from the grid.
'Why aren't we deleting it from the collection and the database?
'Ah! this causes the grid's UnboundDelete event to fire which in turn
'calls CreditCardMgr.DeleteCreditCard()

    gdxCreditCards.Delete
        
    gdxCreditCards.RowSelected(1) = True

'Shouldn't this be in the CreditCardMgr?

    If m_oCreditCardMgr.CreditCards.Count = 0 Then
        m_oCreditCardMgr.NewCreditCard
    End If

    DisplayCreditCard
    cboCCType.SetFocus
        
End Sub


Private Sub cmdNew_Click()
    m_oCreditCardMgr.NewCreditCard
    DisplayCreditCard
    cboCCType.SetFocus
End Sub


Private Sub cmdOK_Click()
    m_RetCode = vbOK
    Me.Hide
End Sub


'This event will never fire for a Dropdown List

Private Sub cboCCType_Change()
    m_oCreditCardMgr.SelCC.State.ClearBits eMask.IsDirty
End Sub


Private Sub cboCCType_Click()
   
    If Not m_bLoading Then

        m_oCreditCardMgr.SelCC.State.ClearBits eMask.IsDirty
        
        m_oBrokenRules.Validate cboCCType
        SetIsValid
    End If
    
    'do we want to do this when loading the combo?
    With m_oCreditCardMgr.SelCC
        .TypeKey = cboCCType.ItemData(cboCCType.ListIndex)
        .TypeID = cboCCType.text
        
        'set mask for selected credit card type in txtCardNo
        txtCardNo.mask = m_oCreditCardMgr.GetCardTypeMask(.TypeKey)
        txtCardNo.PromptChar = "*"
    End With
    
    If Not m_bLoading Then
        'need to also validate CC#, after cboCCType changes & the CCType Mask is loaded in the txtCardNo field
        m_oBrokenRules.Validate txtCardNo
        SetIsValid
    End If
End Sub


Private Sub chkIsPrefCardNbr_Click()
    If Not m_bLoading Then
        If chkIsPrefCardNbr.value = vbUnchecked Then
            Dim aCreditCard As CreditCard
            For Each aCreditCard In m_oCreditCardMgr.CreditCards ' m_cColCC
                If aCreditCard.Preferred And aCreditCard.key <> m_oCreditCardMgr.SelCC.key Then
                    Exit For
                End If
            Next
            MsgBox "Please select another credit card as preferred.", vbInformation
            DisplayCreditCard
        Else
            m_oCreditCardMgr.SelCC.State.SetBits eMask.IsDirty
        End If
    End If
End Sub


Private Sub chkIsPrefCardNbr_LostFocus()
        m_oBrokenRules.Validate chkIsPrefCardNbr
        SetIsValid
        m_oCreditCardMgr.SelCC.Preferred = chkIsPrefCardNbr.value
End Sub


Private Sub txtCardExp_Change()
    If Not m_bLoading Then
        m_oCreditCardMgr.SelCC.State.SetBits eMask.IsDirty
    End If
End Sub


Private Sub txtCardExp_LostFocus()
        m_oBrokenRules.Validate txtCardExp
        SetIsValid
        m_oCreditCardMgr.SelCC.ExpireDate = txtCardExp.text
End Sub


Private Sub txtCardHolderName_Change()
    If Not m_bLoading Then
        m_oCreditCardMgr.SelCC.State.SetBits eMask.IsDirty
    End If
End Sub


Private Sub txtCardHolderName_LostFocus()
'    If Not m_bLoading Then
        m_oBrokenRules.Validate txtCardHolderName
        SetIsValid
        m_oCreditCardMgr.SelCC.CardHolderName = txtCardHolderName.text
'    End If
End Sub


Private Sub txtCardNo_Change()
    If Not m_bLoading Then
        m_oCreditCardMgr.SelCC.State.SetBits eMask.IsDirty
    End If
End Sub


Private Sub txtCardNo_LostFocus()

        m_oCreditCardMgr.SelCC.CardNo = txtCardNo.text
        
        'Check if duplicate creditcard# for the same CustKey by looping through collection of CC, elimating the selected CC
        Dim aCreditCard As CreditCard
        For Each aCreditCard In m_oCreditCardMgr.CreditCards ' m_colCreditCards
            If m_oCreditCardMgr.SelCC.CardNo = aCreditCard.CardNo And m_oCreditCardMgr.SelCC.key <> aCreditCard.key Then
            MsgBox "Credit card number already exists for this customer." & chr(10) & chr(13) & _
                chr(10) & chr(13) & "Please change the credit card number."
                txtCardNo.SetFocus
                  Exit Sub
            End If
        Next
        
        m_oBrokenRules.Validate txtCardNo
        SetIsValid

End Sub


Private Sub txtCardNoMask_GotFocus()
    txtCardNoMask.SelStart = 0
    txtCardNoMask.SelLength = Len(txtCardNoMask.text)
End Sub


Private Sub txtCardNoMask_KeyUp(KeyCode As Integer, Shift As Integer)
    'This text box will display when the credit card is in an editable state.
        'Once the first number is typed in, the txtCardNoMask text box will disappear
        'and txtCardNo SOTAMaskedEdit control will be visible.
        
    txtCardNoMask.Visible = False
    txtCardNo.ClearData
    txtCardNo.Visible = True
    txtCardNo.SetFocus
    'only one new char is typed in at this point
    m_oCreditCardMgr.SelCC.CardNo = Mid$(txtCardNoMask.text, 1, 1)
    txtCardNo.MaskedText = m_oCreditCardMgr.SelCC.CardNo
    txtCardNo.SetSel 1, 2

End Sub


Private Sub txtCardStreetNbr_Change()
    If Not m_bLoading Then
        m_oCreditCardMgr.SelCC.State.SetBits eMask.IsDirty
    End If
End Sub


Private Sub txtCardStreetNbr_LostFocus()
        m_oBrokenRules.Validate txtCardStreetNbr
        SetIsValid
        m_oCreditCardMgr.SelCC.StreetNbr = txtCardStreetNbr.text
End Sub


Private Sub txtCardZipCode_LostFocus()
        'This will fire a Change event and make the card dirty
        'Let's do this in the Update function before saving
        'txtCardZipCode.Text = UCase(txtCardZipCode.Text)
        m_oBrokenRules.Validate txtCardZipCode
        SetIsValid
        m_oCreditCardMgr.SelCC.ZipCode = txtCardZipCode.text
End Sub


Private Sub txtCardZipCode_Change()
    If Not m_bLoading Then
        m_oCreditCardMgr.SelCC.State.SetBits eMask.IsDirty
    End If
End Sub


Private Sub SetIsValid()
    If m_oBrokenRules.Count > 0 Then
        m_oCreditCardMgr.SelCC.State.ClearBits eMask.IsValid
    Else
        m_oCreditCardMgr.SelCC.State.SetBits eMask.IsValid
    End If
End Sub


Private Sub gdxCreditCards_UnboundDelete(ByVal RowIndex As Long, ByVal Bookmark As Variant)
    m_oCreditCardMgr.DeleteCreditCard
End Sub


Private Sub gdxCreditCards_SelectionChange()

    If SaveValidChgs Then
        AddUpdate
    Else
        m_oCreditCardMgr.CancelChanges
    End If

    m_oCreditCardMgr.SelectCreditCard gdxCreditCards.value(1) 'CCKey
    DisplayCreditCard
    
    'these lines of code are already in DisplayCreditCard procedure
'    m_oBrokenRules.Validate
'    Call SetIsValid
        
    If m_oCreditCardMgr.CreditCards.Count = 1 Then
        chkIsPrefCardNbr.Enabled = False
    Else
        chkIsPrefCardNbr.Enabled = True
    End If
End Sub


Private Sub gdxCreditCards_UnboundReadData _
(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As JSRowData)
    Dim lsCardTypeDescr As String
    lsCardTypeDescr = ""

    Values(1) = m_oCreditCardMgr.CreditCards.Item(RowIndex).key
    Values(2) = m_oCreditCardMgr.CreditCards.Item(RowIndex).Preferred
    
    m_oCreditCardMgr.GetCardTypeDescr m_oCreditCardMgr.CreditCards.Item(RowIndex).TypeKey, lsCardTypeDescr
    Values(3) = lsCardTypeDescr
    
    'allow the entire credit card to display if the user has rights to AR:ViewCCNo
    If HasRight(k_sRightARViewCCNo) Then
        Values(4) = m_oCreditCardMgr.CreditCards.Item(RowIndex).CardNo
    Else
        Values(4) = m_oCreditCardMgr.CreditCards.Item(RowIndex).MaskedCCNo
    End If

'***466 9/20/06 removed
'    Values(5) = m_oCreditCardMgr.CreditCards.Item(RowIndex).CID

    Values(6) = m_oCreditCardMgr.CreditCards.Item(RowIndex).ExpireDate
    Values(7) = m_oCreditCardMgr.CreditCards.Item(RowIndex).CardHolderName
    Values(8) = m_oCreditCardMgr.CreditCards.Item(RowIndex).StreetNbr
    Values(9) = m_oCreditCardMgr.CreditCards.Item(RowIndex).ZipCode
End Sub


'*********************************************************************
' Private procedures
'*********************************************************************

Private Sub AddUpdate()
    Dim aCreditCard As CreditCard
    Dim lsDupCCInfo As String

    'don't display message for a misc customer
    If m_oCustomer.HasAccount Then
    
        'check if CC# already exists with for a different CustKey
        m_oCreditCardMgr.DupCreditCardNo lsDupCCInfo
        If Len(lsDupCCInfo) > 0 Then
            lsDupCCInfo = MsgBox("Credit card number " & m_oCreditCardMgr.SelCC.MaskedCCNo & " already exists for: " & chr(10) & chr(13) & _
                lsDupCCInfo & chr(10) & chr(13) & _
                chr(10) & chr(13) & "Do you want to use this credit card number?", vbYesNo, "Credit Card")
            If lsDupCCInfo = vbNo Then
                cboCCType.SetFocus
                Exit Sub
            End If
        End If
    End If

    'Only one credit card can be set to preferred
    If chkIsPrefCardNbr.value = 1 Then
        For Each aCreditCard In m_oCreditCardMgr.CreditCards
            If aCreditCard.Preferred Then
                If aCreditCard.key <> m_oCreditCardMgr.SelCC.key Then
                    MsgBox cboCCType.text & " " & m_oCreditCardMgr.SelCC.MaskedCCNo & " is the preferred credit card."
                    aCreditCard.Preferred = False
                    aCreditCard.Update
                End If
            End If
        Next
    End If

    'Check & save state to a local variable because it get reset after the add/update
    Dim lsHolder As String

    If m_oCreditCardMgr.SelCC.IsNew Then lsHolder = "Add"
    
    m_oCreditCardMgr.AddUpdateCreditCard

    If lsHolder = "Add" Then
        'increment the count (and reload)
        gdxCreditCards.ItemCount = m_oCreditCardMgr.CreditCards.Count
        'select the just added credit card
        gdxCreditCards.RowSelected(gdxCreditCards.RowCount) = True
    Else
        'update the cells in the grid
        gdxCreditCards.RefreshRowIndex (gdxCreditCards.RowIndex(gdxCreditCards.Row))
    End If
End Sub


Private Function SaveValidChgs() As Boolean
    SaveValidChgs = False
    If m_oCreditCardMgr.SelCC.IsDirty And m_oCreditCardMgr.SelCC.IsValid Then

        If MsgBox("Do you want to save changes?", vbYesNo, "Save Changes") = vbYes Then
            SaveValidChgs = True
            Exit Function
        End If
    End If
End Function


Private Sub GridColAutoSize(gdx As GridEX)
    'autosize columns
    Dim i As Integer
    For i = 1 To gdx.Columns.Count
        gdx.Columns(i).AutoSize
    Next
End Sub


Private Sub DisplayCreditCard()
    
    m_bLoading = True
    m_oCreditCardMgr.LoadCardTypeList cboCCType
       
    'reset temp CardNo text box
    txtCardNoMask.text = vbNullString
    
    If m_oCreditCardMgr.SelCC.IsNew Then
        txtCardNoMask.Visible = False
        'set to the first item in the combo box
        cboCCType.ListIndex = 0
    Else
        txtCardNoMask.Visible = True
        'set mask for selected credit card type in txtCardNo
        txtCardNo.mask = m_oCreditCardMgr.GetCardTypeMask(m_oCreditCardMgr.SelCC.TypeKey)
        txtCardNo.PromptChar = "*"
        txtCardNo.FillChar = "0"
        txtCardNo.FillOnJump = True
        
        txtCardNo.text = m_oCreditCardMgr.SelCC.CardNo
        
        'display the entire # if the user has rights to AR:ViewCCNo
        If HasRight(k_sRightARViewCCNo) Then
            txtCardNoMask.text = m_oCreditCardMgr.SelCC.CardNo
        Else
            txtCardNoMask.text = m_oCreditCardMgr.SelCC.MaskedCCNo
        End If
     End If

    txtCardExp.mask = "##-####"
    txtCardExp.PromptChar = "x"
    txtCardExp.text = m_oCreditCardMgr.SelCC.ExpireDate
    
    lblKey.caption = m_oCreditCardMgr.SelCC.key
    
'***466 9/20/06 removed
'    txtCardCID.Text = m_oCreditCardMgr.SelCC.CID

    txtCardHolderName.text = m_oCreditCardMgr.SelCC.CardHolderName
    txtCardZipCode.text = m_oCreditCardMgr.SelCC.ZipCode
    txtCardStreetNbr.text = m_oCreditCardMgr.SelCC.StreetNbr
    
    If m_oCreditCardMgr.SelCC.Preferred = 0 Then
        chkIsPrefCardNbr.value = vbUnchecked
    Else
        chkIsPrefCardNbr.value = vbChecked
    End If
                
    If m_oCreditCardMgr.CreditCards.Count = 0 Then
        chkIsPrefCardNbr.value = vbChecked
    End If
                
    If m_oCreditCardMgr.SelCC.IsNew Then
        If Not m_oCustomer.HasAccount Then
        
            txtCardStreetNbr.text = m_oCustomer.ShipAddr.Addr1
            txtCardZipCode.text = m_oCustomer.ShipAddr.Zip
        Else
            txtCardStreetNbr.text = m_oCustomer.BillAddr.Addr1
            txtCardZipCode.text = m_oCustomer.BillAddr.Zip
        End If
        m_oCreditCardMgr.SelCC.StreetNbr = txtCardStreetNbr.text
        m_oCreditCardMgr.SelCC.ZipCode = txtCardZipCode.text
    End If
    
    'why are these done separately?
    m_oBrokenRules.Validate
    m_oBrokenRules.Validate chkIsPrefCardNbr
    
    SetIsValid

    SetButtons m_oCreditCardMgr.SelCC.State
    
    m_bLoading = False
    
End Sub


Private Sub SetZipCodeMask()
    If Not m_oCustomer.HasAccount Then
    
        txtCardZipCode.mask = PostalCodeMask(Trim$(m_oCustomer.ShipAddr.CountryID))
    Else
        txtCardZipCode.mask = PostalCodeMask(Trim$(m_oCustomer.BillAddr.CountryID))
    End If
End Sub


Private Function PostalCodeMask(CountryID As String) As String
    g_rstCountry.Filter = "CountryID = '" & Trim(CountryID) & "'"

    'It's possible that some addresses' countries are different from "USA", "CAN", or "MEX"  10/25/02 TX
    If Not g_rstCountry.EOF Then
        PostalCodeMask = Trim$(g_rstCountry("PostalCodeMask").value)
    Else
        PostalCodeMask = vbNullString
    End If
    g_rstCountry.Filter = adFilterNone
End Function


Private Sub SetButtons(NewState As BitMap)
        
    cmdDelete.Enabled = Not (NewState.TestBits(eMask.IsNew)) And _
        Not (NewState.TestBits(eMask.IsDirty))

    cmdUndo.Enabled = NewState.TestBits(eMask.IsDirty)
    
    cmdNew.Enabled = Not (NewState.TestBits(eMask.IsNew)) And _
        Not (NewState.TestBits(eMask.IsDirty))
    
    cmdAddUpdate.Enabled = NewState.TestBits(eMask.IsValid) And _
        NewState.TestBits(eMask.IsDirty)
    
    If cmdAddUpdate.Enabled Then
        If NewState.TestBits(eMask.IsNew) Then
            cmdAddUpdate.caption = "Add"
        Else
            cmdAddUpdate.caption = "Update"
        End If
    End If
    
    cmdOK.Enabled = NewState.TestBits(eMask.IsValid) And _
        Not (NewState.TestBits(eMask.IsDirty)) And _
        Not (NewState.TestBits(eMask.IsNew))
        
    If Not m_oCustomer.HasAccount Then
        cmdDelete.Enabled = False
        cmdNew.Enabled = False
        If Not (m_oCreditCardMgr.PrevCC Is Nothing) And m_oCreditCardMgr.CreditCards.Count > 0 Then
            'if the prev cc exist, we can't add a new one - we can only update
            If cmdAddUpdate.caption = "ADD" Then cmdAddUpdate.Enabled = False
        End If
    End If
    
End Sub


Private Sub LoadValidationRules()

    Dim oCtlWrapper As ControlWrapper
    Set m_oBrokenRules = New BrokenRules
    m_oBrokenRules.Form = Me
    
    Set oCtlWrapper = m_oBrokenRules.AddControl(txtCardHolderName, "Card Holder Name", True, True)
    oCtlWrapper.AddRuleRequired "", k_lClassAlways

    Set oCtlWrapper = m_oBrokenRules.AddControl(txtCardNo, "Card Number", True, True)
    oCtlWrapper.AddRuleCCNumber k_lClassAlways

    Set oCtlWrapper = m_oBrokenRules.AddControl(txtCardExp, "Expiration Date", True, True)
    oCtlWrapper.AddRuleRequired "", k_lClassAlways
    oCtlWrapper.AddRuleFutureDate Format(Now, "MMYYYY"), k_lClassAlways

    '***465 SMR 04-17-2006 - AmEx CCs can now be used for purchase by our resale customers***
    'Set oCtlWrapper = m_oBrokenRules.AddControl(cboCCType, "Card Type", True, True)
    'oCtlWrapper.AddRuleAmExCustType k_lClassAlways
    
    Set oCtlWrapper = m_oBrokenRules.AddControl(txtCardStreetNbr, "Street Number", True, True)
    oCtlWrapper.AddRuleRequired "", k_lClassAlways

    Set oCtlWrapper = m_oBrokenRules.AddControl(txtCardZipCode, "Card Zip Code", True, True)
    oCtlWrapper.AddRuleRequired "", k_lClassAlways
    SetZipCodeMask
    If txtCardZipCode.mask <> vbNullString Then oCtlWrapper.AddRuleZipCode txtCardZipCode.mask, k_lClassAlways
    
    'Set oCtlWrapper = m_oBrokenRules.AddControl(txtCardNo, "Credit Card Number", True, True)
    'oCtlWrapper.AddRuleDuplicateCC m_oCreditCardMgr.CreditCards, , m_oCreditCardMgr.SelCC, k_lClassAlways
    
    m_oBrokenRules.EnableClass k_lClassAlways, True
End Sub


