VERSION 5.00
Object = "{0FA91D91-3062-44DB-B896-91406D28F92A}#54.0#0"; "SOTACalendar.ocx"
Begin VB.Form FFixPO 
   ClientHeight    =   1845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3930
   LinkTopic       =   "Form1"
   ScaleHeight     =   1845
   ScaleWidth      =   3930
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   312
      Left            =   2820
      TabIndex        =   1
      Top             =   1440
      Width           =   972
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   312
      Left            =   1800
      TabIndex        =   0
      Top             =   1440
      Width           =   972
   End
   Begin VB.ComboBox cboBuyers 
      Height          =   315
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   180
      Width           =   1452
   End
   Begin SOTACalendarControl.SOTACalendar calReqDate 
      Height          =   288
      Left            =   1320
      TabIndex        =   4
      Top             =   900
      Width           =   1452
      _ExtentX        =   2566
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaskedText      =   "  /  /    "
      Text            =   "  /  /    "
   End
   Begin SOTACalendarControl.SOTACalendar calPODate 
      Height          =   288
      Left            =   1320
      TabIndex        =   3
      Top             =   540
      Width           =   1452
      _ExtentX        =   2566
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaskedText      =   "  /  /    "
      Text            =   "  /  /    "
   End
   Begin VB.Label Label3 
      Caption         =   "Buyer"
      Height          =   192
      Left            =   120
      TabIndex        =   7
      Top             =   240
      Width           =   972
   End
   Begin VB.Label Label2 
      Caption         =   "Request Date"
      Height          =   192
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   1092
   End
   Begin VB.Label Label1 
      Caption         =   "PO Date"
      Height          =   192
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   852
   End
End
Attribute VB_Name = "FFixPO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_lPOKey As Long
Private m_bFixed As Boolean


Public Property Get Fixed() As Boolean
    Fixed = m_bFixed
End Property


Public Sub FixPO(POID As String)
    Dim oRst As Recordset
    Dim sSQL As String
   
    m_bFixed = False
    
    'get the POKey using the POID
    sSQL = "SELECT POKey FROM tpoPurchOrder WHERE tranno LIKE'%" & POID & "'"
    Set oRst = LoadDiscRst(sSQL)
    m_lPOKey = oRst.Fields("POKey").value   'cache this
    Set oRst = Nothing

    'load the Buyer combobox
    Helpers.LoadCombo cboBuyers, g_rstBuyers, "BuyerID", "BuyerKey", , 1
    SetComboByKey cboBuyers, UserNameToBuyerKey(GetUserName)

    'initialize the remaining controls
    Me.Caption = "Fix Drop Ship PO " & POID
    calPODate.value = Now
    calReqDate.value = Now
    
    'show the form modally
    Me.Show vbModal
    
    'on returning
    Unload Me
End Sub


Private Sub cmdCancel_Click()
    m_bFixed = False
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    Dim oCmd As ADODB.Command
    Dim oRst As ADODB.Recordset
    Dim oRstComment As ADODB.Recordset
    Dim sSQL As String
    Dim sCommandText As String
    Dim dUnitCost As Double
    Dim dQtyOrd As Double
    Dim dExtAmt As Double
    Dim dPurchAmt As Double
    Dim sComment As String
    Dim lShipMethKey As Long

    If cboBuyers.List(cboBuyers.ListIndex) = "<none>" Then
        msg "You must select a buyer.", vbExclamation
        Exit Sub
    End If
    
    SetWaitCursor True

    'get data for all PO SPO line items
    'included tpoPOLineDist.ShipMethKey to support persisting UserFld2
    sSQL = "SELECT tpoPOLine.POLineKey, tcpSOLine.SOLineKey, tcpSOLine.OPKey, tcpSOLine.Cost, " & _
    "tpoPOLineDist.QtyOrd, tpoPOLine.UnitCost, tpoPOLine.ExtAmt, tpoPOLineDist.ShipMethKey " & _
    "FROM tpoPOLine INNER JOIN tsoSOLine ON tpoPOLine.POLineKey = tsoSOLine.POLineKey " & _
    "INNER JOIN tcpSOLine ON tcpSOLine.SOLineKey = tsoSOLine.SOLineKey " & _
    "INNER JOIN tpoPOLineDist ON tpoPOLine.POLineKey = tpoPOLineDist.POLineKey " & _
    "WHERE tpoPOLine.POKey = " & m_lPOKey
    
    On Error GoTo EH
    
    g_DB.Connection.BeginTrans
        
    Set oRst = LoadRst(sSQL)

'For each PO line item:
'NOTE: I expect stock items as well as SPOs.
'For stock items the UnitCost and ExtAmt are already in the POLine and in the Order totals.
    
    Do While Not oRst.EOF
        
        If oRst.Fields("UnitCost").value = 0 Then
            'assume this is a SPO
            'build an ExtComment for the Sage PO
                            
            sComment = ""
            If oRst.Fields("SOLineKey").value > 0 Then
                Set oCmd = CreateCommandSP("spcpcPOGetSPODetail")
                oCmd.Parameters("@_iSOLineKey").value = oRst.Fields("SOLineKey").value
                Set oRstComment = oCmd.Execute
                If Not oRstComment.EOF Then
                    If Len(oRstComment.Fields("ModelNbr").value) > 0 Then
                        sComment = sComment & "Model # " & Trim(oRstComment.Fields("ModelNbr").value)
                    End If
                    If Len(oRstComment.Fields("SerialNbr").value) > 0 Then
                        sComment = sComment & "; Serial # " & Trim(oRstComment.Fields("SerialNbr").value)
                    End If
                    sComment = sComment & "; OP " & oRstComment.Fields("OPKey").value
                    sComment = sComment & "; SO " & Trim(oRstComment.Fields("TranKey").value)
                End If
                Set oRstComment = Nothing
            End If
        
            dUnitCost = CDbl(oRst.Fields("Cost"))
            dQtyOrd = CDbl(oRst.Fields("QtyOrd"))
            dExtAmt = dUnitCost * dQtyOrd

            sCommandText = "UPDATE tpoPOLine SET " & _
                    "UnitCost=" & dUnitCost & ", " & _
                    "ExtAmt=" & dExtAmt & ", " & _
                    "ExtCmnt = '" & PrepSQLText(Left(sComment, 254)) & "' " & _
                    "WHERE POLineKey=" & oRst.Fields("POLineKey")
                    
            Set oCmd = CreateCommandSP(sCommandText, adCmdText)
            oCmd.Execute
        Else
        'else it's a stock item
            dUnitCost = CDbl(oRst.Fields("UnitCost").value)
            dExtAmt = CDbl(oRst.Fields("ExtAmt").value)
        End If

        'cache ShipMethKey for use below (all line items should have the same key) 9/23/03 LR
        lShipMethKey = oRst.Fields("ShipMethKey").value
        
        'Set oCmd = New ADODB.Command
        sCommandText = "UPDATE tpoPOLineDist SET " & _
                "ExtAmt=" & dExtAmt & ", " & _
                "OrigOrdered=" & dQtyOrd & ", " & _
                "RequestDate='" & calReqDate.value & "', " & _
                "PromiseDate='" & calReqDate.value & "' " & _
                "WHERE POLineKey=" & oRst.Fields("POLineKey")
        Set oCmd = CreateCommandSP(sCommandText, adCmdText)
        oCmd.Execute

        'keep a running total amount for the order
        dPurchAmt = dPurchAmt + dExtAmt
        
        oRst.MoveNext
    Loop
    
'9/23/03 LR assign the BuyerKey and the ShipMethKey to UserFld1 and UserFld2 of the tpoPurchOrder
    
    'Set oCmd = New ADODB.Command
    sCommandText = "UPDATE tpoPurchOrder SET " & _
            "BuyerKey=" & cboBuyers.ItemData(cboBuyers.ListIndex) & ", " & _
            "UserFld1=" & cboBuyers.ItemData(cboBuyers.ListIndex) & ", " & _
            "UserFld2=" & lShipMethKey & ", " & _
            "TranDate='" & calPODate.value & "', " & _
            "DfltRequestDate='" & calReqDate.value & "', " & _
            "OpenAmt=" & dPurchAmt & ", " & _
            "OpenAmtHC=" & dPurchAmt & ", " & _
            "PurchAmt=" & dPurchAmt & ", " & _
            "PurchAmtHC=" & dPurchAmt & ", " & _
            "TranAmt=" & dPurchAmt & ", " & _
            "TranAmtHC=" & dPurchAmt & " " & _
            "WHERE POKey=" & m_lPOKey
    Set oCmd = CreateCommandSP(sCommandText, adCmdText)
    oCmd.Execute

    g_DB.Connection.CommitTrans

ExitSub:
    SetWaitCursor False

    m_bFixed = True
    Me.Hide
    
    Exit Sub
EH:
    g_DB.Connection.RollbackTrans
    msg "Error fixing dropship PO" & vbCrLf & "(" & Err.Number & ") " & Err.Description & vbCrLf & "Transaction rolled back"
    GoTo ExitSub
    
End Sub

