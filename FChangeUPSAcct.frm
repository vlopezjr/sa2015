VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form FChangeUPSAcct 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7740
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   7740
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   405
      Left            =   6600
      TabIndex        =   2
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&Select"
      Height          =   405
      Left            =   5400
      TabIndex        =   1
      Top             =   3000
      Width           =   975
   End
   Begin GridEX20.GridEX gdxChangeUPS 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   4895
      Version         =   "2.0"
      ScrollToolTips  =   -1  'True
      ShowToolTips    =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      MethodHoldFields=   -1  'True
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      ColumnHeaderHeight=   285
      ColumnsCount    =   9
      Column(1)       =   "FChangeUPSAcct.frx":0000
      Column(2)       =   "FChangeUPSAcct.frx":0150
      Column(3)       =   "FChangeUPSAcct.frx":0290
      Column(4)       =   "FChangeUPSAcct.frx":04E4
      Column(5)       =   "FChangeUPSAcct.frx":0634
      Column(6)       =   "FChangeUPSAcct.frx":0784
      Column(7)       =   "FChangeUPSAcct.frx":08D4
      Column(8)       =   "FChangeUPSAcct.frx":0A0C
      Column(9)       =   "FChangeUPSAcct.frx":0B4C
      FormatStylesCount=   6
      FormatStyle(1)  =   "FChangeUPSAcct.frx":0CA8
      FormatStyle(2)  =   "FChangeUPSAcct.frx":0DE0
      FormatStyle(3)  =   "FChangeUPSAcct.frx":0E90
      FormatStyle(4)  =   "FChangeUPSAcct.frx":0F44
      FormatStyle(5)  =   "FChangeUPSAcct.frx":101C
      FormatStyle(6)  =   "FChangeUPSAcct.frx":10D4
      ImageCount      =   0
      PrinterProperties=   "FChangeUPSAcct.frx":11B4
   End
End
Attribute VB_Name = "FChangeUPSAcct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents m_gw As GridEXWrapper
Attribute m_gw.VB_VarHelpID = -1
Private m_bLoad As Boolean


Public Function SearchUPSAcct(sCustomerID As String, lCustKey As Long, ByVal sUPSAcct As String, Optional ByVal lAddrKey As Long) As String
    Dim sSQL As String
    Dim orst As ADODB.Recordset
    
    '09/09/2003 AVH Add addr type column to UPS Account screens
    sSQL = "Select distinct tciAddress.AddrKey, tciAddress.AddrName, tciAddress.AddrLine1, " _
                & "tciAddress.AddrLine2, rtrim(tciAddress.City) as City, tciAddress.StateID, " _
                & " tciAddress.PostalCode, tcpUPSAcct.UPSAcct, tarCustomer.DfltShipToAddrKey " _
                & " , AddrType = CASE " _
                & " WHEN tciAddress.AddrKey = tarCustomer.DfltBillToAddrKey AND tciAddress.AddrKey = tarCustomer.DfltShipToAddrKey THEN 'B&S' " _
                & " WHEN tciAddress.AddrKey = tarCustomer.DfltBillToAddrKey THEN 'Bill' " _
                & " WHEN tciAddress.AddrKey = tarCustomer.DfltShipToAddrKey THEN 'Ship' " _
                & " Else ' ' " _
                & " End " _
                & " from tciAddress inner join " _
                & "tarCustAddr on tciAddress.Addrkey = tarCustAddr.AddrKey inner join " _
                & "tarCustomer on tarCustomer.CustKey = tarCustAddr.CustKey " _
                & "inner join tcpUPSAcct on tcpUPSAcct.CustAddrKey = tciAddress.AddrKey " _
                & "where tcpUPSAcct.UPSAcct <> '' and tarCustAddr.CustKey = " & lCustKey _
                & " and tcpUPSAcct.UPSAcct <> '" & sUPSAcct & "'"
                
    '09/09/2003 AVH Load only headquarter or this particular branch's UPS account
    If lAddrKey > 0 Then
                sSQL = sSQL & " and ((tciAddress.AddrKey = tarCustomer.DfltShipToAddrKey) " _
                & " or (tciAddress.AddrKey = " & lAddrKey & "))"
    End If
    
    Set orst = LoadDiscRst(sSQL)
    
    If orst.RecordCount = 0 Then
        Msg "Sorry. No additional UPS account available for Bill Recipient.", _
                vbOKOnly + vbExclamation, "Searching Result"
        SearchUPSAcct = ""
    Else
        Me.Caption = "Change UPS account for " & sCustomerID
        SearchUPSAcct = ChooseFromGrid(orst, "UPSAcct")
    End If
    
    Unload Me
End Function


Private Function ChooseFromGrid(i_rst As ADODB.Recordset, i_sKeyField As String) As String
    If i_rst Is Nothing Then Exit Function
    
    With gdxChangeUPS
        .HoldFields
        .HoldSortSettings = True
        Set .ADORecordset = i_rst
    End With
    
    Me.Show vbModal
    If m_bLoad Then
        ChooseFromGrid = m_gw.Value(i_sKeyField)
        'lAddrKey = m_gw.Value("AddrKey")
    End If
    MDIMain.DoRefresh
End Function


Private Sub cmdCancel_Click()
    m_bLoad = False
    Me.Hide
End Sub

Private Sub cmdSelect_Click()
    With gdxChangeUPS
        If .RowIndex(.Row) <= 0 Then
            Msg "Please select the desired UPS Acct from the grid.", , "No UPS Acct Selected"
            Exit Sub
        End If
    End With
    m_bLoad = True
    Me.Hide
End Sub


Private Sub Form_Activate()
     With gdxChangeUPS
        TryToSetFocus gdxChangeUPS
        If .RowCount >= 1 Then
            .Row = 1
        End If
    End With
    
    Set m_gw = New GridEXWrapper
    m_gw.Grid = gdxChangeUPS
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set m_gw = Nothing
End Sub
    
    
Private Sub m_gw_RowChosen()
    cmdSelect_Click
End Sub
