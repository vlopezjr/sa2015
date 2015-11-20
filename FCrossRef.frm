VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form FCrossRef 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cross Reference Searching Result"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7650
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4215
   ScaleWidth      =   7650
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   315
      Left            =   6600
      TabIndex        =   17
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton cmdCreateNew 
      Caption         =   "&New"
      Height          =   315
      Left            =   6600
      TabIndex        =   16
      Top             =   2700
      Width           =   855
   End
   Begin VB.Frame frmCreateNew 
      Caption         =   "Cross Reference Details"
      Height          =   1695
      Left            =   120
      TabIndex        =   5
      Top             =   2400
      Width           =   6375
      Begin VB.CheckBox chkIsDirect 
         Alignment       =   1  'Right Justify
         Caption         =   "Direct Replacement"
         Height          =   255
         Left            =   3000
         TabIndex        =   19
         Top             =   720
         Width           =   2175
      End
      Begin VB.CommandButton cmdLoad 
         Caption         =   "&Load"
         Enabled         =   0   'False
         Height          =   315
         Left            =   5280
         TabIndex        =   18
         Top             =   1140
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Height          =   315
         Left            =   5280
         TabIndex        =   14
         Top             =   300
         Width           =   975
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Enabled         =   0   'False
         Height          =   315
         Left            =   5280
         TabIndex        =   15
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox txtRefRemarks 
         Height          =   525
         Left            =   1320
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   13
         Top             =   1080
         Width           =   3855
      End
      Begin VB.TextBox txtCItemID 
         Height          =   315
         Left            =   3720
         MaxLength       =   30
         TabIndex        =   9
         Top             =   300
         Width           =   1455
      End
      Begin VB.ComboBox cboRefSource 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   690
         Width           =   1335
      End
      Begin VB.TextBox txtCRefItemID 
         Height          =   315
         Left            =   1320
         MaxLength       =   30
         TabIndex        =   7
         Top             =   300
         Width           =   1335
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Ref Remarks"
         Height          =   345
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Ref Source"
         Height          =   345
         Left            =   360
         TabIndex        =   10
         Top             =   690
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Part Nbr"
         Height          =   345
         Left            =   2760
         TabIndex        =   8
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblRefItemID 
         Alignment       =   1  'Right Justify
         Caption         =   "Ref Part Nbr"
         Height          =   345
         Left            =   240
         TabIndex        =   6
         Top             =   300
         Width           =   975
      End
   End
   Begin VB.Frame frmMaintenance 
      Caption         =   "Cross Reference Search"
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7335
      Begin GridEX20.GridEX gdxCrossRef 
         Height          =   1215
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   2143
         Version         =   "2.0"
         ScrollToolTips  =   -1  'True
         ShowToolTips    =   -1  'True
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         MethodHoldFields=   -1  'True
         AllowEdit       =   0   'False
         GroupByBoxVisible=   0   'False
         ColumnHeaderHeight=   285
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   12
         Column(1)       =   "FCrossRef.frx":0000
         Column(2)       =   "FCrossRef.frx":0124
         Column(3)       =   "FCrossRef.frx":0260
         Column(4)       =   "FCrossRef.frx":03A0
         Column(5)       =   "FCrossRef.frx":04DC
         Column(6)       =   "FCrossRef.frx":0640
         Column(7)       =   "FCrossRef.frx":0798
         Column(8)       =   "FCrossRef.frx":08FC
         Column(9)       =   "FCrossRef.frx":0A20
         Column(10)      =   "FCrossRef.frx":0B6C
         Column(11)      =   "FCrossRef.frx":0C78
         Column(12)      =   "FCrossRef.frx":0DB8
         FormatStylesCount=   6
         FormatStyle(1)  =   "FCrossRef.frx":0F04
         FormatStyle(2)  =   "FCrossRef.frx":103C
         FormatStyle(3)  =   "FCrossRef.frx":10EC
         FormatStyle(4)  =   "FCrossRef.frx":11A0
         FormatStyle(5)  =   "FCrossRef.frx":1278
         FormatStyle(6)  =   "FCrossRef.frx":1330
         ImageCount      =   0
         PrinterProperties=   "FCrossRef.frx":1410
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "Fi&nd"
         Height          =   315
         Left            =   4320
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtMItemID 
         Height          =   315
         Left            =   2160
         MaxLength       =   30
         TabIndex        =   2
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Part Number"
         Height          =   315
         Left            =   720
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
   End
End
Attribute VB_Name = "FCrossRef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_lWindowID As Long

Private WithEvents m_gdxCrossRef As GridEXWrapper
Attribute m_gdxCrossRef.VB_VarHelpID = -1

Private m_bCreateNew As Boolean
Private m_bItemIDEnable As Boolean
Private m_bRefItemIDEnable As Boolean
Private m_bRefRemark As Boolean
Private m_bRefSource As Boolean
Private m_bDelete As Boolean
Private m_bLoadCrossRef As Boolean
Private m_bClose As Boolean
Private m_bLoad As Boolean


Public Property Get WindowID() As Long
    WindowID = m_lWindowID
End Property

Public Property Let WindowID(ByVal lNewValue As Long)
    m_lWindowID = lNewValue
End Property


Public Sub SetCaption(ByRef i_sTitle As String)
    Me.caption = i_sTitle
    MDIMain.UpdateCaption Me
End Sub


Public Sub DoShowHelp()
    ShowHelp "FCrossRef", True
End Sub


' Form Events

Private Sub Form_Activate()
    MDIMain.UpdateWindowListSelection Me
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MDIMain.GlobalKeyDownProcessing KeyCode, Shift
End Sub


Private Sub Form_Load()
    If Not m_bLoadCrossRef Then SetCaption "Cross Reference Maintenance"
    Set m_gdxCrossRef = New GridEXWrapper
    m_gdxCrossRef.Grid = gdxCrossRef
    
    Dim rst As ADODB.Recordset
    m_bCreateNew = True
    
    Set rst = LoadDiscRst("Select * from tcpXRefType")
    Helpers.LoadCombo cboRefSource, rst, "XRefTypeDesc", "XRefTypeKey"
End Sub


Private Sub Form_Unload(Cancel As Integer)
    '3/31/05 http://www.devx.com/vb2themax/Tip/18461
    Set m_gdxCrossRef = Nothing
    
    MDIMain.UnloadTool m_lWindowID
End Sub


' Control Event Handlers

Private Sub cboRefSource_LostFocus()
    If Not m_bCreateNew Then
        m_bRefSource = (cboRefSource.ItemData(cboRefSource.ListIndex) <> m_gdxCrossRef.value("XRefSourceKey"))
        cmdSave.Enabled = SaveEnable
    End If
End Sub


Private Sub cmdClose_Click()
    m_bLoad = False
    m_bClose = True
    Unload Me
End Sub


Private Sub cmdCreateNew_Click()
    m_bCreateNew = True
    txtCRefItemID.text = ""
    TryToSetFocus txtCRefItemID
    txtCItemID.text = ""
    txtRefRemarks.text = ""
    SetComboByKey cboRefSource, 1
    cmdSave.Enabled = False
    cmdDelete.Enabled = False
    chkIsDirect.value = vbUnchecked
End Sub


Private Sub cmdDelete_Click()
    Dim cmd As ADODB.Command
    
     If vbYes = msg("Are you sure that you want to delete this Cross Reference record?", _
                    vbExclamation + vbYesNo, "Delete Cross Ref") Then
           Set cmd = CreateCommandSP("cpimDeleteXRef")
           cmd.Parameters("@XRefKey") = m_gdxCrossRef.value("XRefKey")
           cmd.Execute
           m_bDelete = True
           cmdFind_Click
     Else
        Exit Sub
     End If
End Sub


Private Sub cmdFind_Click()
    Dim rst As ADODB.Recordset
    cmdSave.Enabled = False
    cmdLoad.Visible = False
    m_bLoadCrossRef = False
    
    Set rst = CallSP("cpimGetXRefMaintenance", "@ItemID", txtMItemID.text)
    
    ' 10/08/03 AVH Hmmm, m_bCreateNew shouldn't be set to False, nor anything at this point
    'm_bCreateNew = False
    
    With gdxCrossRef
        .HoldFields
        .HoldSortSettings = True
        Set .ADORecordset = rst
    End With
    If rst.RecordCount = 0 Then
        If Not m_bDelete Then msg "No record matches this requests."
        txtMItemID.SelStart = 0
        txtMItemID.SelLength = Len(txtMItemID.text)
        TryToSetFocus txtMItemID
        ClearRefDetail
    Else
        gdxCrossRef.Row = 1
        DisplayCrossRefDetail
    End If
    m_bDelete = False
End Sub


Private Sub cmdLoad_Click()
    m_bLoad = True
    Me.Hide
End Sub


Private Sub cmdSave_Click()
    On Error GoTo ErrorHandler

    Dim cmd As ADODB.Command
    If Trim(txtCItemID.text) = "" Then
        msg "Please enter Part Nbr first before saving.", vbOKOnly + vbExclamation, "Enter Part Nbr"
        Exit Sub
    End If
    If Trim(txtCRefItemID.text) = "" Then
        msg "Please enter Ref Part Nbr first before saving.", vbOKOnly + vbExclamation, "Enter Ref Part Nbr"
        Exit Sub
    End If
    If m_bCreateNew Then
        If vbYes = msg("Are you sure that you want to create this new Cross Reference record?", _
                    vbExclamation + vbYesNo, "Create New Cross Ref") Then

                    Set cmd = CreateCommandSP("cpimAddXRef")
                    With cmd
                        .Parameters("@ItemID1").value = Trim(txtCItemID.text)
                        .Parameters("@ItemID2").value = Trim(txtCRefItemID.text)
                        .Parameters("@XRefTypeKey").value = cboRefSource.ItemData(cboRefSource.ListIndex)
                        .Parameters("@IsDirect").value = chkIsDirect.value
                        .Parameters("@XRefRemark").value = Trim(txtRefRemarks.text)
                        .Parameters("@UpdatedBy").value = GetUserID
                        .Execute
                    End With
                    ClearRefDetail
                    TryToSetFocus txtMItemID
        Else
            Exit Sub
        End If
    Else
        If vbYes = msg("Are you sure that you want to update this Cross Reference record?", _
                    vbExclamation + vbYesNo, "Update Cross Ref Record") Then

                Set cmd = CreateCommandSP("cpimUpdateXRef")
                With cmd
                    .Parameters("@XRefKey").value = m_gdxCrossRef.value("XRefKey")
                    .Parameters("@ItemID1").value = Trim(txtCItemID.text)
                    .Parameters("@ItemID2").value = Trim(txtCRefItemID.text)
                    .Parameters("@XRefTypeKey").value = cboRefSource.ItemData(cboRefSource.ListIndex)
                    .Parameters("@IsDirect").value = chkIsDirect.value
                    .Parameters("@XRefRemark").value = Trim(txtRefRemarks.text)
                    .Parameters("@UpdatedBy").value = GetUserID
                    .Execute
                End With
                m_bDelete = True
                cmdFind_Click
        Else
            Exit Sub
        End If
    End If
    Exit Sub
ErrorHandler:
    On Error Resume Next
    ClearWaitCursor
    msg Err.Description, vbOKOnly + vbCritical, Err.Source
End Sub

Private Sub gdxCrossRef_click()
    DisplayCrossRefDetail
End Sub


Private Sub gdxCrossRef_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
        DisplayCrossRefDetail
    End If
End Sub


Private Sub m_gdxCrossRef_RowChosen()
    DisplayCrossRefDetail
    If m_bLoadCrossRef Then
        cmdLoad_Click
    End If
End Sub


Private Sub txtCItemID_LostFocus()
    If m_bClose Then Exit Sub
    
    If Trim(txtCItemID.text) = "" Then Exit Sub
    If m_bCreateNew Then
        cmdSave.Enabled = (Trim(txtCItemID.text) <> "" And Trim(txtCRefItemID.text) <> "")
    Else
        If (Trim(txtCItemID.text) <> m_gdxCrossRef.value("PartNbr")) Then
                m_bItemIDEnable = True
        Else
            m_bItemIDEnable = False
        End If
        cmdSave.Enabled = SaveEnable
    End If
End Sub


Private Sub txtCRefItemID_LostFocus()
    If Trim(txtCRefItemID.text) = "" Then Exit Sub

    If m_bCreateNew Then
        cmdSave.Enabled = (Trim(txtCItemID.text) <> "" And Trim(txtCRefItemID.text) <> "")
    Else
        If (Trim(txtCRefItemID.text) <> m_gdxCrossRef.value("RefPartNbr")) Then
                m_bRefItemIDEnable = True
        Else
            m_bRefItemIDEnable = False
        End If
        cmdSave.Enabled = SaveEnable
    End If
End Sub


Private Sub txtMItemID_KeyDown(KeyCode As Integer, Shift As Integer)
    If Len(txtMItemID.text) > 0 Then
        If KeyCode = vbKeyReturn Then
            cmdFind_Click
        End If
    End If
End Sub


Private Sub txtRefRemarks_KeyPress(KeyAscii As Integer)
    cmdSave.Enabled = (Trim(txtCItemID.text) <> "" And Trim(txtCRefItemID.text) <> "")
End Sub


Private Sub txtRefRemarks_LostFocus()
    Dim sRemarks As String
    If Not m_bCreateNew Then
        If Not IsNull(m_gdxCrossRef.value("XRefRemarks")) Then
            sRemarks = m_gdxCrossRef.value("XRefRemarks")
        End If
        m_bRefRemark = (Trim(txtRefRemarks.text) <> sRemarks)
        cmdSave.Enabled = SaveEnable
    End If
End Sub


' Private Methods

Private Sub ChooseFromGrid(i_rst As ADODB.Recordset, o_lItemKey As Long, o_eItemType As ItemTypeCode, bConfirm As Boolean)
    cmdLoad.Visible = True
    cmdLoad.Enabled = True
    cmdSave.Enabled = False
    cmdDelete.Enabled = False
    cmdCreateNew.Enabled = False
    cmdFind.Enabled = False
    Dim sMsg As String
    
    With gdxCrossRef
        Dim i As Long
        .HoldFields
        .HoldSortSettings = True
        Set .ADORecordset = i_rst
        For i = 1 To .Columns.Count
            .Columns(i).AutoSize
        Next
    End With
    DisplayCrossRefDetail
    Me.Show vbModal
    If m_bLoad Then
        If bConfirm Then
            If vbYes = msg("Part Number " & m_gdxCrossRef.value("RefPartNbr") & " is an obsolete or discontinued item. " _
                        & "It is cross referenced to replacement Part Number " & m_gdxCrossRef.value("PartNbr") & "." & vbCrLf & vbCrLf _
                        & "Would you like to load " & m_gdxCrossRef.value("PartNbr") & "?", vbExclamation + vbYesNo, "Loading Cross Referenced Part?") Then
                    o_eItemType = ConvertSageItemType(m_gdxCrossRef.value("ItemType"))
                    o_lItemKey = m_gdxCrossRef.value("ItemKey")
            End If
        Else
            If Trim(m_gdxCrossRef.value("XRefRemarks")) <> "" Then
                sMsg = vbCrLf & vbCrLf & Trim(m_gdxCrossRef.value("XRefRemarks"))
            End If
            If vbYes = msg("According to " & Trim(m_gdxCrossRef.value("UpdateUserID")) _
                            & " on " & m_gdxCrossRef.value("UpdateDate") & ", based on " _
                            & Trim(m_gdxCrossRef.value("XRefSourceDescr")) & " information," & vbCrLf _
                            & "Part Number " & Trim(m_gdxCrossRef.value("RefPartNbr")) & " is cross " _
                            & "referenced to our Part Number " & Trim(m_gdxCrossRef.value("PartNbr")) & "." _
                            & sMsg & vbCrLf & vbCrLf _
                            & "Would you like to load " & Trim(m_gdxCrossRef.value("PartNbr")) & "?", vbExclamation + vbYesNo, _
                            "Loading Cross Referenced Part?") Then
                            o_eItemType = ConvertSageItemType(m_gdxCrossRef.value("ItemType"))
                            o_lItemKey = m_gdxCrossRef.value("ItemKey")
            Else
                    o_lItemKey = 0
            End If
        End If
    End If
End Sub


Private Sub ClearRefDetail()
    cmdSave.Enabled = False
    cmdDelete.Enabled = False
    
    ' 10/08/03 AVH Hmmm, m_bCreateNew shouldn't be set to False.
    '           Code in this routine blank out controls to set for new entry, set m_bCreateNew to true
    'm_bCreateNew = False
    m_bCreateNew = True
    
    m_bItemIDEnable = False
    m_bRefItemIDEnable = False
    m_bRefRemark = False
    m_bRefSource = False
    txtCRefItemID.text = ""
    txtCItemID.text = ""
    SetComboByKey cboRefSource, 1
    txtRefRemarks.text = ""
    chkIsDirect.value = vbUnchecked
End Sub


Private Sub DisplayCrossRefDetail()
    If Not m_bLoadCrossRef Then
        cmdSave.Enabled = False
        cmdDelete.Enabled = True
        m_bCreateNew = False
        m_bItemIDEnable = False
        m_bRefItemIDEnable = False
        m_bRefRemark = False
        m_bRefSource = False
    End If
    txtCRefItemID.text = m_gdxCrossRef.value("RefPartNbr")
    txtCItemID.text = m_gdxCrossRef.value("PartNbr")
    SetComboByKey cboRefSource, m_gdxCrossRef.value("XRefSourceKey")
    If IsNull(m_gdxCrossRef.value("XRefRemarks")) Then
        txtRefRemarks.text = ""
    Else
        txtRefRemarks.text = m_gdxCrossRef.value("XRefRemarks")
    End If
    chkIsDirect.value = Abs(m_gdxCrossRef.value("IsDirect"))
End Sub


Private Function SaveEnable() As Boolean
    SaveEnable = m_bItemIDEnable Or m_bRefItemIDEnable Or m_bRefRemark Or m_bRefSource
End Function


