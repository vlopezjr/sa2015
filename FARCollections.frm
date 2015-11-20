VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FARCollections 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5190
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   5190
   Begin VB.CommandButton cmdCreateSheets 
      Caption         =   "Create Call Sheets"
      Height          =   312
      Left            =   420
      TabIndex        =   16
      Top             =   480
      Width           =   1632
   End
   Begin VB.Frame Frame2 
      Caption         =   "Create Letters"
      Height          =   2955
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   4692
      Begin VB.CommandButton cmdMark 
         Caption         =   "Mark"
         Height          =   312
         Left            =   420
         TabIndex        =   7
         Top             =   2460
         Width           =   972
      End
      Begin VB.CommandButton cmdMailMerge 
         Caption         =   "Start Word"
         Height          =   312
         Left            =   420
         TabIndex        =   6
         Top             =   1980
         Width           =   975
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "Browse..."
         Height          =   312
         Left            =   3360
         TabIndex        =   5
         Top             =   1560
         Width           =   972
      End
      Begin VB.CommandButton cmdCreate 
         Caption         =   "Create"
         Height          =   312
         Left            =   420
         TabIndex        =   4
         Top             =   1500
         Width           =   972
      End
      Begin VB.OptionButton optOver60 
         Caption         =   "over60"
         Height          =   192
         Left            =   420
         TabIndex        =   3
         Top             =   900
         Width           =   972
      End
      Begin VB.OptionButton optOver45 
         Caption         =   "over 45"
         Height          =   192
         Left            =   420
         TabIndex        =   2
         Top             =   660
         Value           =   -1  'True
         Width           =   972
      End
      Begin VB.OptionButton optOver90 
         Caption         =   "over 90"
         Height          =   192
         Left            =   420
         TabIndex        =   1
         Top             =   1140
         Width           =   972
      End
      Begin VB.Label lblFileName 
         Height          =   252
         Left            =   1560
         TabIndex        =   15
         Top             =   360
         Width           =   2892
      End
      Begin VB.Label Label9 
         Caption         =   "1.   Select type"
         Height          =   192
         Left            =   180
         TabIndex        =   14
         Top             =   360
         Width           =   2352
      End
      Begin VB.Label Label8 
         Caption         =   "customers"
         Height          =   252
         Left            =   1500
         TabIndex        =   13
         Top             =   2520
         Width           =   1152
      End
      Begin VB.Label Label6 
         Caption         =   "4."
         Height          =   252
         Left            =   180
         TabIndex        =   12
         Top             =   2520
         Width           =   252
      End
      Begin VB.Label Label1 
         Caption         =   "3."
         Height          =   252
         Left            =   180
         TabIndex        =   11
         Top             =   2040
         Width           =   252
      End
      Begin VB.Label Label5 
         Caption         =   "and do MailMerge"
         Height          =   252
         Left            =   1500
         TabIndex        =   10
         Top             =   2040
         Width           =   1632
      End
      Begin VB.Label Label4 
         Caption         =   "2."
         Height          =   252
         Left            =   180
         TabIndex        =   9
         Top             =   1560
         Width           =   252
      End
      Begin VB.Label Label3 
         Caption         =   "MailMerge Data Source"
         Height          =   252
         Left            =   1500
         TabIndex        =   8
         Top             =   1560
         Width           =   1812
      End
   End
   Begin MSComDlg.CommonDialog dlgOpen 
      Left            =   4680
      Top             =   120
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   17
      Top             =   4560
      Width           =   5190
      _ExtentX        =   9155
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   0
      TabIndex        =   18
      Top             =   4200
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
End
Attribute VB_Name = "FARCollections"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Created by:   Len Russell

'***** IMPORTANT NOTE *********************************
'
' The MM DLL relies on the registry entries set by OA's
' logon function to decide which database to connect to.
' Make sure to configure properly before running this.
'
'******************************************************

'Private Const ksProdFilePath = "\\cpwebpro\orderpad\Collections\"
'Private Const ksDevFilePath = "\\Mpkfnp2\ca\IT\AR Programs\Collections\"

'Private m_sFilePath As String

'Private collectors(1 To 5) As String
'Private collectors() As String
   
Private m_sFileName As String
Private m_sLetterPath As String

Private m_lWindowID As Long


Public Property Get WindowID() As Long
    WindowID = m_lWindowID
End Property

Public Property Let WindowID(ByVal lNewValue As Long)
    m_lWindowID = lNewValue
End Property

Public Sub SetCaption(ByRef i_sTitle As String)
    Me.Caption = i_sTitle
    MDIMain.UpdateCaption Me
End Sub

Private Sub Form_Activate()
    MDIMain.UpdateWindowListSelection Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MDIMain.GlobalKeyDownProcessing KeyCode, Shift
End Sub


Private Sub Form_Load()

    'Dim collectorCount As Integer
    'ReDim collectors(0)
    'collectorCount = ParseString(collectors(), g_sCollectors, ",")
    
    StatusBar1.Panels(1).width = Me.width / 2
    
    'm_sFilePath = g_CollectionsPath
    
    SetCaption "AR Collections"
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    MDIMain.UnloadTool m_lWindowID
End Sub


Private Sub cmdCreateSheets_Click()
    Dim ocmd As ADODB.Command
    Dim orst As ADODB.Recordset
    Dim oExcel As Object
    Dim oBook As Excel.Workbook
    Dim oSheets As Excel.Worksheets
    Dim oSheet As Excel.Worksheet
    Dim oRange As Excel.Range
    Dim i As Integer
    Dim sDate As String

    On Error GoTo EH
    
    Screen.MousePointer = vbHourglass

    StatusBar1.Panels(1).text = "Open Excel Workbook"
    Set oExcel = New Excel.Application
    Set oBook = oExcel.Workbooks.Add    'add additional sheets (the assumption is you need more)
    'For i = 1 To UBound(collectors) - oBook.Worksheets.Count
    
    g_rstCollectors.Filter = "UserID <> 'InCollection' AND UserID <> 'WriteOff'"
    
    For i = 1 To g_rstCollectors.RecordCount - oBook.Worksheets.Count
        oBook.Worksheets.Add
    Next i
    
    Set ocmd = New ADODB.Command
    With ocmd
        .ActiveConnection = g_DB.Connection
        .CommandType = adCmdStoredProc
        .CommandText = "spcpcGetCollectionsByAgent"
    End With
    
    'setup a sheet for each collector
    
    For i = 1 To g_rstCollectors.RecordCount
    
        g_rstCollectors.AbsolutePosition = i
    
        StatusBar1.Panels(1).text = "Create sheet " & i
        Set oSheet = oBook.Worksheets(i)
        
        ocmd.Parameters("@_iCollector") = g_rstCollectors.Fields("UserID")
        
        Set orst = ocmd.Execute
        oSheet.Range("A2").CopyFromRecordset orst
        orst.Close
        With oSheet
            .Name = g_rstCollectors.Fields("UserID")
            .Range("A1:N1").value = Array("CustID", "CustName", "Contact", "Phone", "Ext", "90+", "60+", "45+", "Total", "Late", "Over", "Terms", "OnHold", "Status")
            .Range("A1:N1").Font.Bold = True
            .Cells.Font.Name = "arial"
            .Cells.Font.Size = 7
            .Rows.RowHeight = 15
            Set oRange = .Range("D1").EntireColumn
            oRange.TextToColumns
            oRange.NumberFormat = "[<=9999999]###-####;(###) ###-####"
            oRange.HorizontalAlignment = xlLeft
            .Range("F2:K2").EntireColumn.NumberFormat = "#,##0.00"
            .Columns.AutoFit
            
            StatusBar1.Panels(1).text = "Layout sheet " & i & " for printing"
            .PageSetup.Orientation = xlLandscape
            .PageSetup.PrintTitleRows = "$1:$1"
            .PageSetup.LeftFooter = oSheet.Name
            .PageSetup.CenterFooter = Date
            .PageSetup.RightFooter = "Page &P"
            .PageSetup.TopMargin = Application.InchesToPoints(0.5)
            .PageSetup.HeaderMargin = Application.InchesToPoints(0.5)
            .PageSetup.LeftMargin = Application.InchesToPoints(0.25)
            .PageSetup.RightMargin = Application.InchesToPoints(0.25)

            .PageSetup.BottomMargin = Application.InchesToPoints(0.6)
            .PageSetup.FooterMargin = Application.InchesToPoints(0.2)
        End With
    Next i

    g_rstCollectors.Filter = adFilterNone
    
    'Save the Workbook and Quit Excel
    StatusBar1.Panels(1).text = "Saving spreadsheet"
    
    oExcel.ActiveWorkbook.SaveAs g_CollectionsPath & "Collections" & Format$(Date, "mmddyyyy") & ".xls"
 
    GoTo Cleanup

EH:
    MsgBox "cmdCreateSheets_Click " & Err.Number & " " & Err.Description

Cleanup:
    oExcel.Quit
    StatusBar1.Panels(1).text = "Done"
    Screen.MousePointer = vbDefault
End Sub


'open the datasourcetemplate excel file
'load it with the selected query
Private Sub cmdCreate_Click()
    Dim ocmd As ADODB.Command
    Dim orst As ADODB.Recordset
    Dim oExcel As Object
    Dim oBook As Excel.Workbook
    Dim oSheets As Excel.Worksheets
    Dim oSheet As Excel.Worksheet
    Dim sFilePrefix As String

    Screen.MousePointer = vbHourglass

    lblFileName.Caption = vbNullString

    Set ocmd = New ADODB.Command
    ocmd.ActiveConnection = g_DB.Connection
    ocmd.CommandType = adCmdText
    
    StatusBar1.Panels(1).text = "Query database"
    
    If optOver45.value = True Then
        ocmd.CommandText = "spcpcARover45letters"
        sFilePrefix = "Over45("
        m_sLetterPath = g_CollectionsPath & "Over45Letter.doc"
    ElseIf optOver60.value = True Then
        ocmd.CommandText = "spcpcARover60letters"
        sFilePrefix = "Over60("
        m_sLetterPath = g_CollectionsPath & "Over60Letter.doc"
    ElseIf optOver90.value = True Then
        ocmd.CommandText = "spcpcARover90letters"
        sFilePrefix = "Over90("
        m_sLetterPath = g_CollectionsPath & "Over90Letter.doc"
    End If

    Set orst = ocmd.Execute

    If orst.RecordCount = 0 Then
        MsgBox "No records found"
        GoTo Cleanup
    End If
    
    StatusBar1.Panels(1).text = "Create Excel workbook"

    Set oExcel = New Excel.Application
    oExcel.Workbooks.Open g_CollectionsPath & "DataSourceTemplate.xls"

    Set oSheet = oExcel.Workbooks(1).Worksheets(1)
    
    'Get out of here if the datasource template is sullied (7/14/10)
    If oSheet.Range("A2").value <> "" Then
        MsgBox "Your Excel datasource template has data in it, and it shouldn't. Call IT to straighten it out.  Thanks.", vbCritical
        GoTo Cleanup
    End If
    
    StatusBar1.Panels(1).text = "Load worksheet"

    oSheet.Range("A2").CopyFromRecordset orst
    
    m_sFileName = sFilePrefix & Replace(FormatDateTime(Date, vbShortDate), "/", "-") & ").xls"

    lblFileName.Caption = m_sFileName & " [" & orst.RecordCount & " records]"
    
    orst.Close
    
'    oExcel.Workbooks(1).SaveAs m_sFilePath & m_sFileName
    oExcel.ActiveWorkbook.SaveAs g_CollectionsPath & m_sFileName
    
    oExcel.Visible = True

    Set oExcel = Nothing
    
Cleanup:
    StatusBar1.Panels(1).text = "Done"
    Screen.MousePointer = vbDefault
End Sub


Private Sub cmdBrowse_Click()
    With dlgOpen
        .CancelError = False
        .DialogTitle = "Get SpreadSheet"
        .InitDir = g_CollectionsPath
        .DefaultExt = "*.xls"
        .Filter = "Excel Files (*.xls)|*.xls"
        .ShowOpen
        m_sFileName = .FileTitle
        lblFileName.Caption = m_sFileName
        If InStr(1, m_sFileName, "Over45") Then
            m_sLetterPath = g_CollectionsPath & "Over45Letter.doc"
        ElseIf InStr(1, m_sFileName, "Over60") Then
            m_sLetterPath = g_CollectionsPath & "Over60Letter.doc"
        ElseIf InStr(1, m_sFileName, "Over90") Then
            m_sLetterPath = g_CollectionsPath & "Over90Letter.doc"
        Else
            MsgBox "Inappropriate selection"
        End If
    End With
End Sub


Private Sub cmdMailMerge_Click()
    'Dim oWord As Word.Application
    Dim oWord As Object
    
    On Error GoTo EH
    Set oWord = New Word.Application
    oWord.Documents.Open m_sLetterPath
    oWord.Visible = True
    oWord.Documents(1).MailMerge.OpenDataSource g_CollectionsPath & m_sFileName
    Exit Sub
EH:
    MsgBox Err.Number & ":" & Err.Description, , "Error Opening Word"
End Sub


'mark entries in tcpCustHold to indicate if they've received a letter (45 and/or 60)

' m_sFileName is set by both cmdCreate_Click() and cmdBrowse_Click()

Private Sub cmdMark_Click()
    Dim oXLcon As ADODB.Connection
    Dim orst As ADODB.Recordset
    Dim ocmd As ADODB.Command
    Dim sSheetName As String
    Dim lCustKey As Long
    Dim k As Long
    Dim sSQL As String
    
    On Error GoTo EH
    
    If m_sFileName = vbNullString Then
        MsgBox "No Data Source Selected."
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    StatusBar1.Panels(1).text = ""
    
    If optOver45.value = True Then
        sSQL = "UPDATE tcpCustHold set letter45 = -1, letter60 = 0, letter90 = 0 where custkey="
    ElseIf optOver60.value = True Then
        sSQL = "UPDATE tcpCustHold set letter45 = 0, letter60 = -1, letter90 = 0 where custkey="
    ElseIf optOver90.value = True Then
        sSQL = "UPDATE tcpCustHold set letter45 = 0, letter60 = 0, letter90 = -1 where custkey="
    End If

    'open the spreadsheet for input
    Set oXLcon = ConnectExcel(g_CollectionsPath & m_sFileName)
    Set orst = New ADODB.Recordset
    orst.Open "SELECT * FROM [Sheet1$]", oXLcon, adOpenStatic, adLockReadOnly
    
    Set ocmd = New ADODB.Command
    ocmd.ActiveConnection = g_DB.Connection
    ocmd.CommandType = adCmdText
    
    g_DB.Connection.BeginTrans

    Do While Not orst.EOF

        lCustKey = IIf(IsNull(orst.Fields("CustKey").value), 0, orst.Fields("CustKey").value)
        If lCustKey <> 0 Then
            ocmd.CommandText = sSQL & lCustKey
            ocmd.Execute
        End If

        orst.MoveNext
        k = k + 1
        StatusBar1.Panels(1).text = k
        DoEvents
    Loop

    g_DB.Connection.CommitTrans
     
    If orst.RecordCount = 0 Then
        MsgBox "No records are available from spreadsheet"
        Exit Sub
    Else
        ProgressBar1.Min = 0
        ProgressBar1.Max = orst.RecordCount
    End If
    
    If vbYes = MsgBox("Ready to update " & orst.RecordCount & " customer records", vbYesNo) Then
        InjectRemarks orst
    End If

    Set orst = Nothing
    
    GoTo Continue
    
EH:
    g_DB.Connection.RollbackTrans
    MsgBox "Error: [" & Err.Number & "] " & Err.Description
    
Continue:
    oXLcon.Close
    Set oXLcon = Nothing
    StatusBar1.Panels(1).text = "Done"
    Screen.MousePointer = vbDefault
End Sub


Private Sub InjectRemarks(ByRef orst As ADODB.Recordset)
    Dim oRemarkContext As MemoMeister.RemarkContext
    Dim i As Long
    Dim sMessage As String

    On Error GoTo EH

    If optOver45.value = True Then
        sMessage = "Sent 45-day letter."
    ElseIf optOver60.value = True Then
        sMessage = "Sent 60-day letter."
    End If
    
    g_DB.Connection.BeginTrans

    orst.MoveFirst

    While Not orst.EOF
        If Not IsNull(orst.Fields("CustID").value) Then
            Set oRemarkContext = New RemarkContext
            oRemarkContext.Load "ARCustLoad", Trim(orst.Fields("CustID").value)
            'Add the remittance date for 90 day accounts. (The day to send them to collections)
            If optOver90.value = True Then
                sMessage = "Sent 90-day letter. Send to Collections on " & Trim(orst.Fields("RemittanceDate").value)
            End If
            
            oRemarkContext.AddRemark "Cust.AR.Coll", sMessage
            
            oRemarkContext.Save True
            Set oRemarkContext = Nothing

            i = i + 1
            ProgressBar1.value = i
            DoEvents
        End If

        orst.MoveNext
    Wend

    g_DB.Connection.CommitTrans
    
    MsgBox "Finished updating customer records"
    ProgressBar1 = 0
    Exit Sub
    
EH:
    g_DB.Connection.RollbackTrans
    MsgBox "Error while updating customer remarks" & vbCrLf & _
        "Pass " & i & " CustID " & orst.Fields("CustID").value & vbCrLf & _
        " " & Err.Number & " " & Err.Description
End Sub



Private Function ConnectExcel(filepath As String) As ADODB.Connection
    Dim ocon As ADODB.Connection
    
    On Error GoTo EH
    
    Set ocon = New ADODB.Connection
    ocon.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
               "Data Source=" & filepath & ";" & _
               "Extended Properties=""Excel 8.0;HDR=yes;"""
    Set ConnectExcel = ocon
    Exit Function
EH:
    MsgBox "ADO Error: " & Err.Number & " " & Err.Description
    Stop
End Function


Private Function ConnectSQL(Database As String) As ADODB.Connection
    Dim ocon As ADODB.Connection

    On Error GoTo EH
    
    Set ocon = New ADODB.Connection
    With ocon
        .ConnectionString = "Provider=SQLOLEDB.1;Data Source=CPSQLPro;Initial Catalog=" & Database & ";User ID=sa"
        .CursorLocation = adUseClient
        .Open
    End With
    Set ConnectSQL = ocon

    Exit Function
EH:
    MsgBox Err.Number & " - " & Err.Description, vbExclamation + vbOKOnly, "Error connecting to database"
End Function

