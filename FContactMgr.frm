VERSION 5.00
Object = "{CAF0FDE4-8332-11CF-BC13-0020AFD6738C}#1.0#0"; "newsota.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form FContactMgr 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6885
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   6885
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraEdit 
      BorderStyle     =   0  'None
      Height          =   5775
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   6975
      Begin VB.ComboBox cboEMailFormat 
         Height          =   315
         ItemData        =   "FContactMgr.frx":0000
         Left            =   5400
         List            =   "FContactMgr.frx":0002
         TabIndex        =   6
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Frame Frame1 
         Caption         =   "Roles"
         Height          =   1815
         Left            =   3480
         TabIndex        =   43
         Top             =   3360
         Width           =   3255
         Begin VB.CheckBox chkRole 
            Caption         =   "Management"
            Enabled         =   0   'False
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   49
            Top             =   1440
            Width           =   3015
         End
         Begin VB.CheckBox chkRole 
            Caption         =   "Reconciles Statements"
            Enabled         =   0   'False
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   48
            Top             =   1200
            Width           =   3015
         End
         Begin VB.CheckBox chkRole 
            Caption         =   "Approves Payments"
            Enabled         =   0   'False
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   47
            Top             =   960
            Width           =   3015
         End
         Begin VB.CheckBox chkRole 
            Caption         =   "Chooses Vendors"
            Enabled         =   0   'False
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   46
            Top             =   720
            Width           =   3015
         End
         Begin VB.CheckBox chkRole 
            Caption         =   "Issues POs"
            Enabled         =   0   'False
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   45
            Top             =   480
            Width           =   3015
         End
         Begin VB.CheckBox chkRole 
            Caption         =   "Places Parts Orders"
            Enabled         =   0   'False
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   44
            Top             =   240
            Width           =   3015
         End
      End
      Begin VB.TextBox txtFirstName 
         Height          =   375
         Left            =   1440
         MaxLength       =   19
         TabIndex        =   3
         Top             =   180
         Width           =   2655
      End
      Begin VB.TextBox txtLastName 
         Height          =   375
         Left            =   1440
         MaxLength       =   20
         TabIndex        =   4
         Top             =   600
         Width           =   2655
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   375
         Left            =   5760
         TabIndex        =   20
         Top             =   5280
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtFaxExt 
         Height          =   375
         Left            =   3120
         MaxLength       =   4
         TabIndex        =   12
         Top             =   2820
         Width           =   495
      End
      Begin VB.TextBox txtPhoneExt 
         Height          =   375
         Left            =   3120
         MaxLength       =   4
         TabIndex        =   9
         Top             =   1980
         Width           =   495
      End
      Begin VB.TextBox txtEmail 
         Height          =   375
         Left            =   1440
         MaxLength       =   40
         TabIndex        =   5
         Top             =   1020
         Width           =   2655
      End
      Begin VB.Frame fraNotify 
         Caption         =   "Notifications"
         Height          =   1815
         Left            =   120
         TabIndex        =   22
         Top             =   3360
         Width           =   3135
         Begin VB.CheckBox chkNotify 
            Caption         =   "Operational"
            Enabled         =   0   'False
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   42
            Top             =   1440
            Width           =   2295
         End
         Begin VB.CheckBox chkNotify 
            Caption         =   "Promotional"
            Enabled         =   0   'False
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   26
            Top             =   1200
            Width           =   2295
         End
         Begin VB.CheckBox chkNotify 
            Caption         =   "Statements"
            Enabled         =   0   'False
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   25
            Top             =   960
            Width           =   1215
         End
         Begin VB.CheckBox chkNotify 
            Caption         =   "Invoices"
            Enabled         =   0   'False
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   24
            Top             =   720
            Width           =   1215
         End
         Begin VB.CheckBox chkNotify 
            Caption         =   "Shipment Notification"
            Enabled         =   0   'False
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   2055
         End
         Begin VB.CheckBox chkNotify 
            Caption         =   "Order Status"
            Enabled         =   0   'False
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   23
            Top             =   480
            Width           =   2175
         End
      End
      Begin VB.CheckBox chkDeclinedEAddr 
         Caption         =   "Declined to provide email address"
         Height          =   255
         Left            =   1440
         TabIndex        =   7
         Top             =   1560
         Width           =   2895
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "&New"
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   5280
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdAddUpdate 
         Caption         =   "&Add"
         Height          =   375
         Left            =   1200
         TabIndex        =   15
         Top             =   5280
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   2280
         TabIndex        =   16
         Top             =   5280
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdUndo 
         Caption         =   "&Undo"
         Height          =   375
         Left            =   3480
         TabIndex        =   17
         Top             =   5280
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   375
         Left            =   4680
         TabIndex        =   18
         Top             =   5280
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   5760
         TabIndex        =   19
         Top             =   5280
         Visible         =   0   'False
         Width           =   975
      End
      Begin NEWSOTALib.SOTAMaskedEdit meFax 
         Height          =   375
         Left            =   1440
         TabIndex        =   11
         Top             =   2820
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   93
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin NEWSOTALib.SOTAMaskedEdit mePhone 
         Height          =   375
         Left            =   1440
         TabIndex        =   8
         Top             =   1980
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   93
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin NEWSOTALib.SOTAMaskedEdit meCellPhone 
         Height          =   375
         Left            =   1440
         TabIndex        =   10
         Top             =   2400
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   93
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   "EMail Format"
         Height          =   255
         Left            =   4320
         TabIndex        =   50
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label lblUpdateUserID 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   4320
         TabIndex        =   41
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label lblCreateUserID 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   4320
         TabIndex        =   40
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "Updated By:"
         Height          =   255
         Left            =   4320
         TabIndex        =   39
         Top             =   2640
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "Created By:"
         Height          =   255
         Left            =   4320
         TabIndex        =   38
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "First"
         Height          =   255
         Left            =   600
         TabIndex        =   37
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Last"
         Height          =   255
         Left            =   600
         TabIndex        =   36
         Top             =   600
         Width           =   615
      End
      Begin VB.Label lblCreateDate 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   5640
         TabIndex        =   34
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Cell Phone"
         Height          =   255
         Left            =   0
         TabIndex        =   32
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Fax"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   2820
         Width           =   1095
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Business Phone"
         Height          =   255
         Left            =   0
         TabIndex        =   30
         Top             =   1980
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "EMail"
         Height          =   255
         Left            =   720
         TabIndex        =   29
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   2940
         TabIndex        =   28
         Top             =   2040
         Width           =   135
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   2940
         TabIndex        =   27
         Top             =   2880
         Width           =   135
      End
      Begin VB.Label lblUpdateDate 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   5640
         TabIndex        =   35
         Top             =   2880
         Width           =   1095
      End
   End
   Begin VB.Frame fraGrid 
      BorderStyle     =   0  'None
      Height          =   3855
      Left            =   0
      TabIndex        =   33
      Top             =   0
      Visible         =   0   'False
      Width           =   6975
      Begin GridEX20.GridEX gdxOldContacts 
         Height          =   2955
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   5212
         Version         =   "2.0"
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         UseEvenOddColor =   -1  'True
         MethodHoldFields=   -1  'True
         Options         =   8
         RecordsetType   =   1
         AllowEdit       =   0   'False
         GroupByBoxVisible=   0   'False
         DataMode        =   1
         ColumnHeaderHeight=   285
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   3
         Column(1)       =   "FContactMgr.frx":0004
         Column(2)       =   "FContactMgr.frx":00F0
         Column(3)       =   "FContactMgr.frx":01DC
         FormatStylesCount=   6
         FormatStyle(1)  =   "FContactMgr.frx":02C0
         FormatStyle(2)  =   "FContactMgr.frx":03A0
         FormatStyle(3)  =   "FContactMgr.frx":04D8
         FormatStyle(4)  =   "FContactMgr.frx":0588
         FormatStyle(5)  =   "FContactMgr.frx":063C
         FormatStyle(6)  =   "FContactMgr.frx":0714
         ImageCount      =   0
         PrinterProperties=   "FContactMgr.frx":07CC
      End
      Begin VB.CommandButton cmdLookUp 
         Caption         =   "Show Old Phone Numbers"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   3360
         Width           =   2415
      End
      Begin GridEX20.GridEX gdxContacts 
         Height          =   2955
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   5212
         Version         =   "2.0"
         AllowRowSizing  =   -1  'True
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         HideSelection   =   2
         UseEvenOddColor =   -1  'True
         MethodHoldFields=   -1  'True
         AllowEdit       =   0   'False
         GroupByBoxVisible=   0   'False
         ItemCount       =   0
         DataMode        =   99
         ColumnHeaderHeight=   285
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   6
         Column(1)       =   "FContactMgr.frx":09A4
         Column(2)       =   "FContactMgr.frx":0AC8
         Column(3)       =   "FContactMgr.frx":0BB4
         Column(4)       =   "FContactMgr.frx":0CCC
         Column(5)       =   "FContactMgr.frx":0DEC
         Column(6)       =   "FContactMgr.frx":0EF8
         FormatStylesCount=   6
         FormatStyle(1)  =   "FContactMgr.frx":1018
         FormatStyle(2)  =   "FContactMgr.frx":10F8
         FormatStyle(3)  =   "FContactMgr.frx":1230
         FormatStyle(4)  =   "FContactMgr.frx":12E0
         FormatStyle(5)  =   "FContactMgr.frx":1394
         FormatStyle(6)  =   "FContactMgr.frx":146C
         ImageCount      =   0
         PrinterProperties=   "FContactMgr.frx":1524
      End
   End
End
Attribute VB_Name = "FContactMgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const k_lClassAlways = 1

Private Enum FormMode
    opEditor = 1
    opManager = 2
End Enum

Private m_Mode As FormMode
Private WithEvents m_oContacts As Contacts
Attribute m_oContacts.VB_VarHelpID = -1
'Our alias for m_oContacts.selContact
Private WithEvents m_oContact As Contact
Attribute m_oContact.VB_VarHelpID = -1
Private WithEvents m_oBrokenRules As BrokenRules
Attribute m_oBrokenRules.VB_VarHelpID = -1
Private m_bLoading As Boolean
'Communicate back to FOrder
Private m_bCancel As Boolean
Private m_oPrevCntrl As Control
Private m_sUserID As String

'*** Public properties

Public Property Get Cancel() As Boolean
    Cancel = m_bCancel
End Property


'*****************************************************************************
' Public Methods
'*****************************************************************************

'Entry point for the Contacts Manager

'Public Function Init(oCust As Customer, bSelect As Boolean) As Integer
'***DH 11/27/07
Public Sub Init(oContacts As Contacts, i_sUserID As String)
'***DH 4/1/08
    m_sUserID = i_sUserID

    m_Mode = opManager
    cmdNew.Visible = True
    cmdAddUpdate.Visible = True
    cmdDelete.Visible = True
    cmdUndo.Visible = True
    cmdClose.Visible = True
    'Need to do this on the fly because only one button can have this property set at a time.
    cmdClose.Cancel = True
    
    ShowGrids

    Me.Caption = "Contact Manager (" + oContacts.OwnerID + " " + oContacts.OwnerName + ")"

    Set m_oContacts = oContacts
    
    gdxContacts.Visible = True
    gdxOldContacts.Visible = False
    
    GetOldPhoneNbrs
    LoadValidationRules

    m_bLoading = True
    
    'Unbound grid loaded & columns autosized here...
    'ItemCount assignment fires off gdxContacts_UnboundReadData
    gdxContacts.ItemCount = m_oContacts.Count
    GridColAutoSize gdxContacts

    m_bLoading = False
    
    If m_oContacts.Count > 0 Then
        'setup our alias
        Set m_oContact = m_oContacts(1)
        Call DisplayContact
    Else
        Call NewContact
    End If
    gdxContacts.Row = 1
    
    Me.Show vbModal
    Unload Me
End Sub


'Entry Point for the Contact object Editor

Public Sub Edit(ByRef oContact As Contact, sName As String, ByVal oType As OwnerType, lOwnerKey As Long)
    Dim i As Integer
    
    m_Mode = opEditor
    
    'Local alias
    Set m_oContact = oContact
    
    LoadValidationRules
        
    If m_oContact.Key > 0 Then  'Existing Contact
        'this fires off the Form_Load event
        If oType = opOrder Then
            Me.Caption = "Edit Contact (This Order Only)"
        Else
            Me.Caption = "Edit Contact (Customer)"
            
        End If

        DisplayContact
    Else    'New Contact
        'this fires off the Form_Load event
        If oType = opOrder Then
            Me.Caption = "Create Contact (This Order Only)"
        Else
            Me.Caption = "Create Contact (Customer)"
        End If
        
        m_oContact.State.SetBits eMask.IsDirty
        m_oContact.OwnerType = oType
        m_oContact.OwnerKey = lOwnerKey

        'Parse the name into FirstName & LastName
        Dim s As String
        Dim fName As String
        Dim lName As String
        
        s = Trim$(sName)

        If InStr(s, " ") = 0 Then
            fName = s
        Else
            fName = Mid$(s, 1, InStr(s, " ") - 1)
            lName = LTrim$(Mid(s, Len(fName) + 1))
            m_oContact.LastName = lName
        End If
        
        m_oContact.FirstName = fName
        
        m_bLoading = True
        txtFirstName.text = m_oContact.FirstName
        txtLastName.text = m_oContact.LastName
        cboEMailFormat.ListIndex = m_oContact.EMailFormat - 1
        m_bLoading = False
        
        SetButtons m_oContact.State
        m_oBrokenRules.Validate
        SetIsValid
    End If

    cmdSave.Visible = True
    cmdCancel.Visible = True
    
    'Need to do this on the fly because only one button can have this property set at a time.
    cmdCancel.Cancel = True
End Sub


'*****************************************************************************
' Event handlers
'*****************************************************************************

'*** Form Events

Private Sub Form_Load()
    m_bLoading = True
    mePhone.mask = "(###) ###-####"
    meCellPhone.mask = "(###) ###-####"
    meFax.mask = "(###) ###-####"
    gdxContacts.TabKeyBehavior = jgexControlNavigation
    m_bLoading = False

    LoadEmailFormatCombo
End Sub


Private Sub Form_Activate()
    If m_Mode = opManager Then
        If m_oContacts.Count = 0 Then
            txtFirstName.SetFocus
        Else
            If txtFirstName.text = "" Then
                txtFirstName.SetFocus
            ElseIf txtLastName.text = "" Then
                txtLastName.SetFocus
            ElseIf mePhone.text = "" Then
                mePhone.SetFocus
            Else
                cmdClose.SetFocus
            End If
        End If
    Else 'Editor mode
        If txtFirstName.text = "" Then
            txtFirstName.SetFocus
        ElseIf txtLastName.text = "" Then
            txtLastName.SetFocus
        Else
            mePhone.SetFocus
        End If
    End If
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set m_oContact = Nothing
    Set m_oContacts = Nothing
    m_oBrokenRules.Destroy
    Set m_oBrokenRules = Nothing
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
        KeyAscii = 0
    End If
End Sub

'*** Object Events

Private Sub m_oContacts_StateChange(NewState As BitMap)
    SetButtons NewState
End Sub

Private Sub m_oContact_StateChange(NewState As BitMap)
    SetButtons NewState
End Sub

'*** Button Events

Private Sub cmdNew_Click()
    Call NewContact
    
    txtFirstName.SetFocus
End Sub


Private Sub cmdAddUpdate_Click()
    On Error GoTo EH
'***DH 4/1/08
    'add the new contact to the collection
    'or update the selected one
    'm_oContacts.AddUpdate
    'reassign our alias
    'Set m_oContact = m_oContacts.selContact

    m_oContact.UserID = m_sUserID
    
    If m_oContact.State(eMask.IsNew) Then
        'Insert into the database
        m_oContact.Insert
        'Add to the colection in Alphabetical order
        m_oContacts.InsertNewContact m_oContact
    Else
        'Update the database
        m_oContact.Update
    End If
'***
    'Update the details
    DisplayContact
    
    'Update grid
    gdxContacts.SetFocus
    
    gdxContacts.ItemCount = m_oContacts.Count
    GridColAutoSize gdxContacts
    
    gdxContacts.Refetch False
'***DH 3/19/08
'Note: When updating a record's first name that causes it to be inserted into the FIRST row,
'the row gets selected, but the the highlighted row does not change.
'This kludge is to work around the UI behavior.
    If m_oContacts.SelIndex = 1 Then
        gdxContacts.Row = 0
    Else
        gdxContacts.Row = m_oContacts.SelIndex
    End If
    
    Exit Sub
EH:
    If Err.Number = 1 Then
        MsgBox Err.Description, vbExclamation, Err.Source & " Contact"
        gdxContacts.SetFocus
        SetButtons m_oContact.State
        Exit Sub
    Else
        MsgBox Err.Number & " " & Err.Description, vbExclamation
    End If
End Sub


Private Sub cmdDelete_Click()
    If vbYes = MsgBox("Are you sure you want to permanently delete this contact?", vbYesNo, "Delete Contact") Then

'***DH 4/1/08
        m_oContact.UserID = m_sUserID
        'm_oContacts.Delete
        m_oContact.Delete
        m_oContacts.Remove
        
        gdxContacts.ItemCount = m_oContacts.Count
        gdxContacts.Row = 1
        gdxContacts.Refetch False
        
        If m_oContacts.Count > 0 Then
            ChangeSelectedContact gdxContacts.value(1) 'CntctKey
            DisplayContact
        Else
            '***466 if no contacts, manager will be in 'New' contact state.
            NewContact
            MsgBox "Please enter a new contact or click 'Close' to exit.", vbInformation, "Contact Manager"
            txtFirstName.SetFocus
        End If
        
    End If
End Sub


Private Sub cmdUndo_Click()
    m_oContact.Restore
    DisplayContact
    txtFirstName.SetFocus
End Sub


'Manager
Private Sub cmdClose_Click()
    'This assumes we are getting a lost focus event on the
    'control before the this buttom was clicked.
    If Not Me.ActiveControl.Name = "cmdClose" Then
        RunLostFocusEvent Me.ActiveControl
    End If
    
    'Guard against deleting the last contact.
    If Not m_oContact Is Nothing Then
        If m_oContact.IsDirty Then
            Select Case MsgBox("Do you want to save changes?", vbYesNoCancel, "Save Changes")
                Case vbYes
                    If m_oContact.IsValid Then
'***DH 4/1/08
                        m_oContact.UserID = m_sUserID
                        'm_oContacts.AddUpdate
                        If m_oContact.State(eMask.IsNew) Then
                            'Insert into the database
                            m_oContact.Insert
                            'Add to the colection in Alphabetical order
                            m_oContacts.InsertNewContact m_oContact
                        Else
                            'Update the database
                            m_oContact.Update
                        End If
                    Else
                        Exit Sub
                    End If
                Case vbNo
                    m_oContact.Restore
                Case vbCancel
                    Exit Sub
            End Select
        End If
    End If
    
    Unload Me
End Sub

'Editor mode buttons

Private Sub cmdCancel_Click()
    'This assumes we are getting a lost focus event on the
    'control before the this buttom was clicked.
    If Not Me.ActiveControl.Name = "cmdCancel" Then
        RunLostFocusEvent Me.ActiveControl
    End If

    If Not m_oContact.State(eMask.IsNew) Then
        If m_oContact.State(eMask.IsDirty) Then
            Select Case MsgBox("Do you want to save changes?", vbYesNoCancel, "Save Changes")
                Case vbYes
                    If m_oContact.IsValid Then
'***DH 4/1/08
                        m_oContact.UserID = m_sUserID
                        m_oContact.Update
                        m_bCancel = False
                        Me.Hide
                        Exit Sub
                    Else
                        Exit Sub
                    End If
                Case vbNo
                       m_oContact.Restore
                        m_bCancel = True
                        'This form gets unloaded in the calling method. (Contact.Edit)
                        Me.Hide
                Case vbCancel
                    Exit Sub
            End Select
        Else
            m_bCancel = True
            'This form gets unloaded in the calling method. (Contact.Edit)
            Me.Hide
        End If
    Else
        m_bCancel = True
        'This form gets unloaded in the calling method. (Contact.Edit)
        Me.Hide
    End If
    
End Sub


Private Sub cmdSave_Click()
    On Error GoTo EH

    If m_oContact.State(eMask.IsNew) Then
        m_oContact.Insert
    Else
        m_oContact.Update
    End If
    
    m_bCancel = False
    Me.Hide
    Exit Sub

EH:
    If Err.Number = 1 Then
        MsgBox Err.Description, vbExclamation, Err.Source & " Contact"
        txtFirstName.SetFocus
        SetButtons m_oContact.State
        Exit Sub
    Else
        MsgBox Err.Number & " " & Err.Description, vbExclamation
    End If
End Sub


Private Sub cmdLookUp_Click()
    With cmdLookUp
        If .Caption = "Show Old Phone Numbers" Then
            .Caption = "Show New Contacts"
            gdxContacts.Visible = False
            gdxOldContacts.Visible = True
            ClearDisplay
            EnableControls False
        Else
            .Caption = "Show Old Phone Numbers"
            gdxOldContacts.Visible = False
            gdxContacts.Visible = True
            EnableControls True
            DisplayContact
        End If
    End With
End Sub

'*** Control Events

Private Sub txtFirstName_Change()
    If m_bLoading Then Exit Sub
    
    m_oContact.State.SetBits eMask.IsDirty
End Sub

Private Sub txtFirstName_GotFocus()
    Set m_oPrevCntrl = txtFirstName
    
    txtFirstName.SelStart = 0
    txtFirstName.SelLength = Len(txtFirstName.text)
End Sub

Private Sub txtFirstName_LostFocus()
    m_oBrokenRules.Validate txtFirstName
    SetIsValid
    
    If LCase(m_oContact.FirstName) <> LCase(Trim$(txtFirstName.text)) Then
        m_oContact.FirstName = txtFirstName.text
    End If
End Sub

Private Sub txtLastName_Change()
    If m_bLoading Then Exit Sub
    
    m_oContact.State.SetBits eMask.IsDirty
End Sub

Private Sub txtLastName_GotFocus()
    Set m_oPrevCntrl = txtLastName
    
    txtLastName.SelStart = 0
    txtLastName.SelLength = Len(txtLastName.text)
End Sub

Private Sub txtLastName_LostFocus()
    m_oBrokenRules.Validate txtLastName
    SetIsValid
    m_oContact.LastName = txtLastName.text
End Sub

Private Sub txtEmail_Change()
    If m_bLoading Then Exit Sub
    
    m_oContact.State.SetBits eMask.IsDirty
End Sub

Private Sub txtEmail_GotFocus()
    Set m_oPrevCntrl = txtEmail
    
    txtEmail.SelStart = 0
    txtEmail.SelLength = Len(txtEmail.text)
End Sub

Private Sub txtEmail_LostFocus()
    m_oBrokenRules.Validate txtEmail
    
    If Len(txtEmail.text) = 0 Or Not m_oBrokenRules.IsValid(txtEmail) Then
        chkNotify(1).value = vbUnchecked
        chkNotify(1).Enabled = False
    Else 'Valid Email address
        chkNotify(1).Enabled = True
    End If
    
    SetIsValid
    m_oContact.emailaddr = txtEmail.text
End Sub

Private Sub cboEMailFormat_Click()
    If m_bLoading Then Exit Sub
    
    m_oContact.EMailFormat = cboEMailFormat.ItemData(cboEMailFormat.ListIndex)
    m_oContact.State.SetBits eMask.IsDirty
End Sub

Private Sub chkDeclinedEAddr_Click()
    If m_bLoading Then Exit Sub
    
    m_oContact.DeclinedEmailAddr = IIf(chkDeclinedEAddr.value = vbChecked, True, False)
    m_oContact.State.SetBits eMask.IsDirty
End Sub


Private Sub mePhone_KeyPress(KeyAscii As Integer)
    'Not Numbers 0-9
    If KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
End Sub

Private Sub mePhone_GotFocus()
    Set m_oPrevCntrl = mePhone
End Sub

Private Sub mePhone_Change()
    If m_bLoading Then Exit Sub
    m_oContact.State.SetBits eMask.IsDirty
End Sub

Private Sub mePhone_LostFocus()
    m_oBrokenRules.Validate mePhone
    SetIsValid
    m_oContact.Phone = mePhone.text
End Sub


Private Sub txtPhoneExt_GotFocus()
    Set m_oPrevCntrl = txtPhoneExt
End Sub

Private Sub txtPhoneExt_Change()
    If m_bLoading Then Exit Sub
    
    m_oContact.State.SetBits eMask.IsDirty
End Sub

Private Sub txtPhoneExt_LostFocus()
    m_oContact.PhoneExt = txtPhoneExt.text
End Sub


Private Sub meCellPhone_GotFocus()
    Set m_oPrevCntrl = meCellPhone
End Sub

Private Sub meCellPhone_KeyPress(KeyAscii As Integer)
    'Not Numbers 0-9
    If KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
End Sub

Private Sub meCellPhone_Change()
    If m_bLoading Then Exit Sub
    
    m_oContact.State.SetBits eMask.IsDirty
End Sub

Private Sub meCellPhone_LostFocus()
    m_oBrokenRules.Validate meCellPhone
    SetIsValid
    m_oContact.CellPhone = meCellPhone.text
End Sub


Private Sub meFax_GotFocus()
    Set m_oPrevCntrl = meFax
End Sub

Private Sub meFax_KeyPress(KeyAscii As Integer)
    'Not Numbers 0-9
    If KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
End Sub

Private Sub meFax_Change()
    If m_bLoading Then Exit Sub
    
    m_oContact.State.SetBits eMask.IsDirty
End Sub
    
Private Sub meFax_LostFocus()
    m_oBrokenRules.Validate meFax
    SetIsValid
    m_oContact.Fax = meFax.text
End Sub


Private Sub txtFaxExt_GotFocus()
    Set m_oPrevCntrl = txtFaxExt
End Sub

Private Sub txtFaxExt_Change()
    If m_bLoading Then Exit Sub
    
    m_oContact.State.SetBits eMask.IsDirty
End Sub

Private Sub txtFaxExt_LostFocus()
    m_oContact.FaxExt = txtFaxExt.text
End Sub


Private Sub chkNotify_Click(Index As Integer)
    If m_bLoading Then Exit Sub
    
    If chkNotify(Index).value = vbChecked Then
        m_oContact.Notifications.Add CLng(Index)
    Else
        m_oContact.Notifications.Remove Index
    End If
    
    m_oContact.State.SetBits eMask.IsDirty
End Sub


'*** Grids

Private Sub gdxContacts_Click()
    If m_bLoading Then Exit Sub

    ChangeSelectedContact gdxContacts.value(1) 'CntctKey
    DisplayContact
End Sub



Private Sub gdxContacts_SelectionChange()
'***DH 3/19/08
    If m_bLoading Then Exit Sub
    'Note:
    'm_oContact.Key = Key from record before row change.
    'gdxContacts.value(1) = Key from newly selected row.
    'If ((gdxContacts.RowCount > 1) & (m_oContact.Key > 0) & (m_oContact.IsDirty = True)) = True Then
    
    If (m_oContact.IsDirty = True) = True Then
        
        'Fire off LostFocus event of the last control if necessary
        If Not m_oPrevCntrl Is Nothing Then
            RunLostFocusEvent m_oPrevCntrl
        End If
        
        Select Case MsgBox("Do you want to save changes?", vbYesNoCancel, "Save Changes")
            Case vbYes
                If m_oContact.IsValid Then
'***DH 4/1/08
                    m_oContact.UserID = m_sUserID
                    'm_oContacts.AddUpdate
                    If m_oContact.State(eMask.IsNew) Then
                        'Insert into the database
                        m_oContact.Insert
                        'Add to the colection in Alphabetical order
                        m_oContacts.InsertNewContact m_oContact
                    Else
                        'Update the database
                        m_oContact.Update
                    End If
                    
                    ChangeSelectedContact gdxContacts.value(1) 'CntctKey
                    
                    'Update the details
                    DisplayContact
'***DH 4/3/08
                    'Update the grid.
                    gdxContacts.ItemCount = m_oContacts.Count
                    GridColAutoSize gdxContacts
                    gdxContacts.Refetch False
                    'This is to supress the SelectionChange event that will fire off by re-setting the selected row.
                    m_bLoading = True
                    gdxContacts.Row = m_oContacts.SelIndex
                    m_bLoading = False
                Else
                    'This is to supress the SelectionChange event that will fire off by re-setting the selected row.
                    m_bLoading = True
                    gdxContacts.Row = m_oContacts.SelIndex
                    m_bLoading = False
                    Exit Sub
                End If
            Case vbNo
                m_oContact.Restore
                ChangeSelectedContact gdxContacts.value(1) 'CntctKey
                DisplayContact
            Case vbCancel
                'This is to supress the SelectionChange event that will fire off by re-setting the selected row.
                m_bLoading = True
                gdxContacts.Row = m_oContacts.SelIndex
                m_bLoading = False
                Exit Sub
        End Select
    Else
        ChangeSelectedContact gdxContacts.value(1) 'CntctKey
                    
        'Update the details
        DisplayContact
    End If
    
End Sub

Private Sub gdxContacts_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    Dim PhoneNbr As String

    Values(1) = m_oContacts(RowIndex).Key
    Values(2) = m_oContacts(RowIndex).Name

    PhoneNbr = FormatPhoneNumber(m_oContacts(RowIndex).Phone)
    If Len(m_oContacts(RowIndex).PhoneExt) > 0 Then
        PhoneNbr = PhoneNbr & " x" & m_oContacts(RowIndex).PhoneExt
    End If
    
    Values(3) = PhoneNbr
    Values(4) = FormatPhoneNumber(m_oContacts(RowIndex).CellPhone)

    PhoneNbr = FormatPhoneNumber(m_oContacts(RowIndex).Fax)
    If Len(m_oContacts(RowIndex).FaxExt) > 0 Then
        PhoneNbr = PhoneNbr & " x" & m_oContacts(RowIndex).FaxExt
    End If
    
    Values(5) = PhoneNbr
    Values(6) = m_oContacts(RowIndex).emailaddr

End Sub



'*****************************************************************************
' Private subroutines
'*****************************************************************************

Private Sub NewContact()
'***DH 3/31/08
'    m_oContacts.NewContact
'    Set m_oContact = m_oContacts.selContact
    Set m_oContact = New Contact
    'm_oContact.Connection = g_DB.Connection
    m_oContact.OwnerType = opCustomer
    m_oContact.OwnerKey = m_oContacts.OwnerKey
    ClearDisplay
    m_oBrokenRules.Validate
    SetIsValid
    m_oContact.Backup
    SetButtons m_oContact.State
    
End Sub


Private Sub GridColAutoSize(gdx As GridEX)
    gdx.HoldFields
    gdx.HoldSortSettings = True

    Dim i As Integer
    For i = 1 To gdx.Columns.Count
        gdx.Columns(i).AutoSize
    Next
End Sub


Private Sub SetButtons(oState As BitMap)
    If m_bLoading Then Exit Sub

    If m_Mode = opManager Then
        cmdNew.Enabled = Not (oState.TestBits(eMask.IsNew)) And _
            Not (oState.TestBits(eMask.IsDirty))
        cmdAddUpdate.Enabled = oState.TestBits(eMask.IsValid) And oState.TestBits(eMask.IsDirty)
        
        If cmdAddUpdate.Enabled Then
            If oState.TestBits(eMask.IsNew) Then
                cmdAddUpdate.Caption = "&Add"
            Else
                cmdAddUpdate.Caption = "U&pdate"
            End If
        End If
        
        cmdDelete.Enabled = Not (oState.TestBits(eMask.IsNew)) And _
            Not (oState.TestBits(eMask.IsDirty))
        cmdUndo.Enabled = oState.TestBits(eMask.IsDirty)
    Else 'Editor mode
        cmdSave.Enabled = oState.TestBits(eMask.IsValid) And oState.TestBits(eMask.IsDirty)
    End If
End Sub

Private Sub SetIsValid()
    If m_oBrokenRules.Count > 0 Then
        m_oContact.State.ClearBits eMask.IsValid
    Else
        m_oContact.State.SetBits eMask.IsValid
    End If
End Sub


Private Sub DisplayContact()
    Dim j As Integer
            
    '***466 SMR 05/11/2006
    'Example: This is called by cmdundo_click.  When the count is zero,
    '(we are adding the first contact) and undo is clicked, clear display
    'and set buttons don't get called.
    'This procedure also called by Init, Edit, cmdLookUp_Click
    'and gdxContacts_SelectionChange

    m_bLoading = True
    ClearDisplay
            
    With m_oContact
        txtFirstName = .FirstName
        txtLastName = .LastName
        txtEmail.text = .emailaddr
        cboEMailFormat.ListIndex = .EMailFormat - 1
        chkDeclinedEAddr.value = IIf(.DeclinedEmailAddr, vbChecked, vbUnchecked)
        mePhone.text = .Phone
        txtPhoneExt.text = .PhoneExt
        meCellPhone.text = .CellPhone
        meFax.text = .Fax
        txtFaxExt.text = .FaxExt
        lblCreateUserID.Caption = .CreateUserID
        
        If .CreateDate = Empty Then
            lblCreateDate.Caption = ""
        Else
            lblCreateDate.Caption = FormatDateTime(.CreateDate, vbShortDate)
        End If
        
        lblUpdateUserID.Caption = .UpdateUserID
        
        If .UpdateDate = Empty Then
            lblUpdateDate.Caption = ""
        Else
            lblUpdateDate.Caption = FormatDateTime(.UpdateDate, vbShortDate)
        End If

        For j = 1 To .Notifications.Count
            chkNotify(.Notifications.Item(j)) = vbChecked
        Next j
        
        chkNotify(1).Enabled = (Len(txtEmail.text) > 0)
    End With

    m_bLoading = False
    SetButtons m_oContact.State
    m_oBrokenRules.Validate
    SetIsValid
End Sub


Private Sub EnableControls(State As Boolean)
    txtFirstName.Enabled = State
    txtLastName.Enabled = State
    txtEmail.Enabled = State
    mePhone.Enabled = State
    txtPhoneExt.Enabled = State
    meCellPhone.Enabled = State
    meFax.Enabled = State
    txtFaxExt.Enabled = State
    cboEMailFormat.Enabled = State
    chkDeclinedEAddr.Enabled = State
    chkNotify(1).Enabled = State
    
    If State = True Then
        If Not m_oContact Is Nothing Then
            SetButtons m_oContact.State
        End If
        txtFirstName.BackColor = &H80000005
        txtLastName.BackColor = &H80000005
        txtEmail.BackColor = &H80000005
        txtPhoneExt.BackColor = &H80000005
        txtFaxExt.BackColor = &H80000005
    Else
        cmdNew.Enabled = False
        cmdAddUpdate.Enabled = False
        cmdDelete.Enabled = False
        cmdUndo.Enabled = False
        cmdClose.Enabled = True
        txtFirstName.BackColor = &H8000000F
        txtLastName.BackColor = &H8000000F
        txtEmail.BackColor = &H8000000F
        txtPhoneExt.BackColor = &H8000000F
        txtFaxExt.BackColor = &H8000000F
    End If
End Sub


'This is called by cmdLookup and DisplayContact. The value of m_bLoading is different
'in these two contexts, hence we're saving and restoring it.

Private Sub ClearDisplay()
    Dim j As Integer
    Dim bLoading As Boolean
    
    bLoading = m_bLoading
    m_bLoading = True
    txtFirstName.text = ""
    txtLastName.text = ""
    txtEmail.text = ""
    mePhone.text = ""
    txtPhoneExt.text = ""
    meCellPhone.text = ""
    meFax.text = ""
    txtFaxExt.text = ""
    lblCreateUserID.Caption = ""
    lblCreateDate.Caption = ""
    lblUpdateUserID.Caption = ""
    lblUpdateDate.Caption = ""

    For j = 0 To chkNotify.Count - 1
        chkNotify(j) = vbUnchecked
        chkNotify(1).Enabled = False
    Next j
    cboEMailFormat.ListIndex = 2 'Default = Plain Text
    m_bLoading = bLoading
End Sub


Private Sub GetOldPhoneNbrs()
    Dim orst As ADODB.Recordset
        
    Set orst = CallSP("spcpcFindPhoneNbrs", _
                        "@CustID", m_oContacts.OwnerID, _
                        "@CustKey", m_oContacts.OwnerKey)

    Set gdxOldContacts.ADORecordset = orst
    GridColAutoSize gdxOldContacts

    Set orst = Nothing
End Sub


Private Sub LoadValidationRules()
    Dim oCtlWrapper As ControlWrapper
    
    Set m_oBrokenRules = New BrokenRules
    
    'This is only used in OP
    'm_oBrokenRules.Form = Me
 
    Set oCtlWrapper = m_oBrokenRules.AddControl(txtFirstName, "First Name", False)
    oCtlWrapper.AddRuleAlpha

    Set oCtlWrapper = m_oBrokenRules.AddControl(txtFirstName, "First Name", True)
    oCtlWrapper.AddRuleRequired "", , , "All contacts must have a first name."
    
    Set oCtlWrapper = m_oBrokenRules.AddControl(txtLastName, "Last Name", False)
    oCtlWrapper.AddRuleAlpha

    Set oCtlWrapper = m_oBrokenRules.AddControl(txtLastName, "Last Name", True)
    oCtlWrapper.AddRuleRequired "", , , "All contacts must have a last name."

    Set oCtlWrapper = m_oBrokenRules.AddControl(mePhone, "Phone Number", False)
    oCtlWrapper.AddRuleLength 10, 10

    Set oCtlWrapper = m_oBrokenRules.AddControl(mePhone, "Phone Number", True)
    oCtlWrapper.AddRuleRequired "", , , "All contacts must have a phone number."
    
    Set oCtlWrapper = m_oBrokenRules.AddControl(meFax, "Fax Number", False)
    oCtlWrapper.AddRuleLength 10, 10
        
    Set oCtlWrapper = m_oBrokenRules.AddControl(meCellPhone, "Cell Phone Number", False)
    oCtlWrapper.AddRuleLength 10, 10
    
    Set oCtlWrapper = m_oBrokenRules.AddControl(txtEmail, "Email Address", False)
    oCtlWrapper.AddRuleEMail

    'turn on the default rule class (which is 0)
    m_oBrokenRules.EnableClass 0, True
End Sub

Private Sub ChangeSelectedContact(CntctKey As String)
    m_oContacts.SelectContact CntctKey
    Set m_oContact = m_oContacts.selContact
End Sub

Private Sub ShowGrids()
    fraGrid.Visible = True
    Me.Height = 10110
    fraEdit.Top = 3840
End Sub

Private Sub LoadEmailFormatCombo()
    With cboEMailFormat
        .AddItem "HTML"
        .ItemData(.NewIndex) = 1
        
        .AddItem "RTF"
        .ItemData(.NewIndex) = 2
        
        .AddItem "Plain Text"
        .ItemData(.NewIndex) = 3
    End With
End Sub

Private Sub RunLostFocusEvent(cntrl As Control)
    Select Case cntrl.Name
        Case "txtFirstName"
           txtFirstName_LostFocus
        Case "txtLastName"
           txtLastName_LostFocus
        Case "txtEmail"
           txtEmail_LostFocus
        Case "mePhone"
            mePhone_LostFocus
        Case "txtPhoneExt"
            txtPhoneExt_LostFocus
        Case "meCellPhone"
            meCellPhone_LostFocus
        Case "meFax"
            meFax_LostFocus
        Case "txtFaxExt"
            txtFaxExt_LostFocus
    End Select
'    Set m_oPrevCntrl = Nothing
End Sub






