VERSION 5.00
Begin VB.Form FXRef 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cross Reference"
   ClientHeight    =   6090
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   10440
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   10440
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4560
      TabIndex        =   2
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      Caption         =   "About the Relationship"
      Height          =   1695
      Left            =   120
      TabIndex        =   15
      Top             =   3840
      Width           =   10215
      Begin VB.CheckBox chkDirectReplacement 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   255
         Left            =   9840
         TabIndex        =   16
         Top             =   300
         Width           =   255
      End
      Begin VB.Label lblHiddenXRefType 
         AutoSize        =   -1  'True
         Caption         =   "lblHiddenXRefType"
         Height          =   195
         Left            =   6480
         TabIndex        =   29
         Top             =   360
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.Label lblUpdateInfo 
         AutoSize        =   -1  'True
         Caption         =   "Updated By:"
         Height          =   195
         Left            =   240
         TabIndex        =   20
         Top             =   360
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Notes:"
         Height          =   195
         Index           =   15
         Left            =   240
         TabIndex        =   19
         Top             =   720
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Is Direct Replacement:"
         Height          =   195
         Index           =   17
         Left            =   8160
         TabIndex        =   18
         Top             =   360
         Width           =   1620
      End
      Begin VB.Label lblNote 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000002&
         Height          =   855
         Left            =   1320
         TabIndex        =   17
         Top             =   720
         Width           =   8775
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3735
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   10215
      Begin VB.TextBox txtXRef 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1935
         Left            =   5760
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   840
         Width           =   4335
      End
      Begin VB.TextBox txtInput 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   2295
         Left            =   1200
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   480
         Width           =   4335
      End
      Begin VB.ComboBox cboXRef 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   5760
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   480
         Width           =   4335
      End
      Begin VB.CommandButton cmdSelectRef 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5760
         MaskColor       =   &H80000002&
         TabIndex        =   1
         Top             =   3120
         Width           =   4335
      End
      Begin VB.CommandButton cmdSelectInput 
         BackColor       =   &H80000002&
         Caption         =   "Not Available"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         TabIndex        =   0
         Top             =   3120
         Width           =   4335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Price:"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   32
         Top             =   2430
         Width           =   405
      End
      Begin VB.Label lblNotAvailable2 
         AutoSize        =   -1  'True
         Caption         =   "This part# no longer available for purchase"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   5760
         TabIndex        =   28
         Top             =   2820
         Visible         =   0   'False
         Width           =   4470
      End
      Begin VB.Label lblNotAvailable1 
         AutoSize        =   -1  'True
         Caption         =   "This part# no longer available for purchase"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   1200
         TabIndex        =   27
         Top             =   2820
         Visible         =   0   'False
         Width           =   4470
      End
      Begin VB.Label lblRelatedParts 
         AutoSize        =   -1  'True
         Caption         =   "Related Parts:"
         Height          =   195
         Left            =   5760
         TabIndex        =   26
         Top             =   240
         Width           =   1005
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "You Entered"
         Height          =   195
         Index           =   0
         Left            =   1200
         TabIndex        =   25
         Top             =   240
         Width           =   885
      End
      Begin VB.Label lblRefAvailable 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000002&
         Height          =   315
         Left            =   5760
         TabIndex        =   24
         Top             =   1920
         Width           =   4335
      End
      Begin VB.Label lblRefQOH 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000002&
         Height          =   315
         Left            =   5760
         TabIndex        =   23
         Top             =   1560
         Width           =   4335
      End
      Begin VB.Label lblRefPartVendor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000002&
         Height          =   315
         Left            =   5760
         TabIndex        =   22
         Top             =   1200
         Width           =   4335
      End
      Begin VB.Label lblRefPartDescr 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000002&
         Height          =   315
         Left            =   5760
         TabIndex        =   21
         Top             =   840
         Width           =   4335
      End
      Begin VB.Label lblInputQOH 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000002&
         Height          =   315
         Left            =   1200
         TabIndex        =   14
         Top             =   1560
         Width           =   4335
      End
      Begin VB.Label lblInputAvailable 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000002&
         Height          =   315
         Left            =   1200
         TabIndex        =   13
         Top             =   1920
         Width           =   4335
      End
      Begin VB.Label lblInputVendor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000002&
         Height          =   315
         Left            =   1200
         TabIndex        =   12
         Top             =   1200
         Width           =   4335
      End
      Begin VB.Label lblInputDescr 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000002&
         Height          =   315
         Left            =   1200
         TabIndex        =   11
         Top             =   840
         Width           =   4335
      End
      Begin VB.Label lblInputPartNbr 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000002&
         Height          =   315
         Left            =   1200
         TabIndex        =   10
         Top             =   480
         Width           =   4335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Available:"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   9
         Top             =   2040
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Qty On Hand:"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   8
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Vendor:"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Descr:"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "PartNbr:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   585
      End
   End
End
Attribute VB_Name = "FXRef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' SPEC
' 1. Exclude related part from selection, if it is NOT Active and Qty On Hand <= 0
' 2. Do not allow user to select a part that is NOT Active and Qty Available <= 0

Option Explicit

Private Enum enumXRefColumn
    xrcItemID
    xrcIsDirect
End Enum

Private mlItemKey As Long
Private mlItemType As ItemTypeCode
Private msOriginalItemID As String
Private msRefSource As String
Private mbCancelSearch As Boolean

Private mlReferencedItemKey As Long
Private mlReferencedItemType As ItemTypeCode
Private mlWhseKey As Long
Private myCustType As Byte


' Public Methods

Public Sub XRefSearch(ByVal sXRefPartNbr As String, _
                        ByRef lItemKey As Long, _
                        ByRef eItemType As ItemTypeCode, _
                        ByRef o_sOriginalItemID As String, _
                        ByRef o_sRefSource As String, _
                        ByRef o_bCancelSearch As Boolean, _
                        ByVal lWhseKey As Long, _
                        ByVal bDoNotShowIfNoXRef As Boolean, _
                        ByVal yCustType As Byte)
    Dim orst As ADODB.Recordset
    
    On Error GoTo EH
    
    sXRefPartNbr = Trim(sXRefPartNbr)
    mlItemKey = lItemKey
    mlItemType = eItemType
    myCustType = yCustType
    If lWhseKey <> 0 Then
        mlWhseKey = lWhseKey 'GetUserWhseKey
    Else
        mlWhseKey = GetUserWhseKey
    End If
    
    Set orst = CallSP("cpoaGetXRef", "@WhseKey", mlWhseKey, "@ItemID", sXRefPartNbr, "@CustType", myCustType)
    
    ' Referenced parts resultset
    With orst
        If .EOF Then
            If bDoNotShowIfNoXRef Then
                ' No reference found
                Unload Me
                Exit Sub
            End If
        End If
        
        ' Load referenced parts
        Do Until .EOF
            cboXRef.AddItem (.Fields("RefPart") & IIf(.Fields("IsDirect") <> vbFalse, " (Direct)", ""))
            cboXRef.ItemData(cboXRef.NewIndex) = .Fields("XRefKey")
            .MoveNext
        Loop
        If cboXRef.ListCount > 1 Then
            lblRelatedParts.Caption = cboXRef.ListCount & " related parts:"
            lblRelatedParts.ForeColor = vbRed
        Else
            lblRelatedParts.Caption = cboXRef.ListCount & " related part:"
            txtXRef.Top = txtInput.Top
            txtXRef.Height = txtInput.Height
            cboXRef.TabStop = False
        End If
        Set orst = .NextRecordset
    End With
    
    ' Detail info on the input part
    lblInputPartNbr = sXRefPartNbr
    txtInput.text = sXRefPartNbr
    With orst
        If Not .EOF Then
            txtInput.text = txtInput.text & vbCrLf & vbCrLf & _
                            .Fields("ShortDesc") & vbCrLf & vbCrLf & _
                            .Fields("VendName") & vbCrLf & vbCrLf
            lblInputDescr = .Fields("ShortDesc") & ""
            lblInputVendor = .Fields("VendName") & ""
            lblNotAvailable1.Visible = (.Fields("Status") & "" <> "Active")
            ' Due to the use of the Convert function in the stored proc, qty is returned
            ' with trailing zeroes
            If IsNumeric(.Fields("QtyOnHand") & "") Then
                lblInputQOH = CLng(.Fields("QtyOnHand"))
                txtInput.text = txtInput.text & CLng(.Fields("QtyOnHand")) & vbCrLf & vbCrLf
            Else
                lblInputQOH = .Fields("QtyOnHand") & ""
                txtInput.text = txtInput.text & .Fields("QtyOnHand") & "" & vbCrLf & vbCrLf
            End If
            If IsNumeric(.Fields("QtyAvailable") & "") Then
                lblInputAvailable = CLng(.Fields("QtyAvailable"))
                txtInput.text = txtInput.text & CLng(.Fields("QtyAvailable")) & vbCrLf & vbCrLf
            Else
                lblInputAvailable = .Fields("QtyAvailable") & ""
                txtInput.text = txtInput.text & .Fields("QtyAvailable") & "" & vbCrLf & vbCrLf
            End If
            
            txtInput.text = txtInput.text & FormatCurrency(.Fields("EffectivePrice"), 2) & "" & vbCrLf & vbCrLf
            
            ' Can't select a part if Qty Available <= 0 and not Active
            If IsNumeric(lblInputAvailable.Caption) Then
                If CLng(lblInputAvailable.Caption) <= 0 And .Fields("Status") & "" <> "Active" Then
                    cmdSelectInput.Caption = "Not Available"
                    cmdSelectInput.Enabled = False
                Else
                    If cboXRef.ListCount = 0 Then
                        ' No reference found
                        Unload Me
                        Exit Sub
                    End If
                    cmdSelectInput.Caption = "Select " & lblInputPartNbr
                    cmdSelectInput.Enabled = True
                End If
            Else
                If cboXRef.ListCount = 0 Then
                    ' No reference found
                    Unload Me
                    Exit Sub
                End If
                ' Qty Available can't be determined, we have to let the user make the
                ' decision themselves whether to use the reference
                cmdSelectInput.Caption = "Select " & lblInputPartNbr
                cmdSelectInput.Enabled = True
            End If
        Else
            If cboXRef.ListCount = 0 Then
                ' No reference found
                Unload Me
                Exit Sub
            End If
        End If
    End With
    Set orst = Nothing
    
    If cboXRef.ListCount > 0 Then
        cboXRef.ListIndex = 0
    End If
    
    Show vbModal
    
    ' Set the ByRef parameters to the values selected in XRef
    If Not mbCancelSearch Then
        lItemKey = mlItemKey
        eItemType = ConvertSageItemType(mlItemType)
        o_sOriginalItemID = msOriginalItemID
        o_sRefSource = msRefSource
    End If
    o_bCancelSearch = mbCancelSearch
    
    Unload Me
    Exit Sub
EH:
    Set orst = Nothing
    MsgBox "Failed to perform cross reference search due to error '" & Err.Number & " " & Err.Description & "'", vbInformation
    Unload Me
End Sub

' Control Event Handlers

Private Sub cboXRef_Click()
    Dim orst As ADODB.Recordset
    Dim sItemID As String
    
    On Error GoTo EH
    
    sItemID = Replace(cboXRef.text, " (Direct)", "")
    
    Set orst = CallSP("cpoaGetXRefRelationship", "@WhseKey", mlWhseKey, "@ItemID", sItemID, "@XRefKey", cboXRef.ItemData(cboXRef.ListIndex), "@CustType", myCustType)
    
    ' Detail info on the referenced part
    With orst
        If Not .EOF Then
            ' Reset display
            txtXRef.text = ""
            
            mlReferencedItemKey = .Fields("ItemKey")
            mlReferencedItemType = .Fields("ItemType")
            
            If cboXRef.ListCount = 1 Then
                ' Only 1 reference, lets put all reference info in the Reference text box for more clarity
                txtXRef.text = cboXRef.text & vbCrLf & vbCrLf
            End If
            
            txtXRef.text = txtXRef.text & .Fields("ShortDesc") & vbCrLf & vbCrLf & _
                            .Fields("VendName") & vbCrLf & vbCrLf
            
            lblRefPartDescr = .Fields("ShortDesc") & ""
            lblRefPartVendor = .Fields("VendName") & ""
            lblNotAvailable2.Visible = (.Fields("Status") & "" <> "Active")
            ' Due to the use of the Convert function in the stored proc, qty is returned
            ' with trailing zeroes
            If IsNumeric(.Fields("QtyOnHand") & "") Then
                lblRefQOH = CLng(.Fields("QtyOnHand"))
                txtXRef.text = txtXRef.text & CLng(.Fields("QtyOnHand")) & vbCrLf & vbCrLf
            Else
                lblRefQOH = .Fields("QtyOnHand") & ""
                txtXRef.text = txtXRef.text & .Fields("QtyOnHand") & "" & vbCrLf & vbCrLf
            End If
            If IsNumeric(.Fields("QtyAvailable") & "") Then
                lblRefAvailable = CLng(.Fields("QtyAvailable"))
                txtXRef.text = txtXRef.text & CLng(.Fields("QtyAvailable")) & vbCrLf & vbCrLf
            Else
                lblRefAvailable = .Fields("QtyAvailable") & ""
                txtXRef.text = txtXRef.text & .Fields("QtyAvailable") & vbCrLf & vbCrLf
            End If
            
            txtXRef.text = txtXRef.text & FormatCurrency(.Fields("EffectivePrice"), 2) & "" & vbCrLf & vbCrLf
            
            ' Can't select a part if Qty Available <= 0 and not Active
            If IsNumeric(lblRefAvailable.Caption) Then
                If CLng(lblRefAvailable.Caption) <= 0 And .Fields("Status") & "" <> "Active" Then
                    cmdSelectRef.Caption = "Not Available"
                    cmdSelectRef.Enabled = False
                Else
                    cmdSelectRef.Caption = "Select " & sItemID
                    cmdSelectRef.Enabled = True
                End If
            Else
                ' Qty Available can't be determined, we have to let the user make the
                ' decision themselves whether to use the reference
                cmdSelectRef.Caption = "Select " & sItemID
                cmdSelectRef.Enabled = True
            End If
        Else
            mlReferencedItemKey = 0
            mlReferencedItemType = 0
            
            ' No data, blank out controls
            lblRefPartDescr = ""
            lblRefPartVendor = ""
            lblRefQOH = ""
            lblRefAvailable = ""
            lblNotAvailable2.Visible = False
            cmdSelectRef.Caption = "Not Available"
            cmdSelectRef.Enabled = False
        End If
        Set orst = .NextRecordset
    End With
    
    ' Relationship info
    With orst
        If Not .EOF Then
            lblUpdateInfo.Caption = "Created by " & .Fields("UpdatedBy") & " on " & .Fields("UpdatedDate") & ", " & .Fields("Note") & ""
            lblHiddenXRefType.Caption = .Fields("XRefTypeDesc")
            If .Fields("IsDirect") = vbFalse Then
                chkDirectReplacement.Value = vbUnchecked
            Else
                chkDirectReplacement.Value = vbChecked
            End If
            lblNote = .Fields("Remark") & ""
        End If
    End With
    
    Set orst = Nothing
    
    Exit Sub
EH:
    Set orst = Nothing
    MsgBox "Failed to load cross reference relationship due to error '" & Err.Number & " " & Err.Description & "'", vbInformation
    Unload Me
End Sub


Private Sub cmdSelectInput_Click()
    Hide
End Sub


Private Sub cmdSelectRef_Click()
    mlItemKey = mlReferencedItemKey
    mlItemType = mlReferencedItemType
    msOriginalItemID = lblInputPartNbr.Caption
    msRefSource = lblHiddenXRefType.Caption
    Hide
End Sub


Private Sub cmdCancel_Click()
    mbCancelSearch = True
    Hide
End Sub


