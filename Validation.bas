Attribute VB_Name = "modValidation"
Option Explicit

'To reset control back colors
Public Const k_lCtlBackColor = VBRUN.SystemColorConstants.vbWindowBackground
Public Const k_lCtlWhiteColor = vbWhite
Public Const k_lCtlForeColor = VBRUN.SystemColorConstants.vbWindowText
Public Const k_lCtlLockColor = VBRUN.ColorConstants.vbBlue
'Public Const k_lCtlMarkColor = &HFFFF00
Public Const k_lCtlMarkColor = &HC0C0FF
 

Public Function NextValidationRuleID() As Long
    Static lNextID As Long
    
    lNextID = lNextID + 1
    NextValidationRuleID = lNextID
End Function


Public Sub ReportError()

End Sub
