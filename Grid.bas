Attribute VB_Name = "modGrid"
Option Explicit


'attach a recordset to a grid

Public Sub AttachGrid(ByRef i_oGrid As GridEX, ByRef i_oRst As ADODB.Recordset)
    With i_oGrid
        Dim i As Long
        .HoldFields
        .HoldSortSettings = True
        Set .ADORecordset = i_oRst
        For i = 1 To .Columns.Count
            .Columns(i).AutoSize
        Next
    End With
End Sub


'Added 8/5/03 LR
'Uses FFontEditor.frm
'NOTE: putting this function in this module requires any project that includes this file
'also include FFontEditor.frm. Not such a good idea.
'Maybe I should include it in the GridWrapper class (though unbound grids don't usually
'employ a gridwrapper.

Public Sub ChangeGridFont(gdx As GridEX)
    Dim oFrm As FFontEditor
    Set oFrm = New FFontEditor
    oFrm.FName = gdx.Font.Name
    oFrm.FSize = gdx.Font.Size
    oFrm.Show vbModal
    gdx.Font.Name = oFrm.FName
    gdx.ColumnHeaderFont.Name = oFrm.FName
    gdx.Font.Size = oFrm.FSize
    gdx.ColumnHeaderFont.Size = oFrm.FSize
    Unload oFrm
    Set oFrm = Nothing
End Sub
  
  
'Added 8/6/03 LR
Public Sub GetGridLayout(UserKey As Long, gdx As GridEX)
    Dim rst As ADODB.Recordset
    
    SetWaitCursor True
    Set rst = CallSP("spcpcGetUserPrefs", "@_iUserKey", UserKey, "@_iGridName", gdx.Name)
    
    If rst Is Nothing Then Exit Sub

    If Not rst.EOF Then
        gdx.LoadLayoutString rst.Fields("LayoutString").Value
    End If
    
    Set rst = Nothing
    SetWaitCursor False
End Sub


