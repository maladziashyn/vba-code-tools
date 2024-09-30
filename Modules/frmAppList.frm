VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAppList 
   Caption         =   "UserForm1"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10440
   OleObjectBlob   =   "frmAppList.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAppList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub bt_Cancel_Click()
    Unload Me
End Sub

Private Sub bt_OK_Click()
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    
    Dim Rg As Range
    
    With Me
        .tb_HomeDir.Text = oAppData.HomeDir
        .tb_BackupDir.Text = oAppData.BackupDir
    End With
    
    Set Rg = ws1.Range(AppRg).CurrentRegion
    Set Rg = Rg.Resize(Rg.Rows.Count - 1, Rg.Columns.Count).Offset(1, 0)
    
    With Me.lsb_Apps
        .RowSource = ws1.Name & "!" & Rg.Address
        .ColumnHeads = True
        .ColumnCount = Rg.Columns.Count
        .ColumnWidths = ListboxColumnWidths(Rg)
        .ListIndex = 0
    End With
    
End Sub

Private Function ListboxColumnWidths(ByRef DataRg As Range) As String
' Return string of column widths.
' Columns may loop a bit wider than needed, but still a good option.
    
    Dim i As Long
    Dim ColWd As Long
    Dim WdArr As String
    
    For i = 1 To DataRg.Columns.Count
        ColWd = DataRg.Columns(i).Width
        WdArr = WdArr & ColWd & ";"
    Next i
    ListboxColumnWidths = Left(WdArr, Len(WdArr) - 1)
    
End Function
