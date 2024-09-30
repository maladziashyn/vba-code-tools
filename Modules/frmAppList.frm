VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAppList 
   Caption         =   "Manage Applications"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8040
   OleObjectBlob   =   "frmAppList.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAppList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub bt_BrowseBackupDir_Click()
    tb_BackupDir.Text = PathFromPicker
End Sub

Private Sub bt_BrowseHomeDir_Click()
    tb_HomeDir.Text = PathFromPicker
End Sub

Private Function PathFromPicker() As String
' Update text box taking value from folder picker.

    Dim fd As FileDialog
    
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    With fd
        .AllowMultiSelect = False
'        .Title = "Pick projects home directory"
    End With
    If fd.Show = 0 Then
        Exit Function
    Else
        PathFromPicker = fd.SelectedItems(1)
    End If
    
End Function

Private Sub bt_Close_Click()
    Unload Me
End Sub

Private Sub bt_OK_Click()
    Unload Me
End Sub

Private Sub bt_SaveChanges_Click()
    
    Dim i As Long
    Dim LastR As Long
    Dim Arr As Variant
    
    If Not IsEmpty(ws1.Range("AppList").Offset(1, 0)) Then
        LastR = ws1.Cells(ws1.Rows.Count, 1).End(xlUp).Row
        ReDim Arr(1 To LastR - 1, 1 To 2)
        For i = 2 To LastR
            Arr(i - 1, 1) = ws1.Cells(i, 1)
            Arr(i - 1, 2) = ws1.Cells(i, 2)
        Next i
    End If
    
    With oAppData
        .HomeDir = tb_HomeDir.Text
        .BackupDir = tb_BackupDir.Text
        .PrintArr = Arr
    End With
    
    Call oAppData.SaveChangesToJson
    
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
