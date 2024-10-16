VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSelectApp 
   Caption         =   "UserForm1"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3705
   OleObjectBlob   =   "frmSelectApp.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSelectApp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnOK_Click()
    
    ws1.Range("CurrentApp").Value = lbxApps.Value
    Set oApp = New clsApp
    rbxUI_VCT.Invalidate
    Unload frmSelectApp
    
End Sub

Private Sub UserForm_Initialize()
    
    frmSelectApp.Caption = MsbTitle
    With lbxApps
        .List = Application.WorksheetFunction.Transpose( _
            ws1.ListObjects("tblApps") _
                .ListColumns("Name") _
                .DataBodyRange)
        .ListIndex = ws1.Range("frmListIndex").Value
    End With
    
End Sub
