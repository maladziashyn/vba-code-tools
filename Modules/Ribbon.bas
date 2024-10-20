Attribute VB_Name = "Ribbon"
Option Explicit

Public Const MsbTitle As String = "VBA Code Tools"

Public rbxUI_VCT As IRibbonUI

Sub VBACodeTools_onLoad(ByRef Ribbon As IRibbonUI)
' Load the custom ribbon tab.

    Set rbxUI_VCT = Ribbon
    #If VBA7 Then ' for 2010 and newer
        rbxUI_VCT.ActivateTab "tabVBACodeTools"
    #End If
    Application.WindowState = xlMaximized
    
End Sub

Sub VBACodeTools_ClickButton(ByRef control As IRibbonControl)

    Select Case control.ID
        Case "btSelectApp"
            frmSelectApp.Show
        Case "itmShowHideSettings"
            Call ShowHideSettings
        Case "btExit"
            ThisWorkbook.Close SaveChanges:=False
        
    End Select
    
End Sub

'Sub VBACodeTools_ClickButton_WithGetPressed(ByRef control As IRibbonControl, _
'        ByRef pressed As Boolean)
'
'    Call ShowHideSettings(pressed)
'
'End Sub

Private Sub ShowHideSettings()
' Turn AddIn mode on/off.
' Turn it off to make change on add-in's worksheets.
    
    With ThisWorkbook
        Call AddWorkbookVBA6
        .IsAddin = Not .IsAddin
    End With

    ' Refresh ribbon
    If rbxUI_VCT Is Nothing Then
        MsgBox "Error: Custom ribbon tab was reset because of an error. " _
            & "Restart Code Tools to restore the tab.", vbCritical, MsbTitle
    Else
        rbxUI_VCT.Invalidate
    End If


End Sub

Sub VBACodeTools_GetEnabled(ByRef control As IRibbonControl, ByRef returnedVal)
    
    returnedVal = Not ThisWorkbook.IsAddin
    
End Sub

Sub VBACodeTools_GetLabel(ByRef control As IRibbonControl, ByRef returnedVal)
    
    Select Case control.ID
        Case "btSelectApp"
            returnedVal = ws1.Range("CurrentApp").Value
        Case "itmShowHideSettings"
            returnedVal = IIf(ThisWorkbook.IsAddin, "Hide Settings", "Show Settings")
        Case Else
            returnedVal = "Unknown..."
    End Select
    
End Sub
