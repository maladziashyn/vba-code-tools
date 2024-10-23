Attribute VB_Name = "Ribbon"
Option Explicit

Public rbxUI_VCT As IRibbonUI

Sub VBACodeTools_onLoad(ByRef Ribbon As IRibbonUI)
' Load the custom ribbon tab.
    
    Call AddWorkbookVBA6(1)
    
    Set rbxUI_VCT = Ribbon
    #If VBA7 Then ' for 2010 and newer
        rbxUI_VCT.ActivateTab MainRibbonTab
    #End If
    Application.WindowState = xlMaximized
    
End Sub

Sub VBACodeTools_ClickButton(ByRef control As IRibbonControl)

    Select Case control.ID
        Case "btSelectApp"
            frmSelectApp.Show
        Case "itmShowHideSettings"
            Call ShowHideSettings
        Case "itmExportSettings"
            Call ExportSettings
        Case "itmImportSettings"
            Call ImportSettings
        Case "btExit"
            Call QuitApp
    End Select
    
End Sub

Private Sub ShowHideSettings()
' Turn AddIn mode on/off.
' Turn it off to make change on add-in's worksheets.
    
    With ThisWorkbook
        .IsAddin = Not .IsAddin
        If .IsAddin Then
            .Save ' save on switching from addin to regular mode
        End If
    End With
    Call AddWorkbookVBA6
    
    ' Refresh ribbon
    If rbxUI_VCT Is Nothing Then
        MsgBox "Error: Custom ribbon tab was reset because of an error. " _
            & "Restart Code Tools to restore the tab.", vbCritical, MsbTitle
    Else
        rbxUI_VCT.Invalidate
    End If


End Sub

Sub VBACodeTools_GetEnabled(ByRef control As IRibbonControl, ByRef returnedVal)
    
    If control.Tag = "IsAddIn" Then
        returnedVal = Not ThisWorkbook.IsAddin
    Else
        returnedVal = ThisWorkbook.IsAddin
    End If
    
End Sub

Sub VBACodeTools_GetLabel(ByRef control As IRibbonControl, ByRef returnedVal)
    
    Select Case control.ID
        Case "btSelectApp"
            returnedVal = ws1.Range(RgCurrApp).Value
        Case "itmShowHideSettings"
            returnedVal = IIf(ThisWorkbook.IsAddin, "Show Settings", "Hide Settings")
        Case Else
            returnedVal = "Unknown..."
    End Select
    
End Sub

Sub VBACodeTools_GetImage(ByRef control As IRibbonControl, ByRef returnedVal)
    
    returnedVal = IIf(ThisWorkbook.IsAddin, "WindowUnhide", "WindowHide")
    
End Sub
