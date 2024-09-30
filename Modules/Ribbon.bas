Attribute VB_Name = "Ribbon"
Option Explicit

Public Const MsbTitle As String = "VBA Code Tools"

Public rbxUI_VCT As IRibbonUI

Sub VBACodeTools_onLoad(ByRef Ribbon As IRibbonUI)
' Load the custom ribbon tab.

    Set rbxUI_VCT = Ribbon
'    rbxUI_VCT.ActivateTab "tabVBACodeTools" ' for 2010 and newer
    Application.WindowState = xlMaximized
    
End Sub

Sub VBACodeTools_ClickButton(ByRef control As IRibbonControl)

    Select Case control.ID
        Case "btExit"
            ThisWorkbook.Close savechanges:=False
        
    End Select
    
End Sub

Sub VBACodeTools_ClickButton_WithGetPressed(ByRef control As IRibbonControl, _
        ByRef pressed As Boolean)

    ' Refresh ribbon
    If rbxUI_VCT Is Nothing Then
        MsgBox "Error: Custom ribbon tab was reset because of an error. " _
            & "Restart Code Tools to restore the tab.", vbCritical, MsbTitle
    Else
        rbxUI_VCT.Invalidate
    End If
    
End Sub

Sub VBACodeTools_GetPressed(ByRef control As IRibbonControl, ByRef returnedVal)
    
    returnedVal = True
        
End Sub

Sub VBACodeTools_GetEnabled(ByRef control As IRibbonControl, ByRef returnedVal)
    
    returnedVal = True
    
End Sub

Sub VBACodeTools_GetVisible(ByRef control As IRibbonControl, ByRef returnedVal)
    
    returnedVal = True
    
End Sub

Sub VBACodeTools_GetLabel(ByRef control As IRibbonControl, ByRef returnedVal)
    
    returnedVal = "Unknown..."
    
End Sub
