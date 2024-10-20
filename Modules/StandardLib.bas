Attribute VB_Name = "StandardLib"
Option Explicit

Function IsWbOpen(ByVal WbName As String) As Boolean
' Check if workbook or add-in is open.
    
    Dim Wbook As Workbook
    Dim AddInWb As Variant
    
    If InStr(WbName, ".xlam") > 0 Then
        ' For newer Excel
        If Val(Application.Version) > 12 Then
            For Each AddInWb In Application.AddIns2  ' 2016 = .AddIns2
                If AddInWb.Name = WbName Then
                    IsWbOpen = True
                    Exit For
                End If
            Next AddInWb
        Else
            ' For old Excel 2007
            On Error Resume Next
            If Workbooks(WbName).Name = WbName Then
                If Err.Number = 9 Then
                    IsWbOpen = False
                    Err.Clear
                Else
                    IsWbOpen = True
                End If
                On Error GoTo 0
                Exit Function
            End If
            On Error GoTo 0
        End If
    Else
        For Each Wbook In Workbooks
            If Wbook.Name = WbName Then
                IsWbOpen = True
                Exit For
            End If
        Next Wbook
    End If

End Function

Sub AddWorkbookVBA6()
    
    #If VBA6 Then
        If Workbooks.Count = 1 Then
            Workbooks.Add
        End If
    #End If
    
End Sub

