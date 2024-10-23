Attribute VB_Name = "StandardLib"
Option Explicit

Sub AddWorkbookVBA6(Optional ByVal MinBooks As Long = 1)
' In Excel 2007, custom ribbon won't work without a workbook open.

    #If VBA6 Then
        If Workbooks.Count <= MinBooks Then ' 1 is PERSONAL workbook
            Workbooks.Add
        End If
    #End If
    
End Sub

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

Function SetRangeFromArrDimensions(ByRef Ws As Worksheet, _
        ByRef UpLeftC As Range, _
        ByVal Arr As Variant, _
        Optional ByVal IsOneDim As Boolean = False) As Range
' Set range with size like the passed in Arr.
' 2-dim array should be dimensioned as Rows-Columns
    
    Dim AddLBoundRows As Long
    Dim AddLBoundCols As Long
    Dim Rg As Range
    
    AddLBoundRows = IIf(LBound(Arr, 1) = 0, 1, 0)
    If Not IsOneDim Then
        AddLBoundCols = IIf(LBound(Arr, 2) = 0, 1, 0)
    End If
    
    Set Rg = Ws.Cells
    If Not IsOneDim Then
        Set SetRangeFromArrDimensions = Ws.Range(UpLeftC, Rg( _
            UpLeftC.Row + UBound(Arr, 1) - 1 + AddLBoundRows, _
            UpLeftC.Column + UBound(Arr, 2) - 1 + AddLBoundCols))
    Else
        Set SetRangeFromArrDimensions = Ws.Range(UpLeftC, Rg( _
            UpLeftC.Row + UBound(Arr) - 1 + AddLBoundRows, _
            UpLeftC.Column))
    End If
    
End Function

Sub SaveStringToFile(ByVal PrintText As String, _
        ByVal FPath As String, _
        Optional ByVal ToOpen As Boolean = False)
    
    Dim FNum As Integer
    
    FNum = FreeFile
    Open FPath For Output As FNum ' alternatively - For Append
    Print #FNum, PrintText
    Close #FNum
    
    If ToOpen Then
        Shell """C:\Program Files\Notepad++\notepad++.exe"" """ & FPath & """", vbNormalFocus
    End If

End Sub

Function ReadFileToString(ByVal FPath As String) As String
    
    Dim FNum As Integer
    
    FNum = FreeFile
    Open FPath For Input As #FNum
    ReadFileToString = input(LOF(FNum), FNum)
    Close FNum
    
End Function
