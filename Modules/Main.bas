Attribute VB_Name = "Main"
Option Explicit

Private Function SetRangeFromArrDimensions(ByRef Ws As Worksheet, _
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

