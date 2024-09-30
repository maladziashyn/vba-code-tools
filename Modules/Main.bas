Attribute VB_Name = "Main"
Option Explicit

Public Const AppsJson As String = "settings.json"
Public Const AppRg As String = "AppList"

Public oAppData As clsAppData

Sub LoadJsonToWs1()
    
    Dim i As Long
    Dim PrintRg As Range
    
    Set oAppData = New clsAppData ' grab data from JSON
    
    ' Clean up range for future printing
    If Not IsEmpty(ws1.Range(AppRg)) Then
        ws1.Range(AppRg).CurrentRegion.Clear
    End If
    
    ' Set address for print range
    Set PrintRg = SetRangeFromArrDimensions(ws1, _
        ws1.Range(AppRg), oAppData.PrintArr)
    
    PrintRg = oAppData.PrintArr ' PRINT OUT
    
    PrintRg.AutoFilter
    With ws1.AutoFilter.Sort
        With .SortFields
            .Clear
            .Add Key:=Range(PrintRg.Columns(1).Address) ', SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        End With
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    ' Prettify
    With PrintRg
        .Font.Name = "Arial"
        .Font.Size = 13
        .Columns.AutoFit
    End With
    
    frmAppList.Show
    
End Sub

Private Function SetRangeFromArrDimensions(ByRef Ws As Worksheet, _
        ByRef UpLeftC As Range, _
        ByVal Arr As Variant) As Range
' Set range with size like the passed in Arr.
' Array should be dimensioned as Rows-Columns
    
    Dim AddLBoundRows As Long
    Dim AddLBoundCols As Long
    Dim Rg As Range
    
    AddLBoundRows = IIf(LBound(Arr, 1) = 0, 1, 0)
    AddLBoundCols = IIf(LBound(Arr, 2) = 0, 1, 0)
    
    Set Rg = Ws.Cells
    Set SetRangeFromArrDimensions = Ws.Range(UpLeftC, Rg( _
        UpLeftC.Row + UBound(Arr, 1) - 1 + AddLBoundRows, _
        UpLeftC.Column + UBound(Arr, 2) - 1 + AddLBoundCols))
    
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
    Open FPath For Input As FNum
    ReadFileToString = input(LOF(FNum), FNum)
    Close FNum
    
End Function

