Attribute VB_Name = "Main"
Option Explicit

Sub ExportSettings()
    
    Dim i As Long, j As Long
    Dim OutJson As String
    Dim TbRows As String
    Dim Itm As Variant
    Dim Rg As Range
    Dim OutD As New Dictionary
    
    Set Rg = ws1.ListObjects(TblApps).DataBodyRange
    For i = 1 To Rg.Rows.Count
        For j = 1 To Rg.Columns.Count
            TbRows = TbRows & Rg(i, j) & JsonSep
        Next j
        TbRows = TbRows & vbLf
    Next i
    
    With OutD
        .Add "TableRows", Left(TbRows, Len(TbRows) - 1)
        For Each Itm In Array(RgHomeDir, RgBackupDir, RgCurrApp, RgFormListIndex)
            .Add Itm, ws1.Range(Itm)
        Next Itm
    End With
    
    OutJson = JsonConverter.ConvertToJson(JsonValue:=OutD, Whitespace:=" ", json_CurrentIndentation:=4)
    Call SaveStringToFile(OutJson, ThisWorkbook.Path & "\" & JsonFName)
    MsgBox "Settings have been exported.", vbInformation, MsbTitle

End Sub

Sub ImportSettings()
    
    Dim i As Long
    Dim InJson As String
    Dim TbRows As Variant
    Dim Itm As Variant
    Dim Cell As Range
    Dim Rg As Range
    Dim InD As Object
    
    ' Clear table
    For i = ws1.ListObjects(TblApps).ListRows.Count To 1 Step -1
        ws1.ListObjects(TblApps).ListRows(i).Delete
    Next i
    
    ' Parse json into InD
    InJson = ReadFileToString(ThisWorkbook.Path & "\" & JsonFName)
    Set InD = JsonConverter.ParseJson(InJson)
    
    ' Print json data
    TbRows = Split(InD("TableRows"), vbLf, -1, vbTextCompare)
    TbRows = WorksheetFunction.Transpose(TbRows)
    Set Rg = SetRangeFromArrDimensions(ws1, ws1.Cells(2, 1), TbRows, True)
    Rg = TbRows
    
    Application.DisplayAlerts = False
    For Each Cell In Rg
        ' Split text to columns
        Cell.TextToColumns Semicolon:=True
    Next Cell
    Application.DisplayAlerts = True
    
    For Each Itm In Array(RgHomeDir, RgBackupDir, RgCurrApp, RgFormListIndex)
        ws1.Range(Itm) = InD(Itm)
    Next Itm
    MsgBox "Settings have been imported.", vbInformation, MsbTitle
    
End Sub

Sub QuitApp()
    
    If MsgBox("Are you sure you want to quit?", _
            vbQuestion + vbOKCancel + vbDefaultButton2, _
            MsbTitle) = vbOK Then
        ThisWorkbook.Close 'SaveChanges:=False
    End If
    
End Sub
