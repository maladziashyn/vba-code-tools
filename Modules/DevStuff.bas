Attribute VB_Name = "DevStuff"
Option Explicit

Sub Dev1()
    
    With ThisWorkbook
        .IsAddin = Not .IsAddin
    End With
    
End Sub

Sub d002()
    
    Dim Txt As String
    Dim FPath As String
    
    FPath = "C:\home\projects\VBACodeTools\apps.csv"
    Txt = ReadFileToString(FPath)
    
    Debug.Print Txt
    
End Sub

Sub d003()
    
    Dim FNum As Integer
    Dim i As Long
    Dim Txt As String
    Dim OneLine As String
    Dim FPath As String
    Dim Rg As Range
    
    FPath = "C:\home\projects\VBACodeTools\apps.csv"
    
    FNum = FreeFile
    Open FPath For Input As FNum
    
    i = 5
    While Not EOF(FNum)
        Line Input #FNum, OneLine
        ws1.Cells(i, 1).Value = OneLine
        i = i + 1
'Debug.Print OneLine
    Wend
    
    Close #FNum
    
    Set Rg = ws1.Cells(5, 1).CurrentRegion
    Rg.TextToColumns Comma:=True
    
End Sub
