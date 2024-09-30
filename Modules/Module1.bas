Attribute VB_Name = "Module1"
Option Explicit

Sub Dev1()
    
    With ThisWorkbook
        .IsAddin = Not .IsAddin
    End With
    
End Sub

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

Sub parsejson1()
' "C:\home\projects\VBADevTools_DONOR\settings.json"
    
    Dim FPath As String
    Dim Json As Object
    Dim UpdatedJson As String
    
    FPath = "C:\home\projects\VBADevTools_DONOR\settings.json"
    
    Set Json = JsonConverter.ParseJson(ReadFileToString("C:\home\projects\VBADevTools_DONOR\settings.json"))
    UpdatedJson = JsonConverter.ConvertToJson(Json, Whitespace:=4)
    
    Call SaveStringToFile(UpdatedJson, FPath)
    
'Debug.Print UpdatedJson

End Sub
