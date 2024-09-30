Attribute VB_Name = "DevStuff"
Option Explicit

Sub Dev1()
    
    With ThisWorkbook
        .IsAddin = Not .IsAddin
    End With
    
End Sub
