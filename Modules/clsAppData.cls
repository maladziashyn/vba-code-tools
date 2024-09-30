VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAppData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const AppsJson As String = "settings.json"

Private Type TAppData
    FPath As String
    Json As Object
    PrintArr As Variant
    HomeDir As String
    BackupDir As String
End Type

Private this As TAppData

Private Sub Class_Initialize()
    
    Dim i As Long
    Dim Itm As Variant
    
    With this
        .FPath = ThisWorkbook.path & "\" & AppsJson
        Set .Json = JsonConverter.ParseJson( _
            ReadFileToString(.FPath) _
        )
        .HomeDir = .Json("home_dir")
        .BackupDir = .Json("backup_dir")
        
        ReDim .PrintArr(0 To .Json("apps").Count, 1 To .Json("headers").Count)
        ' Move data from JSON to array
        ' Fill headers
        i = 1
        For Each Itm In .Json("headers")
            .PrintArr(0, i) = Itm
            i = i + 1
        Next Itm
        ' Fill rows
        i = 1
        For Each Itm In .Json("apps")
            .PrintArr(i, 1) = Itm("name")
            .PrintArr(i, 2) = Itm("ext")
            i = i + 1
        Next Itm
    End With
    
End Sub

Public Property Get HomeDir() As String
    HomeDir = this.HomeDir
End Property
Public Property Get BackupDir() As String
    BackupDir = this.BackupDir
End Property
Public Property Get PrintArr() As Variant
    PrintArr = this.PrintArr
End Property
