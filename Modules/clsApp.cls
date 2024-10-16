VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsApp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type TApp
    Nm As String
    Ext As String
    BuildAs As String
    DoImportCode As Boolean
'    EnableEventsOn As Boolean
'    IsForBuild As Boolean
'    IsForDeployProd As Boolean
'    IsForDeployTest As Boolean
    IsOpened As Boolean
'
    HomeCell As Range
'    FName As String
'    FPath As String
'    FullNameShared As String
'    FullNameRemote As String
'    RepoUrl As String
'    DeployFName As String
'    password As String
'    InjectBuild As String
'    Launcher As String
'
'    HomeDir As String ' CodeDir
'    ShareDir As String ' SharedDir
'    RemoteDir As String ' RemoteDir
'
'    BuildTestBackup As String
'    BuildProdBackup As String
'
'    BashDir As String
'
'    Wb As Workbook
'    VBComps As VBIDE.VBComponents
'    CurrentVBComps As New Dictionary
'    FilesInCodeDir As New Dictionary
'    VBCompDirs As Variant
End Type
Private this As TApp

Private Sub Class_Initialize()

    Dim ArrTmp As Variant
    
    With this
    
        ' CHECK 1. App should be on the list
        Set .HomeCell = ws1.ListObjects(TblApps) _
                .ListColumns(AppNmCol) _
                .DataBodyRange.Find(What:=ws1.Range(RgCurrApp))
        If .HomeCell Is Nothing Then
            Err.Raise 800, "App_Initialize", "App is not on the list."
        End If
        
        .Nm = .HomeCell.Value
        .Ext = .HomeCell.Offset(0, 1).Value
        .BuildAs = .HomeCell.Offset(0, 2).Value
        .DoImportCode = .HomeCell.Offset(0, 3).Value
        .IsOpened = IsWbOpen(.Nm)
        
    
'    this.Launcher = this.HomeCell.Offset(0, 2).Value
'    this.EnableEventsOn = this.HomeCell.Offset(0, 3).Value
'    this.IsForBuild = this.HomeCell.Offset(0, 4).Value
'    this.IsForDeployTest = this.HomeCell.Offset(0, 5).Value
'    this.IsForDeployProd = this.HomeCell.Offset(0, 6).Value
'
'    this.DeployFName = this.HomeCell.Offset(0, 7).Value
'    this.Pwd = this.HomeCell.Offset(0, 8).Value
'    this.Injection = this.HomeCell.Offset(0, 9).Value
'
'    this.FName = wsMain.Range("CurrentApp").Value
'    this.IsOpened = IsWbOpen(this.FName)
'
'    ArrTmp = Split(this.FName, ".")
'    this.ShortName = ArrTmp(0)
'    this.HomeDir = wsMain.Range("CodeDir").Value & "\" & this.ShortName
'    this.FPath = this.HomeDir & "\" & this.FName
'
'    this.ShareDir = wsMain.Range("SharedDir").Value & "\" & this.ShortName
'    If Dir(this.ShareDir, vbDirectory) = "" Then
'        Call MakeDirTree(this.ShareDir)
'    End If
'    this.ShareFPath = this.ShareDir & "\" & this.FName
'
'
'
'    this.RemoteDir = ServerTest & "Attachments\DEPLOY\OtherApps\" & this.ShortName
'    this.RemoteFPath = this.RemoteDir & "\" & this.FName
'
'    this.RepoUrl = "https://github.com/maladziashyn/" & this.ShortName & ".git"
'    this.BldBackupTest = wsMain.Range("SharedDir").Value & "\BuildBackups\Test\" & this.ShortName
'    If Dir(this.BldBackupTest, vbDirectory) = "" Then
'        Call MakeDirTree(this.BldBackupTest)
'    End If
'    this.BldBackupProd = wsMain.Range("SharedDir").Value & "\BuildBackups\Prod\" & this.ShortName
'    If Dir(this.BldBackupProd, vbDirectory) = "" Then
'        Call MakeDirTree(this.BldBackupProd)
'    End If
'
'    ' Git bash needs forward-slashes to "cd"!
'    this.BashDir = Replace(this.HomeDir, "\", "/")
'
'    If this.IsOpened Then
'        Call SetVBComponents
'    End If
    
    End With

End Sub
