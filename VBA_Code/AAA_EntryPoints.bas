Attribute VB_Name = "AAA_EntryPoints"
Option Explicit

Public Type TypeLambdaRecord
    RepoName As String
    LambdaName As String
    RefersTo As String
    Comment As String
End Type

Public Const csRepoStorageName As String = "GitLambdaStorage"


Sub ExportLambaFunctionsFromActiveWorkbookToXml()

    Dim shtLambdaInventory As Worksheet
    Dim LambdaStorage As zLIB_ListStorage
    Dim Lambdas() As TypeLambdaRecord
    Dim sXmlFileExportPath As String
    Dim sHumanReadableInventoryFilePath As String
    Dim wkb As Workbook
    Const cXmlMapName As String = "LambdaMap"

    StandardEntry
    Set wkb = ActiveWorkbook
    
    Set LambdaStorage = CreateLambdaXmlListStorage(wkb, cXmlMapName)
    
    ReadLambdaFormulasInWorkbook wkb, Lambdas
    PopulateLambdaInventoryStorage LambdaStorage, Lambdas()

    sXmlFileExportPath = wkb.Path & Application.PathSeparator & "LambdaFunctions.xml"
    WriteXmlFile wkb, cXmlMapName, sXmlFileExportPath

    sHumanReadableInventoryFilePath = wkb.Path & Application.PathSeparator & "LambdaFunctions.txt"
    WriteHumanReadableLambdaInventory Lambdas, sHumanReadableInventoryFilePath
    
    
    LambdaStorage.Delete
    StandardExit

End Sub




Sub AddGitRepo()

    Dim sRepoUrl As String
    Dim sRepoName As String
    Dim GitRepoStorage As zLIB_ListStorage
    Dim a As Boolean

    sRepoUrl = InputBox("Enter Repo URL")
    
    Set GitRepoStorage = New zLIB_ListStorage
    
    If Not (GitRepoStorage.StorageAlreadyExists(ThisWorkbook, csRepoStorageName)) Then
        GitRepoStorage.CreateStorage ThisWorkbook, csRepoStorageName, Array("RepoName", "RepoUrl")
    End If

'
'    If RepoAlreadyExistsInWorkbook(loRepos, sRepoUrl) Then
'        MsgBox ("This repo name already exists in the active workbook")
'    Else
'        'AddLambdaRepoToWorkbook wkb, sRepoUrl
'        MsgBox ("Repo successfully added")
'    End If


End Sub
