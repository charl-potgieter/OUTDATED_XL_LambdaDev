Attribute VB_Name = "AAA_EntryPoints"
Option Explicit

Public Type TypeLambdaRecord
    Name As String
    RefersTo As String
    Comment As String
End Type


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




'Sub AddGitRepoToActiveWorkbook()
'
'    Dim sRepoUrl As String
'    Dim wkb As Workbook
'    Dim sRepoName As String
'    Dim loRepos As ListObject
'
'    sRepoUrl = InputBox("Enter Repo URL")
'
'    Set wkb = ActiveWorkbook
'    Set loRepos = AssignReposList(wkb)
'
'    If RepoAlreadyExistsInWorkbook(loRepos, sRepoUrl) Then
'        MsgBox ("This repo name already exists in the active workbook")
'    Else
'        'AddLambdaRepoToWorkbook wkb, sRepoUrl
'        MsgBox ("Repo successfully added")
'    End If
'
'
'End Sub
