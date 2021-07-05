Attribute VB_Name = "EntryPoints"
Option Explicit

Public Type TypeLambdaRecord
    Name As String
    RefersTo As String
    Comment As String
End Type

Public Const csRepoSheetName As String = "__LambdaRepos"
Public Const csRepoListObjName As String = "__tbl_LambdaRepos"

Sub ExportLambaFunctionsFromActiveWorkbookToXml()

    Dim shtLambdaInventory As Worksheet
    Dim loLambdaInventoryList As ListObject
    Dim Lambdas() As TypeLambdaRecord
    Dim sXmlFileExportPath As String
    Dim sHumanReadableInventoryFilePath As String
    Dim wkb As Workbook
    Const cXmlMapName As String = "LambdaMap"

    StandardEntry
    Set wkb = ActiveWorkbook
    
    Set shtLambdaInventory = wkb.Sheets.Add
    Set loLambdaInventoryList = CreateLambdaXmlTable(shtLambdaInventory, cXmlMapName)
    
    ReadLambdaFormulasInWorkbook wkb, Lambdas
    PopulateLambdaInventoryList loLambdaInventoryList, Lambdas()
    
    sXmlFileExportPath = wkb.Path & Application.PathSeparator & "LambdaFunctions.xml"
    WriteXmlFile wkb, cXmlMapName, sXmlFileExportPath
    
    sHumanReadableInventoryFilePath = wkb.Path & Application.PathSeparator & "LambdaFunctions.txt"
    WriteHumanReadableLambdaInventory Lambdas, sHumanReadableInventoryFilePath
    
    shtLambdaInventory.Delete
    StandardExit

End Sub




Sub AddGitRepoToActiveWorkbook()

    Dim sRepoUrl As String
    Dim wkb As Workbook
    
    sRepoUrl = InputBox("Enter Repo URL")
    
    Set wkb = ActiveWorkbook
    If Not RepoSheetExists(wkb) Then
        CreateVeryHiddenRepoSheet wkb
    End If
    If RepoAlreadyExistsInWorkbook(wkb, sRepoUrl) Then
        MsgBox ("This repo already exists in the active workbook")
    Else
        AddLambdaRepoToWorkbook wkb, sRepoUrl
        MsgBox ("Repo successfully added")
    End If

End Sub
