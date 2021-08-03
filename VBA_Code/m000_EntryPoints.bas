Attribute VB_Name = "m000_EntryPoints"
Option Explicit

Public Const csRepoStorageName = "GitRepos"
Public Const csFunctionStorageName = "PowerFunctions"
Const csLambdaXmlMapName As String = "LambdaMap"


Public Sub CreateLambdaXmlGeneratorWorkbook()

    Dim LambdaStorage As zLIB_ListStorage
    Dim CategoryStorage As zLIB_ListStorage
    Dim wkb As Workbook
    Dim sht As Worksheet
    Dim LambdaXmlMap As XmlMap
    Const csCategoryStorageName As String = "CategoryStorage"
    Const csLambdaStorageName As String = "LambdaStorage"
    
    
    
    Set LambdaStorage = New zLIB_ListStorage
    Set CategoryStorage = New zLIB_ListStorage
    
    StandardEntry
    Set wkb = Workbooks.Add
    LambdaStorage.CreateStorage wkb, csLambdaStorageName, _
        Array("Name", "RefersTo", "Category", "Author", "Comment")
    CategoryStorage.CreateStorage wkb, csCategoryStorageName, Array("Categories")
    
    'Delete sheets other than above storage
    For Each sht In wkb.Worksheets
        If sht.Name <> csCategoryStorageName And sht.Name <> csLambdaStorageName Then
            sht.Delete
        End If
    Next sht

    FormatLambdaStorageSheet LambdaStorage
    FormatCategoryStorageSheet CategoryStorage
    Set LambdaXmlMap = CreateLambdaXmlMap(wkb, csLambdaXmlMapName)
    AssignXmlMapToStorage LambdaStorage, LambdaXmlMap

    wkb.Activate
    wkb.Sheets(1).Select
    ActiveWindow.WindowState = xlMaximized
    StandardExit
    
End Sub



Sub ExportLambaFunctionsFromActiveWorkbookToXml()

    Dim shtLambdaInventory As Worksheet
    Dim LambdaStorage As zLIB_ListStorage
    Dim sXmlFileExportPath As String
    Dim sHumanReadableInventoryFilePath As String
    Dim wkb As Workbook
    Dim sExportPath As String
    Const csXmlMapName As String = "LambdaMap"

    StandardEntry
    Set wkb = ActiveWorkbook

    If Not WorkbookIsValidForLambdaXmlExport(wkb) Then
        Exit Sub
    End If
    
    sExportPath = wkb.Path & Application.PathSeparator & "PowerFunctionExports"
    If Not FolderExists(sExportPath) Then CreateFolder (sExportPath)
    sXmlFileExportPath = sExportPath & Application.PathSeparator & "LambdaFunctions.xml"
    wkb.XmlMaps(csXmlMapName).Export Url:=sXmlFileExportPath, Overwrite:=True

    Set LambdaStorage = New zLIB_ListStorage
    LambdaStorage.AssignStorage wkb, "LambdaStorage"
    sHumanReadableInventoryFilePath = sExportPath & Application.PathSeparator & "LambdaFunctions.txt"
    WriteHumanReadableLambdaInventory LambdaStorage, sHumanReadableInventoryFilePath

    MsgBox ("Functions exported")

    StandardExit

End Sub



Sub AddGitRepoToActiveWorkbook()

    Dim sRepoUrl As String
    Dim sRepoName As String
    Dim GitRepoStorage As zLIB_ListStorage
    Dim RepoUrlDictionary As Dictionary
    Dim wkb As Workbook

    Set wkb = ActiveWorkbook
    sRepoUrl = InputBox("Enter Repo URL")
    If sRepoUrl = "" Then Exit Sub

    Set GitRepoStorage = New zLIB_ListStorage
    

    If Not (GitRepoStorage.StorageAlreadyExists(wkb, csRepoStorageName)) Then
        GitRepoStorage.CreateStorage wkb, csRepoStorageName, Array("RepoUrl")
    Else
        GitRepoStorage.AssignStorage wkb, csRepoStorageName
    End If

    If RepoHasAlreadyBeenAdded(GitRepoStorage, sRepoUrl) Then
        MsgBox ("This repo URL has previously been captured.  Current action ignored.")
    Else
        Set RepoUrlDictionary = New Dictionary
        RepoUrlDictionary.Add key:="RepoURL", item:=sRepoUrl
        GitRepoStorage.InsertFromDictionary RepoUrlDictionary
        MsgBox ("Repo successfully added")
    End If

End Sub



Sub RefreshAvailableFormulas()

    Dim wkb As Workbook
    Dim FormulaStorage As zLIB_ListStorage
    Dim GitRepoStorage As zLIB_ListStorage
    Dim sRepoUrl As String
    Dim LambdaXmlMap As XmlMap
    Dim i As Integer

    StandardEntry
    Set wkb = ActiveWorkbook

    Set GitRepoStorage = New zLIB_ListStorage
    GitRepoStorage.AssignStorage wkb, "GitRepos"

    Set FormulaStorage = New zLIB_ListStorage
    If Not (FormulaStorage.StorageAlreadyExists(wkb, csFunctionStorageName)) Then
        FormulaStorage.CreateStorage wkb, csFunctionStorageName, Array("Name", "RefersTo", "Category", "Author", "Comment")
    Else
        FormulaStorage.AssignStorage wkb, csFunctionStorageName
    End If

    On Error Resume Next
    wkb.XmlMaps(csLambdaXmlMapName).Delete
    On Error GoTo 0
    Set LambdaXmlMap = CreateLambdaXmlMap(wkb, csLambdaXmlMapName)
    AssignXmlMapToStorage FormulaStorage, LambdaXmlMap
    
    FormulaStorage.ClearData
    For i = 1 To GitRepoStorage.NumberOfRecords
        sRepoUrl = GitRepoStorage.FieldItemByIndex("RepoUrl", i)
        wkb.XmlMaps("LambdaMap").Import Url:=sRepoUrl, Overwrite:=False
    Next i

    wkb.XmlMaps(csLambdaXmlMapName).Delete
    StandardExit

End Sub



