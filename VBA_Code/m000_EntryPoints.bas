Attribute VB_Name = "m000_EntryPoints"
Option Explicit


Public Sub CreateLambdaXmlGeneratorWorkbook()

    Dim LambdaStorage As zLIB_ListStorage
    Dim CategoryStorage As zLIB_ListStorage
    Dim wkb As Workbook
    Dim sht As Worksheet
    Dim LambdaXmlMap As XmlMap
    Const csCategoryStorageName As String = "CategoryStorage"
    Const csLambdaStorageName As String = "LambdaStorage"
    Const csLambdaXmlMapName As String = "LambdaMap"
    
    
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
    wkb.XmlMaps(csXmlMapName).Export Url:=sXmlFileExportPath, OverWrite:=True

    Set LambdaStorage = New zLIB_ListStorage
    LambdaStorage.AssignStorage wkb, "LambdaStorage"
    sHumanReadableInventoryFilePath = sExportPath & Application.PathSeparator & "LambdaFunctions.txt"
    WriteHumanReadableLambdaInventory LambdaStorage, sHumanReadableInventoryFilePath

    StandardExit

End Sub


'
'Sub AddGitRepo()
'
'    Dim sRepoUrl As String
'    Dim sRepoName As String
'    Dim GitRepoStorage As zLIB_ListStorage
'    Dim a As Boolean
'
'    sRepoUrl = InputBox("Enter Repo URL")
'
'    Set GitRepoStorage = New zLIB_ListStorage
'
'    If Not (GitRepoStorage.StorageAlreadyExists(ThisWorkbook, csRepoStorageName)) Then
'        GitRepoStorage.CreateStorage ThisWorkbook, csRepoStorageName, Array("RepoUrl")
'    Else
'        GitRepoStorage.AssignStorage ThisWorkbook, csRepoStorageName
'    End If
'
'
'    If RepoHasAlreadyBeenAdded(GitRepoStorage, sRepoUrl) Then
'        MsgBox ("This repo name has already been added")
'    Else
'        'AddLambdaRepoToWorkbook wkb, sRepoUrl
'        MsgBox ("Repo successfully added")
'    End If
'
'
'End Sub
