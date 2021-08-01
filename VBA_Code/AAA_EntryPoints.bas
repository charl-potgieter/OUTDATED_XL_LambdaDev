Attribute VB_Name = "AAA_EntryPoints"
Option Explicit

Public Type TypeLambdaRecord
    RepoName As String
    LambdaName As String
    RefersTo As String
    Category As String
    Comment As String
End Type




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
        Array("RepoName", "Name", "RefersTo", "Category", "Author", "Comment")
    CategoryStorage.CreateStorage wkb, csCategoryStorageName, Array("Categories")
    
    'Delete sheets other than above storage
    For Each sht In wkb.Worksheets
        If sht.Name <> csCategoryStorageName And sht.Name <> csLambdaStorageName Then
            sht.Delete
        End If
    Next sht
    
    'Add validation to categories field on LambaStorage
    LambdaStorage.AddBlankRow
    wkb.Names.Add Name:="Val_Categories", RefersToR1C1:="=tbl_CategoryStorage[Categories]"
    LambdaStorage.ListObj.ListColumns("Category").DataBodyRange.Validation.Add _
        Type:=xlValidateList, Formula1:="=Val_Categories"

    FormatLambdaStorageSheet LambdaStorage
    FormatCategoryStorageSheet CategoryStorage
    Set LambdaXmlMap = CreateLambdaXmlMap(wkb, csLambdaXmlMapName)
    AssignXmlMapToStorage LambdaStorage, LambdaXmlMap

    wkb.Activate
    wkb.Sheets(1).Select
    ActiveWindow.WindowState = xlMaximized
    StandardExit
    
End Sub




'Sub ExportLambaFunctionsFromActiveWorkbookToXml()
'
'    Dim shtLambdaInventory As Worksheet
'    Dim LambdaStorage As zLIB_ListStorage
'    Dim Lambdas() As TypeLambdaRecord
'    Dim sXmlFileExportPath As String
'    Dim sHumanReadableInventoryFilePath As String
'    Dim wkb As Workbook
'    Const cXmlMapName As String = "LambdaMap"
'
'    StandardEntry
'    Set wkb = ActiveWorkbook
'
'    Set LambdaStorage = CreateLambdaXmlListStorage(wkb, cXmlMapName)
'
'    ReadLambdaFormulasInWorkbook wkb, Lambdas
'    PopulateLambdaInventoryStorage LambdaStorage, Lambdas()
'
'    sXmlFileExportPath = wkb.Path & Application.PathSeparator & "LambdaFunctions.xml"
'    WriteXmlFile wkb, cXmlMapName, sXmlFileExportPath
'
'    sHumanReadableInventoryFilePath = wkb.Path & Application.PathSeparator & "LambdaFunctions.txt"
'    WriteHumanReadableLambdaInventory Lambdas, sHumanReadableInventoryFilePath
'
'
'    LambdaStorage.Delete
'    StandardExit
'
'End Sub
'
'
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
