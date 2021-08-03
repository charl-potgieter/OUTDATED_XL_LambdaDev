Attribute VB_Name = "Sundry"
Option Explicit
Option Private Module


Sub FormatCategoryStorageSheet(ByVal CategoryStorage As zLIB_ListStorage)
    
    CategoryStorage.AddBlankRow
    CategoryStorage.ListObj.ListColumns("Categories").Range.ColumnWidth = 50
    
End Sub


Sub FormatLambdaStorageSheet(ByVal LambdaStorage As zLIB_ListStorage)

    Dim wkb As Workbook

    LambdaStorage.AddBlankRow
    
    With LambdaStorage.ListObj
        .ListColumns("Name").Range.ColumnWidth = 25
        .ListColumns("RefersTo").Range.ColumnWidth = 90
        .ListColumns("Category").Range.ColumnWidth = 25
        .ListColumns("Author").Range.ColumnWidth = 25
        .ListColumns("Comment").Range.ColumnWidth = 40
        .DataBodyRange.RowHeight = 40
        .DataBodyRange.HorizontalAlignment = xlLeft
        .DataBodyRange.VerticalAlignment = xlTop
        .DataBodyRange.WrapText = True
    End With
    
    
    'Add Comment re Category data validation
    With Range("tbl_LambdaStorage[[#Headers],[Category]]")
        .AddComment
        .Comment.Visible = True
        .Comment.Text Text:= _
            "Drop down data validation is based on categories as captured " & _
            "in the second tab of this workbook."
        .Comment.Shape.Left = 500
        .Comment.Shape.Top = 20
        .Comment.Shape.Width = 200
        .Comment.Shape.Height = 50
    End With
    
    
    'Add validation to categories field on LambaStorage
    Set wkb = LambdaStorage.ListObj.Parent.Parent
    
    wkb.Names.Add Name:="Val_Categories", RefersToR1C1:="=tbl_CategoryStorage[Categories]"
    LambdaStorage.ListObj.ListColumns("Category").DataBodyRange.Validation.Add _
        Type:=xlValidateList, Formula1:="=Val_Categories", AlertStyle:=xlValidAlertStop
    
End Sub


Function CreateLambdaXmlMap(ByVal wkb As Workbook, ByVal sXmlMapName As String) As XmlMap

    Dim sMap As String
    Dim LambdaXmlMap As XmlMap
    Dim storage As zLIB_ListStorage

    'Excel needs two elements in map such a below in order to work out the schema
    sMap = "<LambdaDocument> " & vbCrLf & _
            " <Record> " & vbCrLf & _
            "    <Name></Name><RefersTo></RefersTo><Category></Category><Author></Author><Comment></Comment> " & vbCrLf & _
            " </Record> " & vbCrLf & _
            " <Record> " & vbCrLf & _
            "    <Name></Name><RefersTo></RefersTo><Comment></Comment> " & vbCrLf & _
            " </Record> " & vbCrLf & _
            "</LambdaDocument>"

    'Create XML map in sht parent
    On Error Resume Next
    wkb.XmlMaps("LambdaMap").Delete
    On Error GoTo 0
    Set LambdaXmlMap = wkb.XmlMaps.Add(sMap, "LambdaDocument")
    LambdaXmlMap.Name = sXmlMapName

    Set CreateLambdaXmlMap = LambdaXmlMap

End Function


Sub AssignXmlMapToStorage(ByVal LambdaStorage As zLIB_ListStorage, ByVal LambdaXmlMap As XmlMap)

    With LambdaStorage.ListObj
        .ListColumns("Name").XPath.SetValue LambdaXmlMap, "/LambdaDocument/Record/Name"
        .ListColumns("RefersTo").XPath.SetValue LambdaXmlMap, "/LambdaDocument/Record/RefersTo"
        .ListColumns("Category").XPath.SetValue LambdaXmlMap, "/LambdaDocument/Record/Category"
        .ListColumns("Author").XPath.SetValue LambdaXmlMap, "/LambdaDocument/Record/Author"
        .ListColumns("Comment").XPath.SetValue LambdaXmlMap, "/LambdaDocument/Record/Comment"

    End With

End Sub



Function WorkbookIsValidForLambdaXmlExport(ByVal wkb As Workbook) As Boolean

    WorkbookIsValidForLambdaXmlExport = True
    
    On Error Resume Next
    If Err.Number <> 0 Then
        MsgBox ("This workbook is not in the correct format to export lambda functions")
        WorkbookIsValidForLambdaXmlExport = False
    End If
    On Error GoTo 0
    
    If wkb.Path = "" Then
        MsgBox ("Workbook needs to be saved before generation of output")
        WorkbookIsValidForLambdaXmlExport = False
    End If
    

End Function


Sub WriteHumanReadableLambdaInventory(ByRef LambdasStorage As zLIB_ListStorage, _
    ByVal sFilePath As String)
'Requires reference to Microsoft Scripting Runtime
'Writes sStr to a text file
'*** THIS WILL OVERWRITE ANY CURRENT CONTENT OF THE FILE ***

    Dim fso As Object
    Dim oFile As Object
    Dim i As Long

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set oFile = fso.CreateTextFile(sFilePath)

    For i = 1 To LambdasStorage.NumberOfRecords
        oFile.WriteLine ("/*------------------------------------------------------------------------------------------------------------------")
        oFile.WriteLine ("      Formula Name:   " & LambdasStorage.FieldItemByIndex("Name", i))
        oFile.WriteLine ("      Category:       " & LambdasStorage.FieldItemByIndex("Category", i))
        oFile.WriteLine ("      Autohor:        " & LambdasStorage.FieldItemByIndex("Author", i))
        oFile.WriteLine ("      Comment:        " & LambdasStorage.FieldItemByIndex("Comment", i))
        oFile.WriteLine ("------------------------------------------------------------------------------------------------------------------/*")
        oFile.WriteLine (LambdasStorage.FieldItemByIndex("RefersTo", i))
        oFile.WriteLine (vbCrLf)
    Next i

    oFile.Close
    Set fso = Nothing
    Set oFile = Nothing

End Sub



Function RepoHasAlreadyBeenAdded(ByVal GitStorage As zLIB_ListStorage, ByVal sRepoUrl As String) As Boolean

    Dim ArrayOfReposAlreadyAdded
    Dim i As Integer

    RepoHasAlreadyBeenAdded = False

    If Not GitStorage.IsEmpty Then
        ArrayOfReposAlreadyAdded = GitStorage.ItemsInField("RepoUrl")
        i = LBound(ArrayOfReposAlreadyAdded)
        Do While i <= UBound(ArrayOfReposAlreadyAdded) And Not RepoHasAlreadyBeenAdded
            RepoHasAlreadyBeenAdded = (UCase(sRepoUrl) = UCase(ArrayOfReposAlreadyAdded(i)))
            i = i + 1
        Loop
    End If

End Function



'Function CreateLambdaXmlListStorage(ByVal wkb As Workbook, ByVal sXmlMapName As String) As zLIB_ListStorage
'
'    Dim sMap As String
'    Dim LambdaXmlMap As XmlMap
'    Dim storage As zLIB_ListStorage
'
'    'Excel needs two elements in map such a below in order to work out the schema
'    sMap = "<LambdaDocument> " & vbCrLf & _
'            " <Record> " & vbCrLf & _
'            "    <RepoName></RepoName><LambdaName></LambdaName><RefersTo></RefersTo><Comment></Comment> " & vbCrLf & _
'            " </Record> " & vbCrLf & _
'            " <Record> " & vbCrLf & _
'            "    <RepoName></RepoName><LambdaName></LambdaName><RefersTo></RefersTo><Comment></Comment> " & vbCrLf & _
'            " </Record> " & vbCrLf & _
'            "</LambdaDocument>"
'
'    'Create XML map in sht parent
'    On Error Resume Next
'    wkb.XmlMaps("LambdaMap").Delete
'    On Error GoTo 0
'    Set LambdaXmlMap = wkb.XmlMaps.Add(sMap, "LambdaDocument")
'    LambdaXmlMap.Name = sXmlMapName
'
'
'    'Create ListObject and map to XML
'    Set storage = New zLIB_ListStorage
'    storage.CreateStorage wkb, "Lambdas", Array("RepoName", "LambdaName", "RefersTo", "Comment")
'
'    With storage.ListObj
'        .ListColumns("RepoName").XPath.SetValue LambdaXmlMap, "/LambdaDocument/Record/RepoName"
'        .ListColumns("LambdaName").XPath.SetValue LambdaXmlMap, "/LambdaDocument/Record/LambdaName"
'        .ListColumns("RefersTo").XPath.SetValue LambdaXmlMap, "/LambdaDocument/Record/RefersTo"
'        .ListColumns("Comment").XPath.SetValue LambdaXmlMap, "/LambdaDocument/Record/Comment"
'        .Range.NumberFormat = "@"
'    End With
'
'    Set CreateLambdaXmlListStorage = storage
'
'End Function
'
'
'Sub WriteXmlFile(ByVal wkb As Workbook, ByVal sXmlMapName As String, ByVal sFileName)
'
'        wkb.XmlMaps(sXmlMapName).Export Url:=sFileName, OverWrite:=True
'
'End Sub
'
'
'Sub ReadLambdaFormulasInWorkbook(ByVal wkb As Workbook, ByRef Lambdas() As TypeLambdaRecord)
'
'    Dim LambdaRecord As TypeLambdaRecord
'    Dim rngWithFormulas As Range
'    Dim rngCell As Range
'    Dim sht As Worksheet
'    Dim i As Long
'    Dim fso As Scripting.FileSystemObject
'    Dim sRepoName As String
'    Const iMaxAllowableLambdas As Integer = 10000
'
'
'    Set fso = New Scripting.FileSystemObject
'    sRepoName = fso.GetBaseName(wkb.Name)
'
'    i = 0
'    ReDim Lambdas(0 To iMaxAllowableLambdas - 1)
'
'    For Each sht In wkb.Sheets
'
'        If SheetContainsFormulas(sht) Then
'
'            Set rngWithFormulas = sht.Cells.SpecialCells(xlCellTypeFormulas)
'            For Each rngCell In rngWithFormulas
'                If CellContainsLambda(rngCell) Then
'
'                    LambdaRecord.RepoName = sRepoName
'                    LambdaRecord.LambdaName = rngCell.Offset(-1, 0).Value
'                    LambdaRecord.RefersTo = RemoveParametersFromLambda(rngCell.Formula)
'
'                    'Capture of comments are optional
'                    On Error Resume Next
'                    LambdaRecord.Comment = rngCell.Offset(-2, 0).Value
'                    On Error GoTo 0
'
'                    Lambdas(i) = LambdaRecord
'                    i = i + 1
'
'                End If
'            Next rngCell
'        End If
'
'    Next sht
'
'    If i <> 0 Then
'        ReDim Preserve Lambdas(0 To i - 1)
'    End If
'
'End Sub
'
'
'
'Function SheetContainsFormulas(ByVal sht As Worksheet) As Boolean
'
'    Dim i As Double
'
'    On Error Resume Next
'    i = sht.Cells.SpecialCells(xlCellTypeFormulas).Count
'    SheetContainsFormulas = Err.Number = 0
'    On Error GoTo 0
'
'End Function
'
'
'
'
'Function CellContainsLambda(ByVal rngCell As Range) As Boolean
'
'    If Left(rngCell.Formula, 8) <> "=LAMBDA(" Then
'        CellContainsLambda = False
'        Exit Function
'    End If
'
'    'Dont allow errors unless they are #Calc errors where lambda has no parameters
'    If IsError(rngCell) Then
'         If rngCell.Value <> CVErr(xlErrCalc) Then
'            CellContainsLambda = False
'            Exit Function
'        End If
'    End If
'
'    'If formula is in first row it does not have name above and cannot be captured
'    If rngCell.Row = 1 Then
'        CellContainsLambda = False
'        Exit Function
'    End If
'
'    'Check that cell above contains the Lambdas name
'    If rngCell.Offset(-1, 0).Value = "" Then
'        CellContainsLambda = False
'        Exit Function
'    End If
'
'    CellContainsLambda = True
'
'End Function
'
'
'Sub PopulateLambdaInventoryStorage(ByVal LambdaStorage As zLIB_ListStorage, _
'    ByRef Lambdas() As TypeLambdaRecord)
'
'    Dim i As Long
'    Dim dict As Dictionary
'
'    For i = LBound(Lambdas) To UBound(Lambdas)
'        Set dict = New Dictionary
'        dict.Add "RepoName", Lambdas(i).RepoName
'        dict.Add "LambdaName", Lambdas(i).LambdaName
'        dict.Add "RefersTo", Lambdas(i).RefersTo
'        dict.Add "Comment", Lambdas(i).Comment
'        LambdaStorage.InsertFromDictionary dict
'        Set dict = Nothing
'    Next i
'
'
'End Sub
'
'
''
'Function RemoveParametersFromLambda(ByVal sFormula As String) As String
''sFormula contains a lambda formula.  If this includes parameters the parameters are removed
'
'    Dim iCharacterCounter As Long
'    Dim iOpenBracketCount As Integer
'
'    iCharacterCounter = Len("=Lambda(") + 1
'
'    'Set count as 1 as first bracket is included in the prefix and checking
'    'starts after this prefix
'    iOpenBracketCount = 1
'
'    Do While iOpenBracketCount <> 0 And iCharacterCounter <= Len(sFormula)
'        If Mid(sFormula, iCharacterCounter, 1) = "(" Then
'            iOpenBracketCount = iOpenBracketCount + 1
'        ElseIf Mid(sFormula, iCharacterCounter, 1) = ")" Then
'            iOpenBracketCount = iOpenBracketCount - 1
'        End If
'        iCharacterCounter = iCharacterCounter + 1
'    Loop
'
'    RemoveParametersFromLambda = Left(sFormula, iCharacterCounter - 1)
'
'End Function
'
'
'
'




