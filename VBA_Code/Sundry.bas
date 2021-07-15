Attribute VB_Name = "Sundry"
Option Explicit
Option Private Module

Function CreateLambdaXmlListStorage(ByVal wkb As Workbook, ByVal sXmlMapName As String) As zLIB_ListStorage

    Dim sMap As String
    Dim LambdaXmlMap As XmlMap
    Dim storage As zLIB_ListStorage
    
    'Excel needs two elements in map such a below in order to work out the schema
    sMap = "<LambdaDocument> " & vbCrLf & _
            " <Record> " & vbCrLf & _
            "    <Name></Name><RefersTo></RefersTo><Comment></Comment> " & vbCrLf & _
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


    
    
    'Create ListObject and map to XML
    Set storage = New zLIB_ListStorage
    storage.CreateStorage wkb, "Lambdas", Array("Name", "RefersTo", "Comment")
    
'    Set lo = sht.ListObjects.Add(SourceType:=xlSrcRange, Source:=Range("A1:C1"), XlListObjectHasHeaders:=xlYes)
'    lo.HeaderRowRange.Cells(1) = "Name"
'    lo.HeaderRowRange.Cells(2) = "RefersTo"
'    lo.HeaderRowRange.Cells(3) = "Comment"
    
'    lo.ListColumns("Name").XPath.SetValue LambdaXmlMap, "/LambdaDocument/Record/Name"
'    lo.ListColumns("RefersTo").XPath.SetValue LambdaXmlMap, "/LambdaDocument/Record/RefersTo"
'    lo.ListColumns("Comment").XPath.SetValue LambdaXmlMap, "/LambdaDocument/Record/Comment"
'
'    lo.Range.NumberFormat = "@"
'    Set CreateLambdaXmlTable = lo

End Function


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
'    Const iMaxAllowableLambdas As Integer = 10000
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
'                    LambdaRecord.Name = rngCell.Offset(-1, 0).Value
'                    LambdaRecord.RefersTo = RemoveParametersFromLambda(rngCell.Formula)
'
'                    'Capture of comments are optional
'                    On Error Resume Next
'                    LambdaRecord.Comment = rngCell.Offset(-2, 0).Value
'                    On Error GoTo 0
'
'                    Lambdas(i) = LambdaRecord
'                    i = i + 1
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
'Sub PopulateLambdaInventoryList(ByVal loLambdaInventoryList As ListObject, _
'    ByRef Lambdas() As TypeLambdaRecord)
'
'    Dim i As Long
'
'    'Clear listoject data body range
'    On Error Resume Next
'    loLambdaInventoryList.DataBodyRange.EntireRow.Delete
'    On Error GoTo 0
'
'    For i = LBound(Lambdas) To UBound(Lambdas)
'        With loLambdaInventoryList
'            AddOneRowToListObject loLambdaInventoryList
'            .ListColumns("Name").DataBodyRange.Cells(i + 1) = Lambdas(i).Name
'            .ListColumns("RefersTo").DataBodyRange.Cells(i + 1) = Lambdas(i).RefersTo
'            .ListColumns("Comment").DataBodyRange.Cells(i + 1) = Lambdas(i).Comment
'        End With
'
'    Next i
'
'
'End Sub
'
'
'
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
'Sub WriteHumanReadableLambdaInventory(ByRef Lambdas() As TypeLambdaRecord, _
'    ByVal sFilePath As String)
''Requires reference to Microsoft Scripting Runtime
''Writes sStr to a text file
''*** THIS WILL OVERWRITE ANY CURRENT CONTENT OF THE FILE ***
'
'    Dim fso As Object
'    Dim oFile As Object
'    Dim i As Long
'
'    Set fso = CreateObject("Scripting.FileSystemObject")
'    Set oFile = fso.CreateTextFile(sFilePath)
'
'    For i = LBound(Lambdas) To UBound(Lambdas)
'        oFile.WriteLine ("/*------------------------------------------------------------------------------------------------------------------")
'        oFile.WriteLine ("      Name: " & Lambdas(i).Name)
'        oFile.WriteLine ("      Comment: " & Lambdas(i).Comment)
'        oFile.WriteLine ("------------------------------------------------------------------------------------------------------------------/*")
'        oFile.WriteLine (Lambdas(i).RefersTo)
'        oFile.WriteLine (vbCrLf)
'    Next i
'
'    oFile.Close
'    Set fso = Nothing
'    Set oFile = Nothing
'
'End Sub
'
'
'
'
'
'Function AssignReposList(ByVal wkb As Workbook) As ListObject
'
'    Dim sht As Worksheet
'    Dim lo As ListObject
'    Const csRepoSheetName As String = "__LambdaRepos"
'    Const csRepoListName As String = "__tbl_Repos"
'
'    If SheetExists(wkb, csRepoSheetName) Then
'        Set AssignReposList = wkb.Sheets(csRepoSheetName).ListObjects(csRepoListName)
'    Else
'        Set sht = wkb.Sheets.Add
'        sht.Name = csRepoSheetName
'        sht.Range("A1").Value = "RepoName"
'        sht.Range("B1").Value = "RepoUrl"
'        Set lo = sht.ListObjects.Add(SourceType:=xlSrcRange, Source:=sht.Range("$A1:$B1"), XlListObjectHasHeaders:=xlYes)
'        lo.Name = csRepoListName
'    End If
'
'    Set AssignReposList = lo
'
'   ' sht.Visible = xlSheetVeryHidden
'
'End Function
'
'
'Function GetRepoNameFromUrl(sRepoUrl) As String
''Repo name defined as last portion of URL path (=filename without extension)
'
'    Dim iPositionOfLastForwardSlash As Integer
'
'    iPositionOfLastForwardSlash = InStrRev(sRepoUrl, "/")
'    GetRepoNameFromUrl = Right(sRepoUrl, Len(sRepoUrl) - iPositionOfLastForwardSlash)
'    GetRepoNameFromUrl = Replace(GetRepoNameFromUrl, ".xml", "")
'    GetRepoNameFromUrl = Replace(GetRepoNameFromUrl, ".XML", "")
'    GetRepoNameFromUrl = Replace(GetRepoNameFromUrl, ".Xml", "")
'
'End Function
'
'
'Function RepoAlreadyExistsInWorkbook(ByVal loRepos As ListObject, ByVal sRepoUrl As String)
'
'    Dim sRepoName As String
'
'    sRepoName = GetRepoNameFromUrl(sRepoUrl)
'
'    If Not ListHasDataBodyRange(loRepos) Then
'        RepoAlreadyExistsInWorkbook = False
'    Else
'        RepoAlreadyExistsInWorkbook = WorksheetFunction.CountIfs( _
'            loRepos.ListColumns("RepoName").DataBodyRange, _
'            sRepoName) <> 0
'    End If
'
'End Function
'
'
'Sub AddLambdaRepoToList(ByVal loRepos As ListObject, ByVal sRepoUrl As String)
'
'   AddOneRowToListObject loRepos
'   loRepos.ListColumns("RepoUrl").DataBodyRange.Cells(loRepos.DataBodyRange.rows.Count) = sRepoUrl
'
'End Sub
'
'
'
'
'
'
'
'
'
'
'
