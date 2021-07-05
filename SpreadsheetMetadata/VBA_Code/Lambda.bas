Attribute VB_Name = "Lambda"
Option Explicit
Option Private Module


Sub ReadLambdaFormulasInWorkbook(ByVal wkb As Workbook, ByRef Lambdas() As TypeLambdaRecord)

    Dim LambdaRecord As TypeLambdaRecord
    Dim rngWithFormulas As Range
    Dim rngCell As Range
    Dim sht As Worksheet
    Dim i As Long
    Const iMaxAllowableLambdas As Integer = 10000
    
    i = 0
    ReDim Lambdas(0 To iMaxAllowableLambdas - 1)
    
    For Each sht In wkb.Sheets
        
        If SheetContainsFormulas(sht) Then
    
            Set rngWithFormulas = sht.Cells.SpecialCells(xlCellTypeFormulas)
            For Each rngCell In rngWithFormulas
                If CellContainsLambda(rngCell) Then
                    LambdaRecord.Name = rngCell.Offset(-1, 0).Value
                    LambdaRecord.RefersTo = RemoveParametersFromLambda(rngCell.Formula)
                    
                    'Capture of comments are optional
                    On Error Resume Next
                    LambdaRecord.Comment = rngCell.Offset(-2, 0).Value
                    On Error GoTo 0
                    
                    Lambdas(i) = LambdaRecord
                    i = i + 1
                End If
            Next rngCell
        End If
        
    Next sht
    
    If i <> 0 Then
        ReDim Preserve Lambdas(0 To i - 1)
    End If

End Sub


Function CellContainsLambda(ByVal rngCell As Range) As Boolean

    If Left(rngCell.Formula, 8) <> "=LAMBDA(" Then
        CellContainsLambda = False
        Exit Function
    End If
    
    'Dont allow errors unless they are #Calc errors where lambda has no parameters
    If IsError(rngCell) Then
         If rngCell.Value <> CVErr(xlErrCalc) Then
            CellContainsLambda = False
            Exit Function
        End If
    End If
            
    'If formula is in first row it does not have name above and cannot be captured
    If rngCell.Row = 1 Then
        CellContainsLambda = False
        Exit Function
    End If
           
    'Check that cell above contains the Lambdas name
    If rngCell.Offset(-1, 0).Value = "" Then
        CellContainsLambda = False
        Exit Function
    End If

    CellContainsLambda = True

End Function


Sub PopulateLambdaInventoryList(ByVal loLambdaInventoryList As ListObject, _
    ByRef Lambdas() As TypeLambdaRecord)

    Dim i As Long

    'Clear listoject data body range
    On Error Resume Next
    loLambdaInventoryList.DataBodyRange.EntireRow.Delete
    On Error GoTo 0
    
    For i = LBound(Lambdas) To UBound(Lambdas)
        With loLambdaInventoryList
            AddOneRowToListObject loLambdaInventoryList
            .ListColumns("Name").DataBodyRange.Cells(i + 1) = Lambdas(i).Name
            .ListColumns("RefersTo").DataBodyRange.Cells(i + 1) = Lambdas(i).RefersTo
            .ListColumns("Comment").DataBodyRange.Cells(i + 1) = Lambdas(i).Comment
        End With
            
    Next i


End Sub



Function RemoveParametersFromLambda(ByVal sFormula As String) As String
'sFormula contains a lambda formula.  If this includes parameters the parameters are removed

    Dim iCharacterCounter As Long
    Dim iOpenBracketCount As Integer
    
    iCharacterCounter = Len("=Lambda(") + 1
    
    'Set count as 1 as first bracket is included in the prefix and checking
    'starts after this prefix
    iOpenBracketCount = 1
    
    Do While iOpenBracketCount <> 0 And iCharacterCounter <= Len(sFormula)
        If Mid(sFormula, iCharacterCounter, 1) = "(" Then
            iOpenBracketCount = iOpenBracketCount + 1
        ElseIf Mid(sFormula, iCharacterCounter, 1) = ")" Then
            iOpenBracketCount = iOpenBracketCount - 1
        End If
        iCharacterCounter = iCharacterCounter + 1
    Loop
            
    RemoveParametersFromLambda = Left(sFormula, iCharacterCounter - 1)
            
End Function




Sub WriteHumanReadableLambdaInventory(ByRef Lambdas() As TypeLambdaRecord, _
    ByVal sFilePath As String)
'Requires reference to Microsoft Scripting Runtime
'Writes sStr to a text file
'*** THIS WILL OVERWRITE ANY CURRENT CONTENT OF THE FILE ***

    Dim fso As Object
    Dim oFile As Object
    Dim i As Long
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set oFile = fso.CreateTextFile(sFilePath)
    
    For i = LBound(Lambdas) To UBound(Lambdas)
        oFile.WriteLine ("/*------------------------------------------------------------------------------------------------------------------")
        oFile.WriteLine ("      Name: " & Lambdas(i).Name)
        oFile.WriteLine ("      Comment: " & Lambdas(i).Comment)
        oFile.WriteLine ("------------------------------------------------------------------------------------------------------------------/*")
        oFile.WriteLine (Lambdas(i).RefersTo)
        oFile.WriteLine (vbCrLf)
    Next i
    
    oFile.Close
    Set fso = Nothing
    Set oFile = Nothing

End Sub


Function RepoSheetExists(ByVal wkb As Workbook) As Boolean
    
    Dim sTestName As String
    
    On Error Resume Next
    sTestName = wkb.Sheets(csRepoSheetName).Name
    RepoSheetExists = (Err.Number = 0)
    On Error GoTo 0

End Function


Sub CreateVeryHiddenRepoSheet(ByVal wkb As Workbook)

    Dim sht As Worksheet
    Dim lo As ListObject
    
    Set sht = wkb.Sheets.Add
    sht.Name = csRepoSheetName
    sht.Range("A1").Value = "RepoUrl"
    Set lo = sht.ListObjects.Add(SourceType:=xlSrcRange, Source:=sht.Range("$A1:$A1"), XlListObjectHasHeaders:=xlYes)
    lo.Name = csRepoListObjName
    
    
    
   ' sht.Visible = xlSheetVeryHidden

End Sub


Function RepoAlreadyExistsInWorkbook(ByVal wkb As Workbook, ByVal sRepoUrl As String)
        
    'Ignore error in event listobject is empty and data body range does not exist
    On Error Resume Next
    RepoAlreadyExistsInWorkbook = WorksheetFunction.CountIfs(wkb.Sheets(csRepoSheetName).ListObjects(csRepoListObjName).ListColumns("RepoUrl").DataBodyRange, sRepoUrl) <> 0
    If Err.Number <> 0 Then RepoAlreadyExistsInWorkbook = False
    On Error GoTo 0
    
End Function


Sub AddLambdaRepoToWorkbook(ByVal wkb As Workbook, ByVal sRepoUrl As String)
    
   Dim lo As ListObject
   Dim i As Long
   
   Set lo = wkb.Sheets(csRepoSheetName).ListObjects(csRepoListObjName)
   AddOneRowToListObject lo
   lo.ListColumns("RepoUrl").DataBodyRange.Cells(lo.DataBodyRange.Rows.Count) = sRepoUrl
   
End Sub

















