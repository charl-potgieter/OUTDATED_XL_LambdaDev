Attribute VB_Name = "ZZZ_OUTDATED_General"
'Option Explicit
'Option Private Module
'
'Sub StandardEntry()
'    Application.ScreenUpdating = False
'    Application.EnableEvents = False
'    Application.Calculation = xlCalculationManual
'    Application.DisplayAlerts = False
'End Sub
'
'
'Sub StandardExit()
'    Application.ScreenUpdating = True
'    Application.EnableEvents = True
'    Application.Calculation = xlCalculationAutomatic
'    Application.CutCopyMode = False
'    Application.DisplayAlerts = True
'End Sub
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
'Sub AddOneRowToListObject(lo As ListObject)
'
'    Dim str As String
'
'    On Error Resume Next
'    str = lo.DataBodyRange.Address
'    If Err.Number <> 0 Then
'        'Force empty row in databody range if it does not yet exist
'        lo.HeaderRowRange.Cells(1).Offset(1, 0) = " "
'        lo.HeaderRowRange.Cells(1).Offset(1, 0).ClearContents
'    Else
'        lo.Resize lo.Range.Resize(lo.Range.rows.Count + 1)
'    End If
'    On Error GoTo 0
'
'
'End Sub
'
'
'
'
'Function ListHasDataBodyRange(ByVal lo As ListObject) As Boolean
'
'    Dim test
'
'    On Error GoTo 0
'    test = lo.DataBodyRange.rows.Count
'    ListHasDataBodyRange = (Err.Number = 0)
'    On Error Resume Next
'
'End Function
