Attribute VB_Name = "XmlWorkings"
Option Explicit
Option Private Module



Function CreateLambdaXmlTable(ByVal sht As Worksheet, ByVal sXmlMapName As String) As ListObject

    Dim wkb As Workbook
    Dim sMap As String
    Dim LambdaXmlMap As XmlMap
    Dim lo As ListObject
    Dim strXPath As String
    
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
    Set wkb = sht.Parent
    On Error Resume Next
    wkb.XmlMaps("LambdaMap").Delete
    On Error GoTo 0
    Set LambdaXmlMap = wkb.XmlMaps.Add(sMap, "LambdaDocument")
    LambdaXmlMap.Name = sXmlMapName

    'Create ListObject and map to XML
    Set lo = sht.ListObjects.Add(SourceType:=xlSrcRange, Source:=Range("A1:C1"), XlListObjectHasHeaders:=xlYes)
    lo.HeaderRowRange.Cells(1) = "Name"
    lo.HeaderRowRange.Cells(2) = "RefersTo"
    lo.HeaderRowRange.Cells(3) = "Comment"
    lo.ListColumns("Name").XPath.SetValue LambdaXmlMap, "/LambdaDocument/Record/Name"
    lo.ListColumns("RefersTo").XPath.SetValue LambdaXmlMap, "/LambdaDocument/Record/RefersTo"
    lo.ListColumns("Comment").XPath.SetValue LambdaXmlMap, "/LambdaDocument/Record/Comment"

    lo.Range.NumberFormat = "@"
    Set CreateLambdaXmlTable = lo

End Function


Sub WriteXmlFile(ByVal wkb As Workbook, ByVal sXmlMapName As String, ByVal sFileName)

        wkb.XmlMaps(sXmlMapName).Export Url:=sFileName, OverWrite:=True

End Sub




