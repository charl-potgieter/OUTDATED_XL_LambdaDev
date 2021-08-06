Attribute VB_Name = "ZZZ_Test"
Option Explicit

Function testFunc(a, b) As Double

    testFunc = a + b

End Function






Function splitLine2(line As String) As String()

    Dim regex As Object
    Set regex = CreateObject("vbscript.regexp")
    regex.IgnoreCase = True
    regex.Global = True

    'This pattern matches only commas outside quotes
    'Pattern = ",(?=([^"]*"[^"]*")*(?![^"]*"))"
    regex.Pattern = ",(?=([^" & Chr(34) & "]*" & Chr(34) & "[^" & Chr(34) & "]*" & Chr(34) & ")*(?![^" & Chr(34) & "]*" & Chr(34) & "))"

    'regex.replaces will replace the commas outside quotes with semicolons and then the
    'Split function will split the result based on the semicollons
    splitLine2 = Split(regex.Replace(line, ";"), ";")

End Function



Sub testsplitline()

    Dim a
    
    a = splitLine2("sad,ffsdg, ""das,df,sf"" ")
    

End Sub
