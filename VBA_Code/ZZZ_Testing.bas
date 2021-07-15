Attribute VB_Name = "ZZZ_Testing"
Option Explicit
Option Private Module



Sub TestArrays()

    Dim Headers As Variant
    Dim item As Variant
    
    Headers = Array("a", "b", "c")
    
    For Each item In Headers
        Debug.Print item
    Next item
    


End Sub
