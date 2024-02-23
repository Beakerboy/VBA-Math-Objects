Attribute VB_Name = "VectorFactory"
' Function: Create Vector
' Create a vector from a simple array
' 
' Usage: CreateVector(array(1,2,3,4,5))
Function CreateVector(Vec As Variant) As Vector
    Dim NewVec As Variant
    num = UBound(Vec) - LBound(Vec) + 1
    ReDim NewVec(1 To num, 1 To 1)
    Dim i As Integer
    i = 1
    For Each value In Vec
        NewVec(i, 1) = value
        i = i + 1
    Next value
    Dim oResult As New Vector
    oResult.Vec = NewVec
    Set CreateVector = oResult
End Function
