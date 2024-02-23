Attribute VB_Name = "MatrixFactory"
' Function: Scalarmatrix
' Make a matrix of dimensions m x n where every element is the value 'value'
Function ScalarMatrix(value, matrix_length, matrix_width) As Matrix
    Dim vMatrix As Variant
    ReDim vMatrix(1 To matrix_length, 1 To matrix_width)
    For i = 1 To matrix_length
        For j = 1 To matrix_width
            vMatrix(i, j) = value
        Next j
    Next i
    Dim oResult As New Matrix
    oResult.Mat = vMatrix
    Set ScalarMatrix = oResult
End Function

' Function: Identity
' Create an Identity Matrix of a given size
Function Identity(Size_num) As Matrix
    vMatrix = ScalarMatrix(0, Size_num, Size_num).Mat
    For i = 1 To Size_num
        vMatrix(i, i) = 1
    Next i
    Dim oResult As New Matrix
    oResult.Mat = vMatrix
    Set Identity = oResult
End Function

Function DiagonalMatrix(oVector As Vector) As Matrix
    Dim oResult As Matrix
    Set oResult = ScalarMatrix(0, oVector.M, oVector.M)
    Dim vMatrix As Variant
    vMatrix = oResult.Mat
    For i = 1 To oVector.M
        vMatrix(i, i) = oVector.getValue(i)
    Next i
    Set oResult = New Matrix
    oResult.Mat = vMatrix
    Set DiagonalMatrix = oResult
End Function

Function MatrixFromJaggedArray(aData As Variant) As Matrix
    Dim oResult As Matrix
    matrix_length = UBound(aData) + 1
    matrix_width = UBound(aData(0)) + 1
    ReDim vMatrix(1 To matrix_length, 1 To matrix_width)
    Dim i As Integer
    Dim j As Integer
    For i = 1 To matrix_length
        For j = 1 To matrix_width
            vMatrix(i, j) = aData(i - 1)(j - 1)
        Next j
    Next i
    Set oResult = New Matrix
    oResult.Mat = vMatrix
    Set MatrixFromJaggedArray = oResult
End Function

Function CreateMatrix(Mat As Variant) As Matrix
    Dim oResult As New Matrix
    oResult.Mat = Mat
    Set CreateMatrix = oResult
End Function
