' Class: Matrix
' Perform matrix algebra and provide matrix attributes.
'
' Version: 210331

' Variable: dA
' A 2D array of doubles
' dA(intM, intN)
Dim dA() As Variant

' Variable: intM
' The number of rows in the matrix
Dim intM As Integer

' Variable: intM
' The number of columns in the matrix
Dim intN As Integer 'columns

' Constructor: Class_Initialize
' Initializes the object.
Private Sub Class_Initialize()
End Sub

' Property: Mat
' Set the matrix data.
'
' Parameter:
'   vA - A 2D array of numerical data.
Public Property Let Mat(vA As Variant)
    dA = vA
    intM = UBound(vA, 1)
    intN = UBound(vA, 2)
End Property

' Property: N
' Get the number of rows.
Public Property Get N() As Integer
    N = intN
End Property

' Property: M
' Get the number of columns.
Public Property Get M() As Integer
    M = intM
End Property

' Property: Mat
' Get the matrix data.
Public Property Get Mat() As Variant
    Mat = dA
End Property

' Function: isDiaginal
' Is the matrix a diagonal matrix?
'
' Returns: True or False
Public Function isDiagonal()
    If intM = intN Then
        isDiagonal = True
    Else
        isDiagonal = False
    End If
End Function

' Function: Add
' Perform matrix addition. Throws an error if the matrices do not have matching sizes
'
' Parameters:
'   oInput - Addand
'
' Returns:
'   The sum of the two matrices
Public Function Add(oInput As Matrix) As Matrix
    If Me.M <> oInput.M Or Me.N <> oInput.N Then
        Err.Raise Number:=vbObjectError + 513, _
              Description:="Addition requires compatible sized matricies"
    End If
    Dim i As Integer
    Dim vSum() As Variant
    
    ReDim vSum(1 To Me.M, 1 To Me.N)
    For i = 1 To Me.M
        Dim j As Integer
        For j = 1 To Me.N
            vSum(i, j) = dA(i, j) + oInput.getValue(i, j)
        Next j
    Next i
    Dim oResult As New Matrix
    oResult.Mat = vSum
    Set Add = oResult
End Function

' Function: Subtract
' Perform matrix subtraction. Throws an error if the matrices do not have matching sizes
'
' Parameters:
'   oInput - Subtrahand
'
' Returns:
'   The difference of the two matrices
Public Function Subtract(oInput As Matrix) As Matrix
    If Me.M <> oInput.M Or Me.N <> oInput.N Then
        Err.Raise Number:=vbObjectError + 513, _
              Description:="Subtraction requires compatible sized matricies"
    End If
    Dim i As Integer
    Dim vSum() As Variant
    
    ReDim vSum(1 To intM, 1 To intN)
    For i = 1 To Me.M
        Dim j As Integer
        For j = 1 To Me.N
            vSum(i, j) = dA(i, j) - oInput.getValue(i, j)
        Next j
    Next i
    Dim oResult As New Matrix
    oResult.Mat = vSum
    Set Subtract = oResult
End Function

' Function: Multiply
' Perform matrix multiplication. Throws an error if the matrices do not have compatible sizes.
'
' Parameters:
'   oInput - The multiplier.
'
' Returns:
'   The sum of the two matrices.
Public Function Multiply(oInput As Matrix) As Matrix
    If Me.N <> oInput.M Then
        Err.Raise Number:=vbObjectError + 513, _
              Description:="Multiplication requires compatible sized matricies"
    End If
    Dim vProduct() As Variant
    
    ReDim vProduct(1 To Me.M, 1 To oInput.N)
    Dim i As Long
    For i = 1 To Me.M
        Dim oRow As Vector
        Set oRow = Me.GetRow(i)
        Dim j As Long
        For j = 1 To oInput.N
            Dim oColumn As Vector
            Set oColumn = oInput.GetColumn(j)
            vProduct(i, j) = oRow.DotProduct(oColumn)
        Next j
    Next i
    Dim oResult As New Matrix
    oResult.Mat = vProduct
    Set Multiply = oResult
End Function

' Function: ScalarMultiply
' Multiply a matrix by a scalar.
'
' Parameters:
'   oInput - Scalar
'
' Returns:
'   The product of the matrix by the scalar.
Public Function ScalarMultiply(dInput) As Matrix
    Dim i As Integer
    Dim vProduct() As Variant
    
    ReDim vProduct(1 To Me.M, 1 To Me.N)
    For i = 1 To Me.M
        Dim j As Integer
        For j = 1 To Me.N
            vProduct(i, j) = dA(i, j) * dInput
        Next j
    Next i
    Dim oResult As New Matrix
    oResult.Mat = vProduct
    Set ScalarMultiply = oResult
End Function

' Function: ScalarDivide
' Divide a matrix by a scalar.
'
' Parameters:
    '   oInput - The scalar.
'
' Returns:
'   The result of the matrix divided by the scalar.
Public Function ScalarDivide(dInput As Double) As Matrix
    Dim i As Integer
    Dim vProduct() As Variant
    
    ReDim vProduct(1 To Me.M, 1 To Me.N)
    For i = 1 To Me.M
        Dim j As Integer
        For j = 1 To Me.N
            vProduct(i, j) = dA(i, j) / dInput
        Next j
    Next i
    Dim oResult As New Matrix
    oResult.Mat = vProduct
    Set ScalarDivide = oResult
End Function

' Function: Transpose
' Transpose the matrix.
'
' Returns:
'   A new matrix that is the transposition on this matrix.
Public Function Transpose() As Matrix
    Dim vResult As Variant
    ReDim vResult(1 To intN, 1 To intM)
    For i = 1 To intM
        For j = 1 To intN
            vResult(j, i) = Me.getValue(i, j)
        Next j
    Next i
    Dim oResult As New Matrix
    oResult.Mat = vResult
    Set Transpose = oResult
End Function

' Function: GetRow
' Extract a row from the matrix.
'
' Parameters:
'   iRow - The row number.
'
' Returns:
'   A vector that contains the values of the row.
Public Function GetRow(iRow) As Vector
    Dim vResult() As Variant
    ReDim vResult(1 To intN, 1 To 1)
    For i = 1 To intN
        vResult(i, 1) = dA(iRow, i)
    Next i
    Dim oResult As New Vector
    oResult.Vec = vResult
    Set GetRow = oResult
End Function

' Function: GetColumn
' Extract a column from the matrix.
'
' Parameters:
'   iCol - The column number.
'
' Returns:
'   A vector that contains the values of the column.
Public Function GetColumn(iCol) As Vector
    Dim vResult() As Variant
    ReDim vResult(1 To intM, 1 To 1)
    For i = 1 To intM
        vResult(i, 1) = dA(i, iCol)
    Next i
    Dim oResult As New Vector
    oResult.Vec = vResult
    Set GetColumn = oResult
End Function

Public Function getValue(i, j)
    getValue = dA(i, j)
End Function

' Function: AugmentRight
' Combine two matrices, positioning the supplied matrix to the right of this matrix.
'
' Parameters:
'   inMatrix - The matrix whos elements will be on the right.
'
' Returns:
'   A new matrix that is the combination of this and inMatrix.
Public Function AugmentRight(inMatrix As Matrix)
    If Me.M <> inMatrix.M Then
        Err.Raise Number:=vbObjectError + 513, _
    Description:="Augmentation requires compatible sized matricies"
    End If
    Dim vResult As Variant
    vResult = dA
    ReDim Preserve vResult(1 To Me.M, 1 To Me.N + inMatrix.N)
    Dim oResults As New Matrix
    oResults.Mat = vResult
    Dim i As Long
    For i = 1 To inMatrix.N
        Set oResults = oResults.ReplaceColumn(Me.N + i, inMatrix.GetColumn(i))
    Next i
    Set AugmentRight = oResults
End Function

' Function: AugmentBelow
' Combine two matrices, positioning the supplied matrix below this matrix.
'
' Parameters:
'   inMatrix - The matrix whos elements will be below.
'
' Returns:
'   A new matrix that is the combination of this and inMatrix.
Public Function AugmentBelow(inMatrix As Matrix)
    Set AugmentBelow = Me.Transpose().AugmentRight(inMatrix.Transpose()).Transpose()
End Function

' Function: ReplaceColumn
' Replace a specified column with values from a vector
'
' Parameters:
'   iNum     - The column to replace.
'   inColumn - The values to place in the matrix.
'
' Returns:
'   A new matrix with replaced values in the specified column.
Public Function ReplaceColumn(iNum, inColumn As Vector) As Matrix
    If Me.M <> inColumn.M Then
        Err.Raise Number:=vbObjectError + 513, _
            Description:="Replacement requires compatible sized elements"
    End If
    vData = dA
    Dim i As Long
    For i = 1 To Me.M
        vData(i, iNum) = inColumn.getValue(i)
    Next i
    Dim oResult As New Matrix
    oResult.Mat = vData
    Set ReplaceColumn = oResult
End Function

' Function: ReplaceRow
' Replace a specified row with values from a vector
'
' Parameters:
'   iNum     - The row to replace.
'   inColumn - The values to place in the matrix.
'
' Returns:
'   A new matrix with replaced values in the specified row.
Public Function ReplaceRow(iNum, inRow As Vector) As Matrix
    If Me.N <> inRow.N Then
        Err.Raise Number:=vbObjectError + 513, _
            Description:="Replacement requires compatible sized elements"
    End If
    vData = dA
    Dim i As Long
    For i = 1 To Me.N
        vData(iNum, i) = inRow.getValue(i)
    Next i
    Dim oResult As New Matrix
    oResult.Mat = vData
    Set ReplaceRow = oResult
End Function

' Function: ExcludeRow
' Remove a specified row from this matrix.
'
' Parameters:
'   iNum - The row to remove.
'
' Returns:
'   A new matrix with the specified row removed.
Public Function ExcludeRow(iNum) As Matrix
    'Check that row exists
    Dim oResult As New Matrix
    If iNum = 1 Then
        Set oResult = Submatrix(2, intM, 1, intN)
    ElseIf iNum = intM Then
        Set oResult = Submatrix(1, intM - 1, 1, intN)
    Else
        Set oResult = Me.Submatrix(1, iNum - 1, 1, intN).AugmentBelow(Me.Submatrix(iNum + 1, intM, 1, intN))
    End If
    Set ExcludeRow = oResult
End Function

' Function: Trace
' Calculate the trace of a matrix. The matrix must be square.
'
' Returns:
'   The trace of the matrix.
Public Property Get Trace()
    If intM <> intN Then
        Err.Raise Number:=vbObjectError + 513, _
              Description:="Trace Requires a square Matrix"
    End If
    Trace = GetDiagonalElements.Sum
End Property

' Function: GetDiagonalElements
' Create a vector from the diagonal elements of this matrix.
'
' Returns:
'   A vector.
'
' ToDo:
'   The for loop should be i = 1 to min(intM, intN)
Public Function GetDiagonalElements() As Vector
    Dim vResult() As Variant
    ReDim vResult(1 To intM, 1 To 1)
    For i = 1 To intM
        vResult(i, 1) = dA(i, i)
    Next i
    Dim oResult As New Vector
    oResult.Vec = vResult
    Set GetDiagonalElements = oResult
End Function

' Function: SwapColumns
' Swap two columns in a matrix.
'
' Parameters:
'   col1 - One of the columns to swap.
'   col2 - The other column to swap.
'
' Returns:
'   A new matrix with the specified columns swapped.
Public Function SwapColumns(col1, col2) As Matrix
    Dim oResult As Matrix
    Set oResult = Me.ReplaceColumn(col1, Me.GetColumn(col2)).ReplaceColumn(col2, Me.GetColumn(col1))
    
    Set SwapColumns = oResult
End Function

' Function: SubMatrix
' Create a matrix from a subset of this matrix.
'
' Parameters:
'   startm - The 1 indexed starting row for the new matrix.
'   endm   - The 1 indexed ending row for the new matrix.
'   startn - The 1 indexed starting column for the new matrix.
'   endn   - The 1 indexed starting column for the new matrix.
'
' Returns:
'   A new matrix with the specified elements.
Public Function Submatrix(startm, endm, startn, endn)
    ' Check that ends are after starts and that no out of bounds
    Dim vResult As Variant
    ReDim vResult(1 To endm - startm + 1, 1 To endn - startn + 1)
    
    For i = startm To endm
        For j = startn To endn
            vResult(i - startm + 1, j - startn + 1) = dA(i, j)
        Next j
    Next i
    Dim oResult As New Matrix
    oResult.Mat = vResult
    Set Submatrix = oResult
End Function

' Function: Inverse
' Calculate the inverse of a matrix
'
' Returns: The matrix inverse. Currently only works for a diagonal matrix.
'
' ToDo:
'   Use the minverse worksheet function.
Public Function Inverse() As Matrix
    If intM <> intN Then
        Err.Raise Number:=vbObjectError + 513, _
            Description:="Inverse Requires a square Matrix"
    End If
    Dim vResult As Variant
    vResult = ScalarMatrix(0, intM, intM).Mat
    If isDiagonal() Then
        For i = 1 To intM
            vResult(i, i) = 1 / dA(i, i)
        Next i
    End If
    Dim oResult As New Matrix
    oResult.Mat = vResult
    Set Inverse = oResult
End Function
