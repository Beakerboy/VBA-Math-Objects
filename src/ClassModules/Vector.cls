' Class: Vector
' A list of numbers 
' Version: 19091301

' Variable: dA
' A column of doubles
' dA(intM, 1)
Dim dA() As Variant

' Variable: intM
' The number of elements in the vector
Dim intM As Integer  'number of elements

' Constructor: Class_Initialize
' Initializes the object.
Private Sub Class_Initialize()
End Sub

' Property: Value
' Set the value of the vector
Public Property Let Vec(vA As Variant)
    dA = vA
    intM = UBound(vA, 1)
End Property

' Property: M
' Get the number of elements in the vector
Public Property Get M() As Integer
    M = intM
End Property
    
' Property: Vec
' Get the raw data of the vector
Public Property Get Vec() As Variant
    Vec = dA
End Property

' Function: getValue
' Get an individual value from the vecotr
'
' Parameters:
'
'    i - The position - 1 indexed
'
' Returns: The value of the vector at position i
Public Function getValue(i)
    getValue = dA(i, 1)
End Function

Public Function Magnitude()
    Sum = 0
    For i = 1 To Me.M
        Sum = Sum + dA(i, 1) ^ 2
    Next i
    Magnitude = Math.sqr(Sum)
End Function

' Operations between Vectors

Public Function DotProduct(oInput As Vector) As Double
    Dim vProduct() As Double
    newM = oInput.M
    If Me.M <> newM Then
        Err.Raise Number:=vbObjectError + 513, _
              Description:="DotProduct requires equal length Vectors"
    End If
    vInput = oInput.Vec
    Dim Sum As Double
    Sum = 0
    For i = 1 To Me.M
        Sum = Sum + dA(i, 1) * vInput(i, 1)
    Next i

    DotProduct = Sum
End Function

' Aggregation Functions

Public Function Average() As Double
    Average = WorksheetFunction.Average(dA)
End Function

Public Function Stdev() As Double
    Stdev = WorksheetFunction.Stdev(dA)
End Function

Public Function Sum() As Double
    Sum = WorksheetFunction.Sum(dA)
End Function

' Operations with scalars

Public Function ScalarDivide(inNumber) As Vector
    Dim vResult As Variant
    ReDim vResult(1 To M, 1 To 1)
    For i = 1 To M
        vResult(i, 1) = dA(i, 1) / inNumber
    Next i
    Dim oResult As New Vector
    oResult.Vec = vResult
    Set ScalarDivide = oResult
End Function

Public Function ScalarMultiply(inNumber) As Vector
    Dim vResult As Variant
    ReDim vResult(1 To M, 1 To 1)
    For i = 1 To M
        vResult(i, 1) = dA(i, 1) * inNumber
    Next i
    Dim oResult As New Vector
    oResult.Vec = vResult
    Set ScalarMultiply = oResult
End Function

Public Function pow(exp) As Vector
    Dim vResult As Variant
    ReDim vResult(1 To M, 1 To 1)
    For i = 1 To M
        vResult(i, 1) = dA(i, 1) ^ exp
    Next i
    Dim oResult As New Vector
    oResult.Vec = vResult
    Set pow = oResult
End Function

' Unary Operations

Public Function sqr() As Vector
    Dim vResult As Variant
    ReDim vResult(1 To M, 1 To 1)
    For i = 1 To M
        vResult(i, 1) = Math.sqr(dA(i, 1))
    Next i
    Dim oResult As New Vector
    oResult.Vec = vResult
    Set sqr = oResult
End Function

Public Function Inverse()
    Dim vResult As Variant
    ReDim vResult(1 To M, 1 To 1)
    For i = 1 To M
        vResult(i, 1) = 1 / dA(i, 1)
    Next i
    Dim oResult As New Vector
    oResult.Vec = vResult
    Set sqr = oResult
End Function

' Function: subVector
' Generate a Vector that is a subset of this
'
' Parameters:
'
'    iOffset - The starting value, 1 indexed
'    iLength - How many items to return
Public Function subVector(iOffset As Integer, Optional iLength = 0) As Vector
    If iLength < 1 Or iLength + iOffset - 1 > intM Then
        iLength = intM - iOffset + 1
    End If
    Dim vResult As Variant
    ReDim vResult(1 To iLength, 1 To 1)
    For i = 1 To iLength
        vResult(i, 1) = dA(i + iOffset - 1, 1)
    Next i
    Dim oResult As New Vector
    oResult.Vec = vResult
    Set subVector = oResult
End Function

' Function: toRowmatrix
' Generate a row matrix version of this vector
' Should the matrix factory do this instead?
Public Function toRowMatrix() As Matrix
    Dim oResult As New Matrix
    oResult.Mat = dA
    Set toRowMatrix = oResult.Transpose()
End Function
