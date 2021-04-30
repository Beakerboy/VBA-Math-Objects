Implements iTestCase

Dim M As New Matrix
Dim TestCase As iTestCase

' Constructor: Class_Initialize
' Initialize class members
Private Sub Class_Initialize()
    Set TestCase = CreateTestCase()
End Sub

Sub iTestCase_Setup()
   
End Sub
Sub iTestCase_TearDown()
   
End Sub

Sub MAndNProviderTest(Inputs, Expected, Message As String)
    Set M = ScalarMatrix(0, Inputs(0), Inputs(1))
    AssertEquals M.M, Expected(0), "Incorrect Number Of Rows"
    AssertEquals M.N, Expected(1), "Incorrect Number of Columns"
    AssertEquals M.GetValue(1, 1), 0, "Incorrect Value at position 1,1)"
End Sub

Function MAndNProviderTest_Data()
    MAndNProviderTest_Data = Array( _
        Array(Array(1, 1), Array(1, 1), ""), _
        Array(Array(2, 2), Array(2, 2), ""), _
        Array(Array(1, 2), Array(1, 2), "") _
    )
End Function

Sub GetRowTest()
    Set M = ScalarMatrix(6.2, 4, 4)
    Dim V As Vector
    Set V = M.GetRow(1)
    Dim aExpected As Variant
    aExpected = Array(6.2, 6.2, 6.2, 6.2)
    ReDim Preserve aExpected(1 To 4)
    Dim Expected As Vector
    Set Expected = CreateVector(aExpected)
    AssertTrue V.isEqual(Expected), "Vectors are not equal"
End Sub

Sub GetColumnTest()
    Set M = ScalarMatrix(6.2, 4, 4)
    Dim V As Vector
    Set V = M.GetColumn(1)
    Dim aExpected As Variant
    aExpected = Array(6.2, 6.2, 6.2, 6.2)
    ReDim Preserve aExpected(1 To 4)
    Dim Expected As Vector
    Set Expected = CreateVector(aExpected)
    AssertTrue V.isEqual(Expected), "Vectors are not equal"
End Sub

Sub GetDiagonalTest()
    Set M = ScalarMatrix(6.2, 4, 4)
    Dim V As Vector
    Set V = M.GetDiagonalElements()
    Dim aExpected As Variant
    aExpected = Array(6.2, 6.2, 6.2, 6.2)
    ReDim Preserve aExpected(1 To 4)
    Dim Expected As Vector
    Set Expected = CreateVector(aExpected)
    AssertTrue V.isEqual(Expected), "Vectors are not equal"
End Sub

Sub SubmatrixTest()
    Set M = Identity(3)
    Dim Expected As Matrix
    Set Expected = Identity(2)
    AssertMatrixEqual M.Submatrix(1, 2, 1, 2), Expected, "2x2 submatrix of identity is an identity"
End Sub

Sub SubmatrixOutOfBoundsProviderTest(Inputs, Expected, Message As String)
    ExpectError
    Set M = Identity(3).Submatrix(Inputs(0), Inputs(1), Inputs(2), Inputs(3))
End Sub

' Provide cases where one number is less than 1
' The second number of each pair is less than the first
' Numbers are greater than 3
Function SubmatrixOutOfBoundsProviderTest_Data()
    SubmatrixOutOfBoundsProviderTest_Data = Array( _
        Array(Array(0, 1, 1, 1), Array(), ""), _
        Array(Array(1, 0, 1, 1), Array(), ""), _
        Array(Array(1, 1, 0, 1), Array(), ""), _
        Array(Array(1, 1, 1, 0), Array(), ""), _
        Array(Array(2, 1, 2, 2), Array(), ""), _
        Array(Array(2, 2, 2, 1), Array(), ""), _
        Array(Array(3, 4, 1, 1), Array(), ""), _
        Array(Array(1, 1, 3, 4), Array(), "") _
    )
End Function

' Sub: iTestCase_RunTest
' Run a specific test.
Public Sub iTestCase_RunTest(Test As String, Optional clsObj = Nothing)
    TestCase.RunTest Test, Me
End Sub

' Sub: iTestCase_RunAllTests
' Run all tests.
Public Sub iTestCase_RunAllTests(Optional ByVal clsObj = Nothing)
    TestCase.RunAllTests Me
End Sub
