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

Sub ScalarIsNotDiagonalTest()
    Set M = ScalarMatrix(1, 2, 2)
    AssertFalse M.isDiagonal, "Scaler Matrix is not a Diagonal Matrix"
End Sub

Sub RectangularIsNotDiagonalTest()
    Set M = ScalarMatrix(0, 3, 2)
    AssertFalse M.isDiagonal, "Rectangular Matrix is not a Diagonal Matrix"
End Sub

Sub IdentityIsDiagonalTest()
    Set M = Identity(1)
    AssertTrue M.isDiagonal, "Identity Matrix is a Diagonal Matrix"
End Sub

Function iTestCase_GetAllTests()
    iTestCase_GetAllTests = Array("ScalarIsNotDiagonalTest", "IdentityIsDiagonalTest", "RectangularIsNotDiagonalTest")
End Function

Function iTestCase_GetObject()
    Set iTestCase_GetObject = New PropertyTests
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
