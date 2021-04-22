Implements iTestCase

Dim M As New Matrix

Sub iTestCase_Setup()
   
End Sub
Sub iTestCase_TearDown()
   
End Sub

Sub AdditionProviderTest(Inputs, Expected, Message As String)
    Set M = Inputs(0)
    Dim M2 As Matrix
    Set M2 = Inputs(1)
    Dim M3 As Matrix
    Set M3 = Expected(0)
    AssertTrue M.Add(M2).isEqual(M3), Message
End Sub

Function AdditionProviderTest_Data()
    AdditionProviderTest_Data = Array( _
        Array( _
            Array(ScalarMatrix(3, 3, 3), Identity(3)), _
            Array( _
                MatrixFromJaggedArray( _
                    Array( _
                        Array(4, 3, 3), _
                        Array(3, 4, 3), _
                        Array(3, 3, 4) _
                    ) _
                ) _
            ), _
            "Adding two 3x3 matricies" _
        ) _
    )
End Function

Sub SubtractionProviderTest(Inputs, Expected, Message As String)
    Set M = Inputs(0)
    Dim M2 As Matrix
    Set M2 = Inputs(1)
    Dim M3 As Matrix
    Set M3 = Expected(0)
    AssertTrue M.Subtract(M2).isEqual(M3), Message
End Sub

Function SubtractionProviderTest_Data()
    SubtractionProviderTest_Data = Array( _
        Array( _
            Array(ScalarMatrix(3, 3, 3), Identity(3)), _
            Array(MatrixFromJaggedArray( _
                Array( _
                    Array(2, 3, 3), _
                    Array(3, 2, 3), _
                    Array(3, 3, 2) _
                )) _
            ), _
            "Subtraction on two 3x3 matricies" _
        ) _
    )
End Function

Sub MultiplicationProviderTest(Inputs, Expected, Message As String)
    Set M = Inputs(0)
    Dim M2 As Matrix
    Set M2 = Inputs(1)
    Dim M3 As Matrix
    Set M3 = Expected(0)
    AssertTrue M.Multiply(M2).isEqual(M3), Message
End Sub

Function MultiplicationProviderTest_Data()
    MultiplicationProviderTest_Data = Array( _
        Array( _
            Array(ScalarMatrix(6.2, 1, 1), ScalarMatrix(2, 1, 1)), _
            Array(ScalarMatrix(12.4, 1, 1)), _
            "Multiplication of two 1x1 float matrices" _
        ), _
        Array( _
            Array(ScalarMatrix(3, 3, 3), Identity(3)), _
            Array(ScalarMatrix(3, 3, 3)), _
            "Multiplication of a 3x3 by the identity" _
        ) _
    )
End Function

Sub ScalarMultiplicationProviderTest(Inputs, Expected, Message As String)
    Set M = Inputs(0)
    Dim M2 As Matrix
    Set M2 = Expected(0)
    AssertTrue M.ScalarMultiply(Inputs(1)).isEqual(M2), Message
End Sub

Function ScalarMultiplicationProviderTest_Data()
    ScalarMultiplicationProviderTest_Data = Array( _
        Array( _
            Array(ScalarMatrix(4, 3, 3), 2), _
            Array(ScalarMatrix(8, 3, 3)), _
            "Multiplication of a 3x3 matrix by an Integer" _
        ), _
        Array( _
            Array(Identity(3), 6.2), _
            Array( _
                MatrixFromJaggedArray( _
                    Array( _
                        Array(6.2, 0, 0), _
                        Array(0, 6.2, 0), _
                        Array(0, 0, 6.2) _
                    ) _
                ) _
            ), _
            "Multiplication of a 3x3 identity matrix by a float" _
        ) _
    )
End Function

Function AddExceptionTest()
    Set M = Identity(2)
    Dim M2 As Matrix
    Set M2 = ScalarMatrix(2, 2, 3)
    Dim M3 As Matrix
    ExpectError
    Set M3 = M.Add(M2)
End Function

Function iTestCase_GetAllTests()
    iTestCase_GetAllTests = Array("AdditionProviderTest", "SubtractionProviderTest", "MultiplicationProviderTest", "ScalarMultiplicationProviderTest", "AddExceptionTest")
End Function

Function iTestCase_GetObject()
    Set iTestCase_GetObject = New ArithmeticTests
End Function