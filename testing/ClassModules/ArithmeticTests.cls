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
            Array(MatrixFromJaggedArray( _
                Array( _
                    Array(4, 3, 3), _
                    Array(3, 4, 3), _
                    Array(3, 3, 4) _
                )) _
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

Function SubtrsactionProviderTest_Data()
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

Function iTestCase_GetAllTests()
    iTestCase_GetAllTests = Array("AdditionProviderTest")
End Function

Function iTestCase_GetObject()
    Set iTestCase_GetObject = New ArithmeticTests
End Function
