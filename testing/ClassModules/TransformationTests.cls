Implements iTestCase

Dim M As Matrix
Dim M2 As Matrix
Sub iTestCase_Setup()
   
End Sub
Sub iTestCase_TearDown()
   
End Sub

Sub TransposeProviderTest(Inputs, Expected, Message As String)
    Set M = MatrixFromJaggedArray(Inputs)
    Dim M2 As Matrix
    Set M2 = MatrixFromJaggedArray(Expected)
    Dim Transposed As Matrix
    Set Transposed = M.Transpose
    Test = Transposed.isEqual(M2)
    AssertMatrixEqual Transposed, M2
End Sub

Function TransposeProviderTest_Data()
    Test1 = Array( _
        Array( _
            Array(4, 3, 1), _
            Array(3, 4, 3), _
            Array(3, 3, 4) _
        ), _
        Array( _
            Array(4, 3, 3), _
            Array(3, 4, 3), _
            Array(1, 3, 4) _
        ), _
        "Transpose a 3x3 matrix" _
    )
    Test2 = Array( _
        Array( _
            Array(4, 3, 1), _
            Array(3, 4, 3) _
        ), _
        Array( _
            Array(4, 3), _
            Array(3, 4), _
            Array(1, 3) _
        ), _
        "Transpose a 2x3 matrix" _
    )
    TransposeProviderTest_Data = Array(Test1, Test2)
End Function

Function iTestCase_GetAllTests()
    iTestCase_GetAllTests = Array( _
        "TransposeProviderTest" _
    )
End Function

Function iTestCase_GetObject()
    Set iTestCase_GetObject = New TransformationTests
End Function

