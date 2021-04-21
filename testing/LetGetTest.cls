Implements iTestCase

Dim M As New Matrix

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

Function iTestCase_GetAllTests()
    iTestCase_GetAllTests = Array("MAndNProviderTest")
End Function

Function iTestCase_GetObject()
    Set iTestCase_GetObject = New LetGetTest
End Function
