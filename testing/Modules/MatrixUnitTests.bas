Public Function RunTests()
    
    Dim TestConfig As iTestableProject
    Dim MatrixTestConfig As New MatrixTestConfig
    Set TestConfig = MatrixTestConfig
    
    TestConfig.Run
End Function

Public Function AssertMatrixEqual(MyTest, Expected, Optional Message = "")

End Function
