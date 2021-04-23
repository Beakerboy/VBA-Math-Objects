Public Function RunTests()
    
    Dim TestConfig As iTestableProject
    Dim MatrixTestConfig As New MatrixTestConfig
    Set TestConfig = MatrixTestConfig
    
    TestConfig.Run
End Function

Public Function AssertMatrixEqual(MyTest As Matrix, Expected As Matrix, Optional Message As String = "")
    Dim Rows As Variant
    Rows = MyTest.toJaggedArray
    Dim Rows1 As Variant
    Rows1 = Expected.toJaggedArray
    Dim i As Integer
    i = 0
    For Each Row In Rows1
        If i = 0 Then
            Line = "Expecting "
            Line1 = " Given "
        Else
            Line = "          "
            Line1 = "       "
        End If
        Line = Line & "[" & Join(Rows1(i), ", ") & "]" & Line1 & "[" & Join(Rows(i), ", ") & "]"
        Rows(i) = Line
        i = i + 1
    Next Row
    Message = Join(Rows, vbNewLine)
    AssertMatrixEqual = AssertTrue(MyTest.isEqual(Expected), Message)
End Function
