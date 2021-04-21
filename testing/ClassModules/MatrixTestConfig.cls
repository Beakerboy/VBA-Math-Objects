Implements iTestableProject

Dim bOutputFailures As Boolean

Dim TestRunner As TestRunner

' Constructor: Class_Initialize
' Configure the test runner
Private Sub Class_Initialize()
    ' Print each failure as it occurs
    bOutputFailures = True
End Sub

Public Property Get iTestableProject_GetOutputFailures()
    iTestableProject_GetOutputFailures = bOutputFailures
End Property

Public Function iTestableProject_GetTestCases()
        iTestableProject_GetTestCases = Array(New LetGetTest, New PropertyTests, New ArithmeticTests)
End Function

' Function: Run
' Perform all the actions.
Public Function iTestableProject_Run()
    Configure
    Run
    Report
End Function

Private Sub Configure()
    Set TestRunner = CreateTestRunner()
    
    ' Pass this object to the runner so it has access to the necessary data
    Set TestRunner.TestConfig = New MatrixTestConfig
End Sub

Private Sub Run()
    ' We want to run all Test Cases
    TestRunner.TestAllCases
End Sub
Private Sub Report()
    ' We want to see the summary report in a message box at the end
    TestReporter.ResultsMsgBox
    
    ' We also want to see a list of all the failures
    TestReporter.PrintFailureLog
End Sub
