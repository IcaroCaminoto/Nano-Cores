Option Explicit

Public TotalTests As Long
Public PassedTests As Long
Public FailedTests As Long

Public Sub ResetTestCounters()
    TotalTests = 0
    PassedTests = 0
    FailedTests = 0
End Sub

Public Sub AssertEqual(ByVal actual As Variant, ByVal expected As Variant, ByVal testName As String)
    TotalTests = TotalTests + 1
    If actual = expected Then
        PassedTests = PassedTests + 1
        Debug.Print " [PASS] " & testName
    Else
        FailedTests = FailedTests + 1
        Debug.Print "![FAIL] " & testName & " | Expected: [" & expected & "] but got: [" & actual & "]"
    End If
End Sub

Public Sub AssertDoubleApprox(ByVal actual As Double, ByVal expected As Double, ByVal tolerance As Double, ByVal testName As String)
    TotalTests = TotalTests + 1
    If Abs(actual - expected) <= tolerance Then
        PassedTests = PassedTests + 1
        Debug.Print " [PASS] " & testName
    Else
        FailedTests = FailedTests + 1
        Debug.Print "![FAIL] " & testName & " | Expected ~[" & expected & "] +/- " & tolerance & " but got: [" & actual & "]"
    End If
End Sub

Public Sub PrintTestSummary()
    Debug.Print "=========================================="
    Debug.Print "TEST RUN COMPLETE: " & TotalTests & " Tests Executed."
    Debug.Print " Passed: " & PassedTests & " | Failed: " & FailedTests
    Debug.Print "=========================================="
    If FailedTests > 0 Then
        Err.Raise vbObjectError + 999, "TestRunner", "One or more contract assertion tests failed!"
    End If
End Sub