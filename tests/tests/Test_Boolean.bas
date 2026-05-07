Attribute VB_Name = "Test_Boolean"
Sub Test_True()
    Debug.Assert True
End Sub
Sub Test_False()
    Debug.Assert False
End Sub
Sub Test_And()
    TestValue = True And True
    Debug.Assert TestValue
End Sub
Sub Test_eval()
    Debug.Assert Foo() = 1
End Sub
