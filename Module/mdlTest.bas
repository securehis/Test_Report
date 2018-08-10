Attribute VB_Name = "mdlTest"
Public Function sD_mTest() As String
On Error GoTo ErrorLine:

    sD_mTest = "Module Test"

    Exit Function
    
ErrorLine:

    Debug.Print "Module Error!!"
End Function

