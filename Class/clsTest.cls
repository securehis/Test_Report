VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function sD_cTest() As String
On Error GoTo ErrorLine:

<<<<<<< HEAD
    sD_cTest = "Class Test Alter2"
=======
    sD_cTest = "Class Test Alter"
>>>>>>> 6b60a56b547ac048401bf5220acc290b45efbfea

ErrorLine:

    Debug.Print "Class Error!!"
End Function
