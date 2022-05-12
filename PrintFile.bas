Attribute VB_Name = "Module2"
' If/Else allows functionality on 32 and 64 bit
Option Explicit

#If VBA7 Then
    Declare PtrSafe Function apiShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
        ByVal hwnd As Long, _
        ByVal lpOperation As String, _
        ByVal lpFile As String, _
        ByVal lpParameters As String, _
        ByVal lpDirectory As String, _
        ByVal nShowCmd As Long) _
        As Long
#Else
    Declare Function apiShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
        ByVal hwnd As Long, _
        ByVal lpOperation As String, _
        ByVal lpFile As String, _
        ByVal lpParameters As String, _
        ByVal lpDirectory As String, _
        ByVal nShowCmd As Long) _
        As Long
#End If
 
' Public function for printing files using Window's apiShellExecute
Public Sub PrintFile(ByVal strPathAndFilename As String)
 
    Call apiShellExecute(0, "print", strPathAndFilename, vbNullString, vbNullString, 0)
 
End Sub
 
' Example of function call
Sub Test()


    PrintFile ("Drive:\File Path\file.extension")


End Sub

