VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTestLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Purpose: This class is a helper class to output the progress of the tests to the VBA Immediate window
'This is an example of how to use for a class called oClassToTest which has an implementation of INameProvider:

'   ' Create a new instance of class to test
'    Set oClassToTest = New ClassToTest
'
'    ' Instantiate clsTestLogger object
'    Set objLogger = New clsTestLogger
'    objLogger.SetClass oClassToTest
'
'    (*** test code written here *** )
'
'   Set objLogger = Nothing


Option Explicit

Implements INameProvider

Private moClass As INameProvider

Private Sub Class_Terminate()
    Call EndTest
    Set moClass = Nothing
End Sub

Private Function INameProvider_GetClassName() As String
    INameProvider_GetClassName = "clsTestLogger"
End Function

Public Sub SetClass(ByVal Obj As INameProvider)
    Set moClass = Obj
    Call StartTest
End Sub

Private Sub StartTest()
    Call LogMessage("Starting testing")
End Sub

Private Sub EndTest()
    Call LogMessage("End testing")
End Sub

Private Sub LogMessage(ByVal strMessage As String)
    If Not moClass Is Nothing Then
        Debug.Print fncGetDateTimeStamp & vbTab & moClass.GetClassName & " : " & strMessage
    End If
End Sub

Private Function fncGetDateTimeStamp() As String
    fncGetDateTimeStamp = Format(Now, "yyyy-MM-dd hh:mm:ss")
End Function

