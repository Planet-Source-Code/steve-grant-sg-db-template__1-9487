Attribute VB_Name = "libErrorHandling"
' Module     : libErrorHandling
' Description:
' Procedures : MsgErr
' Author     : Steve Grant
'--------------------------------------------------------------------------------
Option Explicit

Private Const mcstrMod As String = "libErrorHandling"


Sub MsgErr(strProc As String, intErrNumber As Integer, strErrDesc As String)
  ' Comments  : When an error occurs, this proc is called and it displays a msg to
  '             the user.
  ' Parameters: strProc     : The procedure that generated the error.
  '             intErrNumber: The error number.
  '             strErrDesc  : The description of the error.
  ' Returns   : Nothing.
  ' Author    : Steve Grant
  '------------------------------------------------------------------------------
  On Error GoTo PROC_ERR
  
  Dim strMsg As String
  
  strMsg = "An error occured in the folowing procedure: <" & strProc & ">." & _
           vbCrLf & vbCrLf & "Please report this error to:." & _
           vbCrLf & vbCrLf & CStr(intErrNumber) & " - " & strErrDesc
  
  MsgBox strMsg, vbOKOnly + vbCritical, "Error"
  
PROC_EXIT:
  Exit Sub
  
PROC_ERR:
  MsgBox ("An error occured in the module <lib_ErrorHandling>" & _
          vbCrLf & vbCrLf & "Please report this error to:")
  Resume PROC_EXIT
End Sub



