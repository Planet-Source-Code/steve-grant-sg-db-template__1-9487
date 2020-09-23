Attribute VB_Name = "libHelp"
' Module     : libHelp
' Description: Library for Help.
' Procedures : DisplayAbout
'              DisplayManual
' Author     : Steve Grant
'--------------------------------------------------------------------------------
Option Explicit

Private Const mcstrMod As String = "libHelp"


Sub DisplayAbout()
  ' Comments  : Display the About dialog.
  ' Parameters: None.
  ' Returns   : Nothing.
  ' Author    : Steve Grant
  '------------------------------------------------------------------------------
  Const cstrProc As String = "DisplayAbout"
  'On Error GoTo PROC_ERR
  
  NIF
  
PROC_EXIT:
  Exit Sub
  
PROC_ERR:
  Call MsgErr(mcstrMod & "." & cstrProc, Err.Number, Err.Description)
  Resume PROC_EXIT
End Sub


Sub DisplayManual()
  ' Comments  : Display the manual.
  ' Parameters: None.
  ' Returns   : Nothing.
  ' Author    : Steve Grant
  '------------------------------------------------------------------------------
  Const cstrProc As String = "DisplayManual"
  'On Error GoTo PROC_ERR
  
  NIF
  
PROC_EXIT:
  Exit Sub
  
PROC_ERR:
  Call MsgErr(mcstrMod & "." & cstrProc, Err.Number, Err.Description)
  Resume PROC_EXIT
End Sub


