Attribute VB_Name = "libUtilities"
' Module     : libUtilities
' Description:
' Procedures : NIF
' Author     : Steve Grant
'--------------------------------------------------------------------------------
Option Explicit

Private Const mcstrMod As String = "libUtilities"

Sub NIF()
  ' Comments  : Beeps and displays a message box with the information
  '             icon and OK button with the msg "Not in function".
  '             Usefull while developping.
  ' Parameters: Nothing
  ' Returns   : None
  ' Author    : Steve Grant
  '------------------------------------------------------------------------------
  On Error GoTo PROC_ERR

  Beep
  MsgBox "Not in function", vbInformation, "Developpement"

PROC_EXIT:
  Exit Sub

PROC_ERR:
  Resume PROC_EXIT
End Sub
