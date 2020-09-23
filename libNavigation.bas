Attribute VB_Name = "libNavigation"
' Module     : libNavigation
' Description: Library for record navigation.
' Procedures : CheckNavControls
'              DisplayCurrentRecord
'              GetTag
'              IsFirstRecord
'              IsLastRecord
'              NavDeleteRecord
'              NavEditRecord
'              NavMoveFirstRecord
'              NavMoveLastRecord
'              NavNewRecord
'              NavMoveNextRecord
'              NavMovePrevious
'              NavRequery
'              NavSaveRecord
'              NavUndoChange
'              PosNonNum
'              RecValidateTag
'              UpdateCmdControls
' Author     : Steve Grant
'--------------------------------------------------------------------------------
Option Explicit

Private Const mcstrMod As String = "libNavigation"


Function CheckNavControls(frmIn As Form, rstIn As Recordset, fEditRecord, _
                             fNewRecord) As Boolean
  ' Comments  : Checks the navigation controls on the form and disables/enables
  '             them depending on the status of the record.
  ' Parameters: frmIn      : The form on wich to update the nav controls.
  '             rstIn      : The recordset that is used with the form.
  '             fEditRecord: Is the Record in Edit Mode.
  '             fNewRecord : Is the record on New Mode
  ' Returns   : True : There is a record to display.
  '             False: There is no record to display.
  ' Author    : Steve Grant
  '------------------------------------------------------------------------------
  Const cstrProc As String = "CheckNavControls"
  'On Error GoTo PROC_ERR
  Dim fStop As Boolean
  Dim fReturn As Boolean
  
  If rstIn.AbsolutePosition = -1 Then ' There is no current record
  
    If fEditRecord Or fNewRecord Then
      UpdateCmdControls 2, 2, 2, 2, 2, 1, 1, 2, 2, 2, 3, frmIn
    Else
      UpdateCmdControls 2, 2, 1, 2, 2, 2, 2, 2, 2, 2, 3, frmIn
    End If
    fReturn = False
  
  Else ' There is at least 1 record in the table
  
    ' If the record is being edited or a new record
    If fEditRecord Or fNewRecord Then
      UpdateCmdControls 2, 2, 2, 2, 2, 1, 1, 2, 2, 2, 3, frmIn
      fStop = True
      
      If fNewRecord Then ' Don't display the record
        fReturn = False
      Else
        fReturn = True
      End If
      
    Else ' The record is not being edited or not a new record
      UpdateCmdControls 1, 1, 1, 1, 1, 2, 2, 1, 1, 1, 3, frmIn
      fStop = False
      fReturn = True
    End If
    
    ' If the record is on the last record
    If IsLastRecord(rstIn) And Not fStop Then
      UpdateCmdControls 3, 3, 3, 2, 2, 3, 3, 3, 3, 3, 3, frmIn
      fReturn = True
    End If
    
    ' If the record is on the first record
    If IsFirstRecord(rstIn) And Not fStop Then
      UpdateCmdControls 2, 2, 3, 3, 3, 3, 3, 3, 3, 3, 3, frmIn
      fReturn = True
    End If
        
  End If

  CheckNavControls = fReturn
  
PROC_EXIT:
  Exit Function
  
PROC_ERR:
  Call MsgErr(mcstrMod & "." & cstrProc, Err.Number, Err.Description)
  Resume PROC_EXIT
End Function


Sub NavDeleteRecord(frmIn As Form, rstIn As Recordset, fEditRecord, fNewRecord)
  ' Comments  : Delete the current record.
  ' Parameters: frmIn      : The form on wich to update the nav controls.
  '             rstIn      : The recordset that is used with the form.
  '             fEditRecord: Is the Record in edit Mode.
  '             fNewRecord : Is the record on New Mode
  ' Returns   : Nothing.
  ' Author    : Steve Grant
  '------------------------------------------------------------------------------
  Const cstrProc As String = "NavDeleteRecord"
  'On Error GoTo PROC_ERR
  Dim fFirstRecord As Boolean
  Dim fLastRecord As Boolean
  
  ' Get this info before the delete, because after the delete, the
  ' AbsolutePosition is set to -1
  fFirstRecord = IsFirstRecord(rstIn)
  fLastRecord = IsLastRecord(rstIn)
  
  rstIn.Delete
  
  ' If a delete occurs on the first record, we want to goto the next record.
  ' If it's not on the first record, we want to go to the previous record.
  ' This way, we can display the next logical record on the form.
  If fFirstRecord And Not fLastRecord Then
    rstIn.MoveNext
  Else
    rstIn.MovePrevious
  End If
  
  DisplayCurrentRecord frmIn, rstIn, CheckNavControls(frmIn, rstIn, _
                          fEditRecord, fNewRecord), True
  
PROC_EXIT:
  Exit Sub
  
PROC_ERR:
  Call MsgErr(mcstrMod & "." & cstrProc, Err.Number, Err.Description)
  Resume PROC_EXIT
End Sub


Sub DisplayCurrentRecord(frmIn As Form, rstIn As Recordset, _
                                    fDisplay As Boolean, fLocked As Boolean)
  ' Comments  : Write the current record to the form.
  ' Parameters: frmIn   : The form we need to put the display the data.
  '             rstIn   : The recordset we use to get the data.
  '             fDisplay: True  - We must fill it with the current record of rstIn.
  '                       False - We must empty the form.
  '             fLocked : True  - Locks the controls.
  '                       False - Unlocks the controls.
  ' Returns   : Nothing.
  ' Author    : Steve Grant
  '------------------------------------------------------------------------------
  Const cstrProc As String = "DisplayCurrentRecord"
  'On Error GoTo PROC_ERR
  Dim ctl As Control
  Dim strFieldName As String
  
    ' Loop through all the controls and do an action on them
    For Each ctl In frmIn.Controls
    
      strFieldName = GetTag(ctl.Tag, ";", FieldName)
      If strFieldName <> "NoValue" Then ' It's a field
        
        ' If fDisplay display the fields on the form else empty fields on form
        If fDisplay Then
          ctl.Text = IIf(IsNull(rstIn.Fields(strFieldName)), "", _
                         rstIn.Fields(strFieldName))
        Else
          ctl.Text = ""
        End If
  
        ' Change the BackColor and assign the Locked property depending
        ' on the lock status.
        If fLocked Then
          ctl.BackColor = gcstLockedBC
          ctl.Locked = True
        Else
          ctl.BackColor = gcstUnLockedBC
          ctl.Locked = False
        End If
        
      End If ' strFieldName
    Next ctl
     
PROC_EXIT:
  Exit Sub
  
PROC_ERR:
  Call MsgErr(mcstrMod & "." & cstrProc, Err.Number, Err.Description)
  Resume PROC_EXIT
End Sub


Function GetTag(strIn As String, chrDelimit As String, Tag As enTagInfoTxtCtrls) _
                As String
  ' Comments  : Determines if the passed string is a field. If it is, it sends the
  '             field name, if not, it sends "NoValue"
  ' Parameters: strIn   : The string to check.
  '             chrDelim: The delimiter.
  '             Tag     : IsField
  '                       FieldName
  '                       FieldValue
  '                       FieldType
  '                       FieldDup
  '                       DefaultValue
  '                       FieldUserName
  '                       NullsPermited
  '                       NavDesc
  ' Returns   : FieldName: The field name if it's a field "NoValue" otherwise.
  '             FieldType: The field Type Alpha or Num if it's a field "NoValue"
  '                        otherwise.
  '             FieldDup : DupAllowed or NoDup if it's a field, "NoValue"
  '                        otherwise.
  '             DefaultValue: The default value to use if none is present in the ctl.
  '             FieldUserName: The name the users see on the form.
  '             NullsPermited: Are nulls permited for this field AllowNull or
  '                            NoNull
  ' Author    : Steve Grant
  '------------------------------------------------------------------------------
  Const cstrProc As String = "GetTag"
  Dim strTmp As String
  Dim aintPosDelim(9) As Integer ' The pos of the different ;
  
  aintPosDelim(1) = InStr(Trim$(strIn), chrDelimit)
  aintPosDelim(2) = InStr(aintPosDelim(1) + 1, Trim$(strIn), chrDelimit)
  aintPosDelim(3) = InStr(aintPosDelim(2) + 1, Trim$(strIn), chrDelimit)
  aintPosDelim(4) = InStr(aintPosDelim(3) + 1, Trim$(strIn), chrDelimit)
  aintPosDelim(5) = InStr(aintPosDelim(4) + 1, Trim$(strIn), chrDelimit)
  aintPosDelim(6) = InStr(aintPosDelim(5) + 1, Trim$(strIn), chrDelimit)
  aintPosDelim(7) = InStr(aintPosDelim(6) + 1, Trim$(strIn), chrDelimit)
  aintPosDelim(8) = InStr(aintPosDelim(7) + 1, Trim$(strIn), chrDelimit)
  aintPosDelim(9) = InStr(aintPosDelim(8) + 1, Trim$(strIn), chrDelimit)
  
  GetTag = "NoValue"
  If aintPosDelim(1) > 1 Then ' There is at least one ;
    If Left$(strIn, aintPosDelim(1) - 1) = "Field" Then
      Select Case Tag
        Case FieldName ' Get the first section of the tag
          strTmp = Mid$(strIn, aintPosDelim(1) + 1, aintPosDelim(2) - aintPosDelim(1) - 1)
          If Len(strTmp) = 0 Then
            GetTag = "NoValue"
          Else
            GetTag = strTmp
          End If
    
        Case FieldType ' Get the second section of the tag
          strTmp = Mid$(strIn, aintPosDelim(2) + 1, aintPosDelim(3) - aintPosDelim(2) - 1)
          If Len(strTmp) = 0 Then
            GetTag = "NoValue"
          Else
            GetTag = strTmp
          End If
        
        Case FieldDup ' Get the third section of the tag
          strTmp = Mid$(strIn, aintPosDelim(3) + 1, aintPosDelim(4) - aintPosDelim(3) - 1)
          If Len(strTmp) = 0 Then
            GetTag = "NoValue"
          Else
            GetTag = strTmp
          End If
        
        Case DefaultValue ' Get the fourth section of the tag
          strTmp = Mid$(strIn, aintPosDelim(4) + 1, aintPosDelim(5) - aintPosDelim(4) - 1)
          If Len(strTmp) = 0 Then
            GetTag = "NoValue"
          Else
            GetTag = strTmp
          End If
      
        Case FieldUserName ' Get the fifth section of the tag
          strTmp = Mid$(strIn, aintPosDelim(5) + 1, aintPosDelim(6) - aintPosDelim(5) - 1)
          If Len(strTmp) = 0 Then
            GetTag = "NoValue"
          Else
            GetTag = strTmp
          End If
    
        Case NullsPermited ' Get the sixth section of the tag
          strTmp = Mid$(strIn, aintPosDelim(6) + 1, aintPosDelim(7) - aintPosDelim(6) - 1)
          If Len(strTmp) = 0 Then
            GetTag = "NoValue"
          Else
            GetTag = strTmp
          End If
    
        Case NavDesc ' Get the seventh section of the tag
          strTmp = Mid$(strIn, aintPosDelim(7) + 1, aintPosDelim(8) - aintPosDelim(7) - 1)
          If Len(strTmp) = 0 Then
            GetTag = "NoValue"
          Else
            GetTag = strTmp
          End If
          
      End Select
    End If ' Field
  End If ' aintPosDelim(1) = 0
PROC_EXIT:
  Exit Function
  
PROC_ERR:
  Call MsgErr(mcstrMod & "." & cstrProc, Err.Number, Err.Description)
  Resume PROC_EXIT
End Function


Function IsFirstRecord(rstIn As Recordset) As Boolean
  ' Comments  : Checks if the current record is the first record in the rst.
  ' Parameters: rstIn: The recordset in wich to check.
  ' Returns   : True if record is the first record, false otherwise
  ' Author    : Steve Grant
  '------------------------------------------------------------------------------
  Const cstrProc As String = "IsFirstRecord"
  'On Error GoTo PROC_ERR
  If rstIn.AbsolutePosition = 0 Then
    IsFirstRecord = True
  Else
    IsFirstRecord = False
  End If

PROC_EXIT:
  Exit Function
  
PROC_ERR:
  Call MsgErr(mcstrMod & "." & cstrProc, Err.Number, Err.Description)
  Resume PROC_EXIT
End Function


Function IsLastRecord(rstIn As Recordset) As Boolean
  ' Comments  : Checks if the current record is the last record in the rst.
  ' Parameters: rstIn: The recordset in wich to check.
  ' Returns   : True if record is the last record, false otherwise
  ' Author    : Steve Grant
  '------------------------------------------------------------------------------
  Const cstrProc As String = "IsLastRecord"
  'On Error GoTo PROC_ERR
  
  ' Check if focus is on the last record. Do a MoveNext if AbsolutePosition = -1
  ' that means there is no current record (so we we are on the last record) after,
  ' do a MovePrevious in order to not loose the record's position.
  rstIn.MoveNext
  
  If rstIn.AbsolutePosition = -1 Then
    IsLastRecord = True
  Else
    IsLastRecord = False
  End If
  
  rstIn.MovePrevious

PROC_EXIT:
  Exit Function
  
PROC_ERR:
  Call MsgErr(mcstrMod & "." & cstrProc, Err.Number, Err.Description)
  Resume PROC_EXIT
End Function


Sub NavEditRecord(frmIn As Form, rstIn As Recordset, _
                       fNewRecord As Boolean, fEditRecord As Boolean)
  ' Comments  : Edit the record. This is all handled by the DisplayCurrentRecord
  '             proc. This proc is a wrapper in case we need to had something
  '             else later on.
  ' Parameters: frmIn      : The form on wich to do the validation.
  '             rstIn      : The recordset that is used with the form.
  '             fEditRecord: Is the Record in edit Mode.
  '             fNewRecord : Is the record on New Mode
  ' Returns   : Nothing.
  ' Author    : Steve Grant
  '------------------------------------------------------------------------------
  Const cstrProc As String = "NavEditRecord"
  'On Error GoTo PROC_ERR
  
  DisplayCurrentRecord frmIn, rstIn, CheckNavControls(frmIn, rstIn, fEditRecord, _
                       fNewRecord), False
                       
PROC_EXIT:
  Exit Sub
  
PROC_ERR:
  Call MsgErr(mcstrMod & "." & cstrProc, Err.Number, Err.Description)
  Resume PROC_EXIT
End Sub


Sub NavMoveFirstRecord(frmIn As Form, rstIn As Recordset, _
                       fNewRecord As Boolean, fEditRecord As Boolean)
  ' Comments  : Move to the first record.
  ' Parameters: frmIn      : The form on wich to do the validation.
  '             rstIn      : The recordset that is used with the form.
  '             fEditRecord: Is the Record in edit Mode.
  '             fNewRecord : Is the record on New Mode
  ' Returns   : Nothing.
  ' Author    : Steve Grant
  '------------------------------------------------------------------------------
  Const cstrProc As String = "NavMoveFirstRecord"
  'On Error GoTo PROC_ERR
  
  rstIn.MoveFirst
  DisplayCurrentRecord frmIn, rstIn, CheckNavControls(frmIn, rstIn, fEditRecord, _
                       fNewRecord), True
                       
PROC_EXIT:
  Exit Sub
  
PROC_ERR:
  Call MsgErr(mcstrMod & "." & cstrProc, Err.Number, Err.Description)
  Resume PROC_EXIT
End Sub


Sub NavMoveLastRecord(frmIn As Form, rstIn As Recordset, _
                       fNewRecord As Boolean, fEditRecord As Boolean)
  ' Comments  : Move to the last record.
  ' Parameters: frmIn      : The form on wich to do the validation.
  '             rstIn      : The recordset that is used with the form.
  '             fEditRecord: Is the Record in edit Mode.
  '             fNewRecord : Is the record on New Mode
  ' Returns   : Nothing.
  ' Author    : Steve Grant
  '------------------------------------------------------------------------------
  Const cstrProc As String = "NavMoveLastRecord"
  'On Error GoTo PROC_ERR
  
  rstIn.MoveLast
  DisplayCurrentRecord frmIn, rstIn, CheckNavControls(frmIn, rstIn, fEditRecord, _
                       fNewRecord), True
                       
PROC_EXIT:
  Exit Sub
  
PROC_ERR:
  Call MsgErr(mcstrMod & "." & cstrProc, Err.Number, Err.Description)
  Resume PROC_EXIT
End Sub


Sub NavNewRecord(frmIn As Form, rstIn As Recordset, _
                       fNewRecord As Boolean, fEditRecord As Boolean)
  ' Comments  : New record. This is all handled by the DisplayCurrentRecord
  '             proc. This proc is a wrapper in case we need to had something
  '             else later on.
  ' Parameters: frmIn      : The form on wich to do the validation.
  '             rstIn      : The recordset that is used with the form.
  '             fEditRecord: Is the Record in edit Mode.
  '             fNewRecord : Is the record on New Mode
  ' Returns   : Nothing.
  ' Author    : Steve Grant
  '------------------------------------------------------------------------------
  Const cstrProc As String = "NavNewRecord"
  'On Error GoTo PROC_ERR
  
  DisplayCurrentRecord frmIn, rstIn, CheckNavControls(frmIn, rstIn, fEditRecord, _
                       fNewRecord), False
                       
PROC_EXIT:
  Exit Sub
  
PROC_ERR:
  Call MsgErr(mcstrMod & "." & cstrProc, Err.Number, Err.Description)
  Resume PROC_EXIT
End Sub


Sub NavMoveNextRecord(frmIn As Form, rstIn As Recordset, _
                       fNewRecord As Boolean, fEditRecord As Boolean)
  ' Comments  : Move to the next record.
  ' Parameters: frmIn      : The form on wich to do the validation.
  '             rstIn      : The recordset that is used with the form.
  '             fEditRecord: Is the Record in edit Mode.
  '             fNewRecord : Is the record on New Mode
  ' Returns   : Nothing.
  ' Author    : Steve Grant
  '------------------------------------------------------------------------------
  Const cstrProc As String = "NavMoveNextRecord"
  'On Error GoTo PROC_ERR
  
  rstIn.MoveNext
  DisplayCurrentRecord frmIn, rstIn, CheckNavControls(frmIn, rstIn, fEditRecord, _
                       fNewRecord), True
                       
PROC_EXIT:
  Exit Sub
  
PROC_ERR:
  Call MsgErr(mcstrMod & "." & cstrProc, Err.Number, Err.Description)
  Resume PROC_EXIT
End Sub


Sub NavMovePreviousRecord(frmIn As Form, rstIn As Recordset, _
                          fNewRecord As Boolean, fEditRecord As Boolean)
  ' Comments  : Move to the previous record.
  ' Parameters: frmIn      : The form on wich to do the validation.
  '             rstIn      : The recordset that is used with the form.
  '             fEditRecord: Is the Record in edit Mode.
  '             fNewRecord : Is the record on New Mode
  ' Returns   : Nothing.
  ' Author    : Steve Grant
  '------------------------------------------------------------------------------
  Const cstrProc As String = "NavMovePreviousRecord"
  'On Error GoTo PROC_ERR
  
  rstIn.MovePrevious
  DisplayCurrentRecord frmIn, rstIn, CheckNavControls(frmIn, rstIn, fEditRecord, _
                       fNewRecord), True
                       
PROC_EXIT:
  Exit Sub
  
PROC_ERR:
  Call MsgErr(mcstrMod & "." & cstrProc, Err.Number, Err.Description)
  Resume PROC_EXIT
End Sub


Sub NavRequery(frmIn As Form, rstIn As Recordset, _
                       fNewRecord As Boolean, fEditRecord As Boolean)
  ' Comments  : Requery the recordset.
  ' Parameters: frmIn      : The form on wich to do the validation.
  '             rstIn      : The recordset that is used with the form.
  '             fEditRecord: Is the Record in edit Mode.
  '             fNewRecord : Is the record on New Mode
  ' Returns   : Nothing.
  ' Author    : Steve Grant
  '------------------------------------------------------------------------------
  Const cstrProc As String = "NavRequery"
  'On Error GoTo PROC_ERR
  
  rstIn.Requery
  DisplayCurrentRecord frmIn, rstIn, CheckNavControls(frmIn, rstIn, fEditRecord, _
                       fNewRecord), True
                       
PROC_EXIT:
  Exit Sub
  
PROC_ERR:
  Call MsgErr(mcstrMod & "." & cstrProc, Err.Number, Err.Description)
  Resume PROC_EXIT
End Sub


Sub NavSaveRecord(frmIn As Form, rstIn As Recordset, _
                       fNewRecord As Boolean, fEditRecord As Boolean)
  ' Comments  : Save record.
  ' Parameters: frmIn      : The form on wich to do the validation.
  '             rstIn      : The recordset that is used with the form.
  '             fEditRecord: Is the Record in edit Mode.
  '             fNewRecord : Is the record on New Mode
  ' Returns   : Nothing.
  ' Author    : Steve Grant
  '------------------------------------------------------------------------------
  Const cstrProc As String = "NavSaveRecord"
  'On Error GoTo PROC_ERR
  Dim ctl As Control
  Dim strFieldName As String
  
  ' *** EDIT MODE ***
  If fEditRecord Then ' The record is in edit mode
    rstIn.Edit
      For Each ctl In frmIn.Controls
        strFieldName = GetTag(ctl.Tag, ";", FieldName)
        If strFieldName <> "NoValue" Then ' It's a field
          rstIn.Fields(strFieldName) = ctl.Text
        End If ' strFieldName
      Next ctl
    rstIn.Update
    fEditRecord = False
  End If ' fEditRecord

  ' *** NEW RECORD ***
  If fNewRecord Then   ' The record is a new record
    rstIn.AddNew
      For Each ctl In frmIn.Controls
        strFieldName = GetTag(ctl.Tag, ";", FieldName)
        If strFieldName <> "NoValue" Then ' It's a field
          rstIn.Fields(strFieldName) = ctl.Text
        End If ' strFieldName
      Next ctl
    rstIn.Update
    ' MoveLast because that's where the new record is
    rstIn.MoveLast
    fNewRecord = False
  End If ' fNewRecord
  
  DisplayCurrentRecord frmIn, rstIn, CheckNavControls(frmIn, rstIn, fEditRecord, _
                       fNewRecord), True

PROC_EXIT:
  Exit Sub
  
PROC_ERR:
  Call MsgErr(mcstrMod & "." & cstrProc, Err.Number, Err.Description)
  Resume PROC_EXIT
End Sub

Sub NavUndoChange(frmIn As Form, rstIn As Recordset, _
                       fNewRecord As Boolean, fEditRecord As Boolean)
  ' Comments  : Undo changes. This is all handled by the DisplayCurrentRecord
  '             proc. This proc is a wrapper in case we need to had something
  '             else later on.
  ' Parameters: frmIn      : The form on wich to do the validation.
  '             rstIn      : The recordset that is used with the form.
  '             fEditRecord: Is the Record in edit Mode.
  '             fNewRecord : Is the record on New Mode
  ' Returns   : Nothing.
  ' Author    : Steve Grant
  '------------------------------------------------------------------------------
  Const cstrProc As String = "NavUndoChange"
  'On Error GoTo PROC_ERR
  
  DisplayCurrentRecord frmIn, rstIn, CheckNavControls(frmIn, rstIn, fEditRecord, _
                       fNewRecord), True
    
PROC_EXIT:
  Exit Sub
  
PROC_ERR:
  Call MsgErr(mcstrMod & "." & cstrProc, Err.Number, Err.Description)
  Resume PROC_EXIT
End Sub


Function PosNonNum(strIn As String) As Integer
  ' Comments  : Finds the first occurence of a non numeric character in a string.
  ' Parameters: strIn : The string in wich to search.
  ' Returns   : The intPos of the first non-numeric character. If the function
  '             didn't find anything, return 0.
  ' Created   : Steve Grant
  '------------------------------------------------------------------------------
  Const cstrProc As String = "PosNonNum"
  'On Error GoTo PROC_ERR
  
  Dim intPos As Integer
  Dim strArray() As String
  Dim intLength As Integer
  Dim i As Integer
  
  If Len(strIn) > 0 Then ' There is something in the string
    intLength = Len(strIn)
    ReDim strArray(intLength)
    
    For i = 1 To intLength
      strArray(i) = Mid$(strIn, i, 1)
    Next i
    
    intPos = 0
    i = 1
    Do Until (i = intLength + 1) Or (intPos > 0)
      If strArray(i) < "0" Or strArray(i) > "9" Then   ' It's not numeric
        intPos = i
      End If
      i = i + 1
    Loop
  
  Else ' The strIn is empty
    intPos = 0
  End If
  
  PosNonNum = intPos
PROC_EXIT:
  Exit Function
  
PROC_ERR:
  Call MsgErr(mcstrMod & "." & cstrProc, Err.Number, Err.Description)
  Resume PROC_EXIT
End Function


Function RecValidateTag(frmIn As Form, rstIn As Recordset, _
                        fNewRecord As Boolean, fEditRecord As Boolean) As String
  ' Comments  : Validates the record. This function uses the information located
  '             in the tag property of the txt controls.
  ' Parameters: frmIn      : The form on wich to do the validation.
  '             rstIn      : The recordset that is used with the form.
  '             fEditRecord: Is the Record in edit Mode.
  '             fNewRecord : Is the record on New Mode
  ' Returns   : "Valid" if passed validation, a msg containing the error(s) found
  '             otherwise.
  ' Author    : Steve Grant
  '------------------------------------------------------------------------------
  Const cstrProc As String = "RecValidateTag"
  'On Error GoTo PROC_ERR
  Dim fValidate As Boolean
  Dim ctl As Control
  Dim strMsg As String
  Dim strCriteria As String
  Dim strCurrent As String
  Dim strFieldName As String     ' The name of the field in the table
  Dim strFieldType As String     ' The type of field Alpha or Num
  Dim strFieldDup As String      ' Are dups allowed (AllowDup) or no (NoDup)
  Dim strDefaultValue  As String ' The default value if any
  Dim strFieldUserName As String ' The name that the users sees on the form
  Dim strNullsPermited As String ' Are nulls permited for this field AllowNull
                                 ' or NoNull
  Dim varBookmark As Variant
  
  strMsg = "The record could not be saved because of the folowing:" & vbCrLf & vbCrLf
  fValidate = True
  
  On Error Resume Next ' If there is no record in the table we get an error
  varBookmark = rstIn.Bookmark
  On Error GoTo 0
  
  ' Go thru each record and validate it with the different tag options on the
  ' txtField.
  For Each ctl In frmIn.Controls
    
    strFieldName = GetTag(ctl.Tag, ";", FieldName)
    If strFieldName <> "NoValue" Then ' It's a field so do some validation
      
      ' Get the other info
      strFieldType = GetTag(ctl.Tag, ";", FieldType)
      strFieldDup = GetTag(ctl.Tag, ";", FieldDup)
      strDefaultValue = GetTag(ctl.Tag, ";", DefaultValue)
      strFieldUserName = GetTag(ctl.Tag, ";", FieldUserName)
      strNullsPermited = GetTag(ctl.Tag, ";", NullsPermited)
      
      strCurrent = ctl.Text ' Keep track of the current value, because it will be
                            ' changed when we do FindFirst.
      
      ' *** Validate NEW RECORD ***
      ' If it's a new record, do not allow DUPS
      If fNewRecord And strFieldDup = "NoDup" Then
        strCriteria = "" & strFieldName & " = '" & ctl.Text & "'"
        rstIn.FindFirst strCriteria
        
        If Not rstIn.NoMatch Then ' There is a Match
          fValidate = False
          strMsg = strMsg & "The " & strFieldUserName & " already exists." & vbCrLf
        End If ' rstIn.NoMatch
      
      End If ' fNewRecord
    
      
      ' *** Validate EDIT RECORD ***
      ' If a record is being edited we must check all fields that do not allow
      ' DUPS. But when we do the validation, we must compare to the initial
      ' value of the record.
      If fEditRecord And strFieldDup = "NoDup" Then
        If rstIn.Fields(strFieldName) <> strCurrent Then
          strCriteria = "" & strFieldName & " = '" & ctl.Text & "'"
          rstIn.FindFirst strCriteria
          
          If Not rstIn.NoMatch Then ' There is a Match but different than strCurrent
            fValidate = False
            strMsg = strMsg & "You tried to rename the " & strFieldUserName & _
                              ". But this name already exists." & vbCrLf
          End If ' rstIn.NoMatch
        
        End If ' rstIn.fields <> strCurrent
      End If ' fEditRecord
  
    
      ' *** Validate NULLS ***
      If strNullsPermited = "NoNull" Then
        If Len(Trim$(ctl.Text)) = 0 Then
          strMsg = strMsg & "The " & strFieldUserName & " cannot be empty." & vbCrLf
          fValidate = False
        End If ' Len
      End If ' strNullsPermited
    
        
      ' *** Assign DEFAULT VALUES ***
      If strDefaultValue <> "NoValue" Then ' There is a default value
        If Len(ctl.Text) = 0 Then ctl.Text = strDefaultValue
      End If ' strDefaultValue
      
      
      ' *** Validate NUMERIC ***
      ' If it's alpha, we don't need to validate it. If it's num, we need to
      ' check that the user has entered only numerics.
      If strFieldType = "Num" Then
        If PosNonNum(ctl.Text) > 0 Then
          strMsg = strMsg & "The " & strFieldUserName & " must be numeric." & vbCrLf
          fValidate = False
        End If ' PosNonNum
      End If ' strFieldType
    
    End If ' strField is field
  Next ctl
  
  ' If the record is valid, destroy the info in strMsg and put "Valid", that's
  ' what will be sent by the function.
  If fValidate Then ' Doesn't pass the validation
    strMsg = "Valid"
  End If
  
  On Error Resume Next
  rstIn.Bookmark = varBookmark
  On Error GoTo 0
  
  RecValidateTag = strMsg
  
PROC_EXIT:
  Exit Function
  
PROC_ERR:
  Call MsgErr(mcstrMod & "." & cstrProc, Err.Number, Err.Description)
  Resume PROC_EXIT
End Function


Sub UpdateCmdControls(bFirstRecord As Byte, _
                      bPreviousRecord As Byte, _
                      bNewRecord As Byte, _
                      bNextRecord As Byte, _
                      bLastRecord As Byte, _
                      bSaveRecord As Byte, _
                      bUndoChange As Byte, _
                      bDeleteRecord As Byte, _
                      bEditRecord As Byte, _
                      bRequery As Byte, _
                      bExit As Byte, _
                      frmIn As Form)
  ' Comments  : Update the Command controls on the form and synchronize with
  '             the menu options.
  ' Parameters: bFirstRecord   : 1 = true 2= false 3 = NoChange
  '             bPreviousRecord: 1 = true 2= false 3 = NoChange
  '             bNewRecord     : 1 = true 2= false 3 = NoChange
  '             bNextRecord    : 1 = true 2= false 3 = NoChange
  '             bLastRecord    : 1 = true 2= false 3 = NoChange
  '             bSaveRecord    : 1 = true 2= false 3 = NoChange
  '             bUndoChange    : 1 = true 2= false 3 = NoChange
  '             bDeleteRecord  : 1 = true 2= false 3 = NoChange
  '             bEditRecord    : 1 = true 2= false 3 = NoChange
  '             bRequery       : 1 = true 2= false 3 = NoChange
  '             frmIn          : Form on wich to do the changes.
  ' Returns   : Nothing.
  ' Author    : Steve Grant
  '------------------------------------------------------------------------------
  Const cstrProc As String = "UpdateCmdControls"
  'On Error GoTo PROC_ERR
    
  Select Case bFirstRecord
    Case 1
      frmIn.cmdNav(0).Enabled = True
      frmIn.mnuNavFirstRecord.Enabled = True
    Case 2
      frmIn.cmdNav(0).Enabled = False
      frmIn.mnuNavFirstRecord.Enabled = False
    Case 3
      ' No Change
  End Select
  
  Select Case bPreviousRecord
    Case 1
      frmIn.cmdNav(1).Enabled = True
      frmIn.mnuNavPreviousRecord.Enabled = True
    Case 2
      frmIn.cmdNav(1).Enabled = False
      frmIn.mnuNavPreviousRecord.Enabled = False
    Case 3
      ' No Change
  End Select
  
  Select Case bNewRecord
    Case 1
      frmIn.cmdNav(2).Enabled = True
      frmIn.mnuNavNewRecord.Enabled = True
    Case 2
      frmIn.cmdNav(2).Enabled = False
      frmIn.mnuNavNewRecord.Enabled = False
    Case 3
      ' No Change
  End Select
    
  Select Case bNextRecord
    Case 1
      frmIn.cmdNav(3).Enabled = True
      frmIn.mnuNavNextRecord.Enabled = True
    Case 2
      frmIn.cmdNav(3).Enabled = False
      frmIn.mnuNavNextRecord.Enabled = False
    Case 3
      ' No Change
  End Select
  
  Select Case bLastRecord
    Case 1
      frmIn.cmdNav(4).Enabled = True
      frmIn.mnuNavLastRecord.Enabled = True
    Case 2
      frmIn.cmdNav(4).Enabled = False
      frmIn.mnuNavLastRecord.Enabled = False
    Case 3
      ' No Change
  End Select
  
  Select Case bSaveRecord
    Case 1
      frmIn.cmdNav(5).Enabled = True
      frmIn.mnuNavSaveRecord.Enabled = True
    Case 2
      frmIn.cmdNav(5).Enabled = False
      frmIn.mnuNavSaveRecord.Enabled = False
    Case 3
      ' No Change
  End Select
  
  Select Case bUndoChange
    Case 1
      frmIn.cmdNav(6).Enabled = True
      frmIn.mnuNavUndoChange.Enabled = True
    Case 2
      frmIn.cmdNav(6).Enabled = False
      frmIn.mnuNavUndoChange.Enabled = False
    Case 3
      ' No Change
  End Select
  
  Select Case bDeleteRecord
    Case 1
      frmIn.cmdNav(7).Enabled = True
      frmIn.mnuNavDeleteRecord.Enabled = True
    Case 2
      frmIn.cmdNav(7).Enabled = False
      frmIn.mnuNavDeleteRecord.Enabled = False
    Case 3
      ' No Change
  End Select
  
  Select Case bRequery
    Case 1
      frmIn.cmdNav(8).Enabled = True
      frmIn.mnuNavRequery.Enabled = True
    Case 2
      frmIn.cmdNav(8).Enabled = False
      frmIn.mnuNavRequery.Enabled = False
    Case 3
      ' No Change
  End Select
  
  Select Case bEditRecord
    Case 1
      frmIn.cmdNav(9).Enabled = True
      frmIn.mnuNavEditRecord.Enabled = True
    Case 2
      frmIn.cmdNav(9).Enabled = False
      frmIn.mnuNavEditRecord.Enabled = False
    Case 3
      ' No Change
  End Select
  
  Select Case bExit
    Case 1
      frmIn.cmdNav(10).Enabled = True
      frmIn.mnuFileExit_Click.Enabled = True
    Case 2
      frmIn.cmdNav(10).Enabled = False
      frmIn.mnuFileExit_Click.Enabled = False
    Case 3
      ' No Change
  End Select
  
PROC_EXIT:
  Exit Sub
  
PROC_ERR:
  Call MsgErr(mcstrMod & "." & cstrProc, Err.Number, Err.Description)
  Resume PROC_EXIT
End Sub

