Attribute VB_Name = "libGlobal"
' Module     : libGlobal
' Description: Used for global declaration.
' Procedures :
' Author     : Steve Grant
'--------------------------------------------------------------------------------
Option Explicit

Private Const mcstrMod As String = "libGlobal"

Global Const gcstLockedBC = &H80000000
Global Const gcstUnLockedBC = &H80000005

' The text to display when there is a mouse over the nav buttons
Global Const mcstrTxtFirstRecord = "Goto the FIRST record."
Global Const mcstrTxtPreviousRecord = "Goto the PREVIOUS record."
Global Const mcstrTxtNewRecord = "Create a NEW record."
Global Const mcstrTxtNextRecord = "Goto the NEXT record."
Global Const mcstrTxtLastRecord = "Goto the LAST record."
Global Const mcstrTxtSaveRecord = "SAVE the current record."
Global Const mcstrTxtUndoChange = "UNDO the changes made on the current record."
Global Const mcstrTxtDeleteRecord = "DELETE the current record."
Global Const mcstrTxtRequery = "REQUERY the data. This will reorder the data " & _
                               "(usefull after the creation of a new record."
Global Const mcstrTxtEditRecord = "EDIT the record."
Global Const mcstrTxtExit = "EXIT the form."

' Used in GetTag
Enum enTagInfoTxtCtrls
  IsField = 0       ' Insert the text "Field" to indicate it's a field anything
                    ' else won't be considered as a field. Usefull if we want to
                    ' use the tag property of other controls.
  FieldName = 1     ' The name of the field in the table
  FieldType = 2     ' The type of field Alpha or Num
  FieldDup = 3      ' Are dups allowed (AllowDup) or no (NoDup)
  DefaultValue = 4  ' The default value if any
  FieldUserName = 5 ' The name that the users sees on the form
  NullsPermited = 6 ' Are nulls permited for this field (AllowNull) or (NoNull)
  NavDesc = 7       ' The description that will show on the txtNavDesc control.
End Enum
