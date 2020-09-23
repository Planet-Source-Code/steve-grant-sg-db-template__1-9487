VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Clans"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   6990
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   6990
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtData 
      Height          =   285
      Index           =   1
      Left            =   5760
      TabIndex        =   14
      Tag             =   "Field;CL_Year;Num;DupAllowed;0;Year;AllowNull;The first year in wich this Clan can be used. Use 0 if always available.;"
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txtData 
      Height          =   285
      Index           =   0
      Left            =   720
      TabIndex        =   13
      Tag             =   "Field;CL_Name;Alpha;NoDup;;Name;NoNull;Clan Name;"
      Top             =   120
      Width           =   3975
   End
   Begin VB.TextBox txtNavDesc 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   3960
      Width           =   4575
   End
   Begin VB.CommandButton cmdNav 
      Appearance      =   0  'Flat
      Caption         =   "First"
      Height          =   495
      Index           =   0
      Left            =   120
      Picture         =   "frmMain.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   3360
      Width           =   735
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   "Previous"
      Height          =   495
      Index           =   1
      Left            =   840
      Picture         =   "frmMain.frx":0075
      Style           =   1  'Graphical
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3360
      Width           =   735
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   "New"
      Height          =   495
      Index           =   2
      Left            =   1560
      Picture         =   "frmMain.frx":00DF
      Style           =   1  'Graphical
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   3360
      Width           =   735
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   "Next"
      Height          =   495
      Index           =   3
      Left            =   2280
      Picture         =   "frmMain.frx":0159
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   3360
      Width           =   735
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   "Last"
      Height          =   495
      Index           =   4
      Left            =   3000
      Picture         =   "frmMain.frx":01C3
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   3360
      Width           =   735
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   "Save"
      Height          =   495
      Index           =   5
      Left            =   4080
      Picture         =   "frmMain.frx":0239
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3360
      Width           =   615
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   "Undo"
      Height          =   495
      Index           =   6
      Left            =   4800
      Picture         =   "frmMain.frx":02CA
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3360
      Width           =   615
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   "Delete"
      Height          =   495
      Index           =   7
      Left            =   5520
      Picture         =   "frmMain.frx":033C
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3360
      Width           =   615
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   "Requery"
      Height          =   495
      Index           =   8
      Left            =   4800
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   "Edit"
      Height          =   495
      Index           =   9
      Left            =   5880
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   "Exit"
      Height          =   495
      Index           =   10
      Left            =   6240
      Picture         =   "frmMain.frx":03C5
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3360
      Width           =   615
   End
   Begin VB.TextBox txtData 
      Height          =   2565
      Index           =   2
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Tag             =   "Field;CL_Note;Alpha;DupAllowed;;Note;AllowNull;Note on this Clan.;"
      Top             =   720
      Width           =   6735
   End
   Begin VB.Label lblData 
      Caption         =   "Note:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   17
      Top             =   480
      Width           =   495
   End
   Begin VB.Label lblData 
      Caption         =   "Year (0=all):"
      Height          =   255
      Index           =   1
      Left            =   4800
      TabIndex        =   16
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lblData 
      Caption         =   "Name:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   495
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuNav 
      Caption         =   "&Navigation"
      Begin VB.Menu mnuNavFirstRecord 
         Caption         =   "&First Record"
      End
      Begin VB.Menu mnuNavPreviousRecord 
         Caption         =   "&Previous Record"
      End
      Begin VB.Menu mnuNavNewRecord 
         Caption         =   "&New Record"
      End
      Begin VB.Menu mnuNavNextRecord 
         Caption         =   "N&ext Record"
      End
      Begin VB.Menu mnuNavLastRecord 
         Caption         =   "&Last Record"
      End
      Begin VB.Menu mnuNavSaveRecord 
         Caption         =   "&Save Record"
      End
      Begin VB.Menu mnuNavUndoChange 
         Caption         =   "&Undo Change"
      End
      Begin VB.Menu mnuNavDeleteRecord 
         Caption         =   "&DeleteRecord"
      End
      Begin VB.Menu mnuNavSeparator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNavEditRecord 
         Caption         =   "Ed&it Record"
      End
      Begin VB.Menu mnuNavSeparator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNavRequery 
         Caption         =   "&Requery"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpManual 
         Caption         =   "&Manual"
      End
      Begin VB.Menu mnuHelpSeparaor 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Module     : frmMain
' Description:
' Procedures : cbfRecValidate
'              cmdNav
'              Form_Load
'              Form_Unload
'              mnuFileExit_Click
'              mnuHelpAbout_Click
'              mnuHelpManual_Click
'              mnuHelpRegister_Click
'              mnuNavDeleteRecord_Click
'              mnuNavEditRecord_Click
'              mnuNavFirstRecord_Click
'              mnuNavLastRecord_Click
'              mnuNavNewRecord_Click
'              mnuNavNextRecord_Click
'              mnuNavPreviousRecord_Click
'              mnuNavSaveRecord_Click
'              mnuNavUndoChange_Click
' Author     : Steve Grant
'--------------------------------------------------------------------------------
Option Explicit

Private Const mcstrMod As String = "frmMain"

' The db the form will use.
Private Const mcstrDBName As String = "\DB.mdb"

' The SQL the form will use.
Private Const mcstrSQL = "SELECT tblClans.CL_ID, tblClans.CL_Name, " & _
                         "tblClans.CL_Year, tblClans.CL_Note " & _
                         "FROM tblClans;"

Private mdbs As Database
Private mrst As Recordset

Private mfEditRecord As Boolean ' True if the record is being edited
Private mfNewRecord As Boolean  ' True if the record is a new record


Private Function cbfRecValidate() As Boolean
  ' Comments  : Validates the record. This function will call the RecValidateTag
  '             function that will validate the record based on the info located
  '             in the tag property of the txtFields.
  '             We use cbfRecValidate, to give us a simple mean of doing some more
  '             validation on the record if we need it.
  '             If we need to do more validation, we should insert it after having
  '             called RecValidateTag, because this function will keep the data
  '             integrity of the database.
  ' Parameters: None.
  ' Returns   : True if passed validation, false otherwise.
  ' Author    : Steve Grant
  '------------------------------------------------------------------------------
  Const cstrProc As String = "cbfRecValidate"
  'On Error GoTo PROC_ERR
  Dim fValidate As Boolean
  Dim strMsg
  
  strMsg = RecValidateTag(Me, mrst, mfNewRecord, mfEditRecord)
  
  If strMsg <> "Valid" Then ' The record did not pass validation so
                            ' display the error msg
    MsgBox strMsg, vbOKOnly + vbExclamation, _
           "Validation error"
    fValidate = False
  Else
    fValidate = True
  End If
  
  cbfRecValidate = fValidate
  
PROC_EXIT:
  Exit Function
  
PROC_ERR:
  Call MsgErr(mcstrMod & "." & cstrProc, Err.Number, Err.Description)
  Resume PROC_EXIT
End Function


Private Sub cmdNav_Click(Index As Integer)
  ' Comments  : Navigation buttons.
  '             0 : First Record
  '             1 : Previous Record
  '             2 : New Record
  '             3 : Next Record
  '             4 : Last Record
  '             5 : Save Record
  '             6 : Undo Change
  '             7 : Delete Record
  '             8 : Requery
  '             9 : Edit Record
  '            10 : Exit the form
  ' Parameters: None.
  ' Returns   : Nothing.
  ' Author    : Steve Grant
  '------------------------------------------------------------------------------
  Const cstrProc As String = "cmdNav_Click"
  'On Error GoTo PROC_ERR

  Select Case Index
    Case 0  ' First Record
      NavMoveFirstRecord Me, mrst, mfNewRecord, mfEditRecord

    Case 1  ' Previous Record
      NavMovePreviousRecord Me, mrst, mfNewRecord, mfEditRecord
    
    Case 2  ' New Record
      mfNewRecord = True
      NavNewRecord Me, mrst, mfNewRecord, mfEditRecord
    
    Case 3  ' Next Record
      NavMoveNextRecord Me, mrst, mfNewRecord, mfEditRecord
    
    Case 4  ' Last Record
      NavMoveLastRecord Me, mrst, mfNewRecord, mfEditRecord
 
    Case 5  ' Save Record
      If cbfRecValidate Then ' The record is valid, so proceed
        NavSaveRecord Me, mrst, mfNewRecord, mfEditRecord
        mfNewRecord = False
        mfEditRecord = False
      End If
    
    Case 6  ' Undo Change
      ' When we undo, we are sure that we are not on a New Record or Editing one
      mfEditRecord = False
      mfNewRecord = False
      NavUndoChange Me, mrst, mfNewRecord, mfEditRecord
    
    Case 7  ' Delete Record
      NavDeleteRecord Me, mrst, mfEditRecord, mfNewRecord
    
    Case 8  ' Requery
      NavRequery Me, mrst, mfNewRecord, mfEditRecord
    
    Case 9  ' Edit Record
      mfEditRecord = True
      NavEditRecord Me, mrst, mfNewRecord, mfEditRecord
      
    Case 10 ' Exit the form
      Unload Me
  
  End Select
  
PROC_EXIT:
  Exit Sub
  
PROC_ERR:
  Call MsgErr(mcstrMod & "." & cstrProc, Err.Number, Err.Description)
  Resume PROC_EXIT
End Sub


Private Sub cmdNav_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Comments  : Navigation buttons.
  '             0 : First Record
  '             1 : Previous Record
  '             2 : New Record
  '             3 : Next Record
  '             4 : Last Record
  '             5 : Save Record
  '             6 : Undo Change
  '             7 : Delete Record
  '             8 : Requery
  '             9 : Edit Record
  ' Parameters: None.
  ' Returns   : Nothing.
  ' Author    : Steve Grant
  '------------------------------------------------------------------------------
  Const cstrProc As String = "cmdNav_MouseMove"
  'On Error GoTo PROC_ERR

    Select Case Index
    Case 0  ' First Record
      txtNavDesc.Text = mcstrTxtFirstRecord

    Case 1  ' Previous Record
      txtNavDesc.Text = mcstrTxtPreviousRecord

    Case 2  ' New Record
      txtNavDesc.Text = mcstrTxtNewRecord
    
    Case 3  ' Next Record
      txtNavDesc.Text = mcstrTxtNextRecord
    
    Case 4  ' Last Record
      txtNavDesc.Text = mcstrTxtLastRecord
 
    Case 5  ' Save Record
      txtNavDesc.Text = mcstrTxtSaveRecord
    
    Case 6  ' Undo Change
      txtNavDesc.Text = mcstrTxtUndoChange
    
    Case 7  ' Delete Record
      txtNavDesc.Text = mcstrTxtDeleteRecord
    
    Case 8  ' Requery
      txtNavDesc.Text = mcstrTxtRequery
    
    Case 9  ' Edit Record
      txtNavDesc.Text = mcstrTxtEditRecord
    
    Case 10 ' Exit
      txtNavDesc.Text = mcstrTxtExit
  
  End Select

PROC_EXIT:
  Exit Sub
  
PROC_ERR:
  Call MsgErr(mcstrMod & "." & cstrProc, Err.Number, Err.Description)
  Resume PROC_EXIT
End Sub


Private Sub Form_Load()
  ' Comments  : When the form loads, do the following:
  '             Assign the mdbs and mrst.
  ' Parameters: None.
  ' Returns   : Nothing.
  ' Author    : Steve Grant
  '------------------------------------------------------------------------------
  Const cstrProc As String = "Form_Load"
  'On Error GoTo PROC_ERR
  
  Set mdbs = OpenDatabase(App.Path & mcstrDBName)

  Set mrst = mdbs.OpenRecordset(mcstrSQL)
  
  mfEditRecord = False
  mfNewRecord = False
  
  DisplayCurrentRecord Me, mrst, CheckNavControls(Me, mrst, mfEditRecord, _
                          mfNewRecord), True
  
PROC_EXIT:
  Exit Sub
  
PROC_ERR:
  Call MsgErr(mcstrMod & "." & cstrProc, Err.Number, Err.Description)
  Resume PROC_EXIT
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Comments  : When the mous is not on a button, reset txtNavDesc to "".
  ' Parameters: None.
  ' Returns   : Nothing.
  ' Author    : Steve Grant
  '------------------------------------------------------------------------------
  Const cstrProc As String = "Form_MouseMove"
  'On Error GoTo PROC_ERR
  
  txtNavDesc = ""
  
PROC_EXIT:
  Exit Sub
  
PROC_ERR:
  Call MsgErr(mcstrMod & "." & cstrProc, Err.Number, Err.Description)
  Resume PROC_EXIT
End Sub


Private Sub Form_Unload(Cancel As Integer)
  ' Comments  : When the form unloads, do the following:
  '             Close the mdbs and mrst.
  ' Parameters: None.
  ' Returns   : Nothing.
  ' Author    : Steve Grant
  '------------------------------------------------------------------------------
  Const cstrProc As String = "Form_Unload"
  'On Error GoTo PROC_ERR
  Dim intRet As Integer
  Dim strMsg As String
  
  If mfNewRecord Or mfEditRecord Then
    strMsg = "The record is either a NEW record or is being EDITED. " & _
             "If you exit at this moment, you will lose all the " & _
             "changes made to the current record." & vbCrLf & vbCrLf & _
             "If you wish to exit without saving (and lose the " & _
             "modifications you made) press 'OK'." & vbCrLf & vbCrLf & _
             "If you do not wish to exit at this time, press " & _
             "'Cancel'." & vbCrLf & vbCrLf

    intRet = MsgBox(strMsg, vbOKCancel + vbInformation + _
                     vbDefaultButton2, "New or Edited record")
    Select Case intRet
      Case vbOK: ' The user chose to quit and not save
        mrst.Close
        Set mrst = Nothing
        mdbs.Close
        Set mdbs = Nothing
      
      Case vbCancel: ' The user chose not to exit
        Cancel = True
      Case Else
    End Select
    
  Else
    mrst.Close
    Set mrst = Nothing
    mdbs.Close
    Set mdbs = Nothing
  End If ' mfNewRecord Or mfEditRecord
  
PROC_EXIT:
  Exit Sub
  
PROC_ERR:
  Call MsgErr(mcstrMod & "." & cstrProc, Err.Number, Err.Description)
  Resume PROC_EXIT
End Sub


Private Sub mnuFileExit_Click()
  ' Comments  : Menu options.
  ' Parameters: None.
  ' Returns   : Nothing.
  ' Author    : Steve Grant
  '------------------------------------------------------------------------------
  Const cstrProc As String = "mnuFileExit_Click"
  'On Error GoTo PROC_ERR
  Unload Me
  
PROC_EXIT:
  Exit Sub
  
PROC_ERR:
  Call MsgErr(mcstrMod & "." & cstrProc, Err.Number, Err.Description)
  Resume PROC_EXIT
End Sub


Private Sub mnuHelpAbout_Click()
  ' Comments  : Menu options.
  ' Parameters: None.
  ' Returns   : Nothing.
  ' Author    : Steve Grant
  '------------------------------------------------------------------------------
  Const cstrProc As String = "mnuHelpAbout_Click"
  'On Error GoTo PROC_ERR
  
  DisplayAbout

PROC_EXIT:
  Exit Sub
  
PROC_ERR:
  Call MsgErr(mcstrMod & "." & cstrProc, Err.Number, Err.Description)
  Resume PROC_EXIT
End Sub


Private Sub mnuHelpManual_Click()
  ' Comments  : Menu options.
  ' Parameters: None.
  ' Returns   : Nothing.
  ' Author    : Steve Grant
  '------------------------------------------------------------------------------
  Const cstrProc As String = "mnuHelpManual_Click"
  'On Error GoTo PROC_ERR
  
  DisplayManual

PROC_EXIT:
  Exit Sub
  
PROC_ERR:
  Call MsgErr(mcstrMod & "." & cstrProc, Err.Number, Err.Description)
  Resume PROC_EXIT
End Sub


Private Sub mnuNavDeleteRecord_Click()
  ' Comments  : Menu options.
  ' Parameters: None.
  ' Returns   : Nothing.
  ' Author    : Steve Grant
  '------------------------------------------------------------------------------
  Const cstrProc As String = "mnuNavDeleteRecord_Click"
  'On Error GoTo PROC_ERR
  
  cmdNav_Click (7)

PROC_EXIT:
  Exit Sub
  
PROC_ERR:
  Call MsgErr(mcstrMod & "." & cstrProc, Err.Number, Err.Description)
  Resume PROC_EXIT
End Sub


Private Sub mnuNavEditRecord_Click()
  ' Comments  : Menu options.
  ' Parameters: None.
  ' Returns   : Nothing.
  ' Author    : Steve Grant
  '------------------------------------------------------------------------------
  Const cstrProc As String = "mnuNavEditRecord_Click"
  'On Error GoTo PROC_ERR
  
  cmdNav_Click (9)

PROC_EXIT:
  Exit Sub
  
PROC_ERR:
  Call MsgErr(mcstrMod & "." & cstrProc, Err.Number, Err.Description)
  Resume PROC_EXIT
End Sub


Private Sub mnuNavFirstRecord_Click()
  ' Comments  : Menu options.
  ' Parameters: None.
  ' Returns   : Nothing.
  ' Author    : Steve Grant
  '------------------------------------------------------------------------------
  Const cstrProc As String = "mnuNavFirstRecord_Click"
  'On Error GoTo PROC_ERR
  
  cmdNav_Click (0)

PROC_EXIT:
  Exit Sub
  
PROC_ERR:
  Call MsgErr(mcstrMod & "." & cstrProc, Err.Number, Err.Description)
  Resume PROC_EXIT
End Sub


Private Sub mnuNavLastRecord_Click()
  ' Comments  : Menu options.
  ' Parameters: None.
  ' Returns   : Nothing.
  ' Author    : Steve Grant
  '------------------------------------------------------------------------------
  Const cstrProc As String = "mnuNavLastRecord_Click"
  'On Error GoTo PROC_ERR
  
  cmdNav_Click (4)

PROC_EXIT:
  Exit Sub
  
PROC_ERR:
  Call MsgErr(mcstrMod & "." & cstrProc, Err.Number, Err.Description)
  Resume PROC_EXIT
End Sub


Private Sub mnuNavNewRecord_Click()
  ' Comments  : Menu options.
  ' Parameters: None.
  ' Returns   : Nothing.
  ' Author    : Steve Grant
  '------------------------------------------------------------------------------
  Const cstrProc As String = "mnuNavNewRecord_Click"
  'On Error GoTo PROC_ERR
  
  cmdNav_Click (2)

PROC_EXIT:
  Exit Sub
  
PROC_ERR:
  Call MsgErr(mcstrMod & "." & cstrProc, Err.Number, Err.Description)
  Resume PROC_EXIT
End Sub


Private Sub mnuNavNextRecord_Click()
  ' Comments  : Menu options.
  ' Parameters: None.
  ' Returns   : Nothing.
  ' Author    : Steve Grant
  '------------------------------------------------------------------------------
  Const cstrProc As String = "mnuNavNextRecord_Click"
  'On Error GoTo PROC_ERR
  
  cmdNav_Click (3)

PROC_EXIT:
  Exit Sub
  
PROC_ERR:
  Call MsgErr(mcstrMod & "." & cstrProc, Err.Number, Err.Description)
  Resume PROC_EXIT
End Sub


Private Sub mnuNavPreviousRecord_Click()
  ' Comments  : Menu options.
  ' Parameters: None.
  ' Returns   : Nothing.
  ' Author    : Steve Grant
  '------------------------------------------------------------------------------
  Const cstrProc As String = "mnuNavPreviousRecord_Click"
  'On Error GoTo PROC_ERR
  
  cmdNav_Click (1)

PROC_EXIT:
  Exit Sub
  
PROC_ERR:
  Call MsgErr(mcstrMod & "." & cstrProc, Err.Number, Err.Description)
  Resume PROC_EXIT
End Sub


Private Sub mnuNavRequery_Click()
  ' Comments  : Menu options.
  ' Parameters: None.
  ' Returns   : Nothing.
  ' Author    : Steve Grant
  '------------------------------------------------------------------------------
  Const cstrProc As String = "mnuNavRequery_Click"
  'On Error GoTo PROC_ERR
  
  cmdNav_Click (8)

PROC_EXIT:
  Exit Sub
  
PROC_ERR:
  Call MsgErr(mcstrMod & "." & cstrProc, Err.Number, Err.Description)
  Resume PROC_EXIT

End Sub

Private Sub mnuNavSaveRecord_Click()
  ' Comments  : Menu options.
  ' Parameters: None.
  ' Returns   : Nothing.
  ' Author    : Steve Grant
  '------------------------------------------------------------------------------
  Const cstrProc As String = "mnuNavSaveRecord_Click"
  'On Error GoTo PROC_ERR
  
  cmdNav_Click (5)

PROC_EXIT:
  Exit Sub
  
PROC_ERR:
  Call MsgErr(mcstrMod & "." & cstrProc, Err.Number, Err.Description)
  Resume PROC_EXIT
End Sub


Private Sub mnuNavUndoChange_Click()
  ' Comments  : Menu options.
  ' Parameters: None.
  ' Returns   : Nothing.
  ' Author    : Steve Grant
  '------------------------------------------------------------------------------
  Const cstrProc As String = "mnuNavUndoChange_Click"
  'On Error GoTo PROC_ERR
  
  cmdNav_Click (6)

PROC_EXIT:
  Exit Sub
  
PROC_ERR:
  Call MsgErr(mcstrMod & "." & cstrProc, Err.Number, Err.Description)
  Resume PROC_EXIT
End Sub


Private Sub txtData_MouseMove(Index As Integer, Button As Integer, _
                              Shift As Integer, X As Single, Y As Single)
  ' Comments  : When the data moves on the control, display the info in
  '             txtNavInfo.
  ' Parameters: None.
  ' Returns   : Nothing.
  ' Author    : Steve Grant
  '------------------------------------------------------------------------------
  Const cstrProc As String = "txtData_MouseMove"
  'On Error GoTo PROC_ERR
  Dim strDisplayTxt As String
  
  strDisplayTxt = GetTag(txtData(Index).Tag, ";", NavDesc)
  
  If strDisplayTxt <> "NoValue" Then
    txtNavDesc.Text = strDisplayTxt
  Else ' Display the name of the field
    txtNavDesc.Text = GetTag(txtData(Index).Tag, ";", FieldUserName)
  End If

PROC_EXIT:
  Exit Sub
  
PROC_ERR:
  Call MsgErr(mcstrMod & "." & cstrProc, Err.Number, Err.Description)
  Resume PROC_EXIT
End Sub
