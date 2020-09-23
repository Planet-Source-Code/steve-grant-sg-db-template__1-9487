VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStopWatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Module     : StopWatch
' Description: From "VBA Developer's Handbook".
'              Class that calculates the time it took to process a piece of code.
'
' Example    : A simple stopwatch class
'              Dim sw As StopWatch
'
'              Set sw = New StopWatch
'              sw.StartTimer
'              Piece of code
'              Debug.Print sw.EndTimer
' Procedures : StartTimer
'              EndTimer
' Author     : Steve Grant
'--------------------------------------------------------------------------------
Option Explicit

Private plngStart As Long
Private Declare Function timeGetTime Lib "winmm.dll" () As Long


Public Sub StartTimer()
  plngStart = timeGetTime
End Sub


Public Function EndTimer() As Long
  EndTimer = (timeGetTime - plngStart)
End Function

