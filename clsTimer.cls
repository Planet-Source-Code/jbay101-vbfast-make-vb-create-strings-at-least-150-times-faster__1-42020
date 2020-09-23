VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function QueryPerformanceCounter Lib "kernel32" ( _
    lpPerformanceCount As Currency) _
As Long

Private Declare Function QueryPerformanceFrequency Lib "kernel32" ( _
    lpFrequency As Currency) _
As Long

Private m_curFrequency As Currency
Private m_HasCounter As Boolean
Private m_StartTime As Currency
Private m_StopTime As Currency

Private tmp As Boolean


' Determine if the computer supports a high performance counter
Private Sub Class_Initialize()
    m_HasCounter = QueryPerformanceFrequency(m_curFrequency)
    m_curFrequency = m_curFrequency * 10000
End Sub


' Allow the programmers to verify that the high performance
' counter is supported
Public Property Get HasCounter() As Boolean
    HasCounter = m_HasCounter
End Property


' Returns the frequency of the high performance counter
Public Property Get TicksPerSecond() As Long
    TicksPerSecond = m_curFrequency
End Property

' Call immediately prior to execution of the code being timed
Public Sub ResetTimer()
    tmp = QueryPerformanceCounter(m_StartTime)
End Sub


' Call immediatley after execution of the code being timed
Public Sub StopTimer()
    tmp = QueryPerformanceCounter(m_StopTime)
End Sub


' Returns the number of microseconds that elapsed
Public Function Elapsed() As Double
    Elapsed = (m_StopTime - m_StartTime) / m_curFrequency * 10000 * 1000
End Function


' Returns the a nicely formatted string containing the
' number of microseconds that have elpased
Public Function strElapsed() As String
    strElapsed = "Execution took " & Format$((m_StopTime - m_StartTime) / m_curFrequency * 10000 * 1000, "###,##0.000") & " ms"
End Function
