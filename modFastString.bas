Attribute VB_Name = "modFastString"
Option Explicit
Private Declare Sub RtlMoveMemory Lib "kernel32" (dst As Any, src As Any, ByVal nBytes&)
Private Declare Function SysAllocStringByteLen& Lib "oleaut32" (ByVal olestr&, ByVal BLen&)

' method 1 - the Space function
Public Function AllocString_Space(ByVal lSize As Long) As String
  AllocString_Space = Space$(lSize)
End Function

' method 2 - the String function with a character code
Public Function AllocString_StringASC(ByVal lSize As Long) As String
  AllocString_StringASC = String$(lSize, 32)
End Function

' method 3 - the String function with a character
Public Function AllocString_StringCHR(ByVal lSize As Long) As String
  AllocString_StringCHR = String$(lSize, " ")
End Function

' method 4 - the advanced **FAST** method
Public Function AllocString_ADVANCED(ByVal lSize As Long) As String
  RtlMoveMemory ByVal VarPtr(AllocString_ADVANCED), _
    SysAllocStringByteLen(0&, lSize + lSize), 4&
End Function

