Attribute VB_Name = "modInformation"
Option Explicit

Dim s As String

Function InitialiseRTF(sFile As String)
    
    Open sFile For Binary Access Read As #1
        s = modFastString.AllocString_ADVANCED(LOF(1))
        Get #1, , s
    Close #1
    
End Function

Function SetVar(sVar As String, sValue As String)
s = Replace(s, sVar, sValue)
End Function

Function Update()
On Error Resume Next
Kill App.Path & "\x.htm"
Open App.Path & "\x.htm" For Binary Access Write As #1
    Put #1, , s
Close #1

Form1.wb1.Navigate App.Path & "\x.htm"

End Function
