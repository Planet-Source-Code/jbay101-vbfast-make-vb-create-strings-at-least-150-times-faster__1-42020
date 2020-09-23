VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FastString benchmark"
   ClientHeight    =   7425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9555
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7425
   ScaleWidth      =   9555
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   675
      Left            =   3045
      ScaleHeight     =   615
      ScaleWidth      =   2850
      TabIndex        =   1
      Top             =   2220
      Width           =   2910
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Running benchmark..."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   165
         TabIndex        =   2
         Top             =   195
         Width           =   2595
      End
   End
   Begin SHDocVwCtl.WebBrowser wb1 
      Height          =   4275
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   8820
      ExtentX         =   15557
      ExtentY         =   7541
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Show
Me.Refresh
Picture1.Refresh
StringBenchmark 30000 '20 MB
Picture1.Visible = False
wb1.Visible = True
End Sub

Function StringBenchmark(lKB As Long)
Dim sResult As String
Dim lSize As Long
Dim vTimer As New clsTimer
Dim sOutput As String

InitialiseRTF App.Path & "\format.htm"

SetVar "$name$", "FastString"

SetVar "$introduction$", "Visual Basic stores it's strings in a type refered to in C++ as a BSTR." & vbCrLf & _
                         "This is an ActiveX-type format which is defined in the OLE Automation library." & vbCrLf & _
                         "In Visual Basic, when you dynamicly create a string, VB automaticly fills it with data." & vbCrLf & _
                         "However, when you deal with large strings, this can take a lot of time, and most of the time" & vbCrLf & _
                         "you don't want this. So, included in this project is a function which creates the string and doesn't fill it" & vbCrLf & _
                         "with data. This is a lot faster, as this benchmark shows."


lSize = 1024 * lKB
SetVar "$buffer$", Format(lKB, "###,###,###,###,###")

SetVar "$function1$", "AllocString_Space"
SetVar "$description1$", "Allocate storage using the Visual Basic 'Space' function."
vTimer.ResetTimer
    sResult = modFastString.AllocString_Space(lSize)
vTimer.StopTimer
 
SetVar "$time1$", Round(vTimer.Elapsed, 2)

SetVar "$function2$", "AllocString_StringASC"
SetVar "$description2$", "Allocate storage using the Visual Basic 'String' function and passing it a character code."
vTimer.ResetTimer
sResult = modFastString.AllocString_StringASC(lSize)
vTimer.StopTimer
SetVar "$time2$", Round(vTimer.Elapsed, 2)

SetVar "$function3$", "AllocString_StringCHR"
SetVar "$description3$", "Allocate storage using the Visual Basic 'String' function and passing it a character."
vTimer.ResetTimer
sResult = modFastString.AllocString_StringCHR(lSize)
vTimer.StopTimer
SetVar "$time3$", Round(vTimer.Elapsed, 2)


SetVar "$function4$", "AllocString_ADVANCED"
SetVar "$description4$", "Using the API to allocate the storage."
vTimer.ResetTimer
sResult = modFastString.AllocString_ADVANCED(lSize)
vTimer.StopTimer
SetVar "$time4$", Round(vTimer.Elapsed, 2)
sResult = ""
Update
End Function

Private Sub Form_Resize()
wb1.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight

End Sub
