Attribute VB_Name = "file_utilities"
Option Explicit

' This file is part of the Minnesota Population Center's VBA libraries project.
' For copyright and licensing information, see the NOTICE and LICENSE files
' in this project's top-level directory, and also on-line at:
'   https://github.com/mnpopcenter/vba-libs


Public Function FileExists(file_path As String) As Boolean
    If Application.OperatingSystem Like "*Mac*" Then
        On Error GoTo DirErr
    End If

    FileExists = Dir(file_path) <> ""
    Exit Function

DirErr:
    ' Excel for Mac raises an error if Dir() function called with a path that
    ' does not exist.  For details, see:
    '   http://answers.microsoft.com/en-us/mac/forum/macoffice2011-macexcel/lets-talk-about-mac-excel-2011-bugs/a3653864-e889-4413-aab0-ac118c03d65e
    If Err.Number = 68 Then
        FileExists = False
    Else
        MsgBox Err.Description & " (" & Err.Number & ")", , "Run-time Error"
        Stop
    End If
End Function
