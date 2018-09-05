Attribute VB_Name = "update_core"
Option Explicit

' This file is part of the Minnesota Population Center's VBA libraries project.
' For copyright and licensing information, see the NOTICE and LICENSE files
' in this project's top-level directory, and also on-line at:
'   https://github.com/mnpopcenter/vba-libs

' Update the core modules in the current toolkit by loading revised source
' code from their corresponding code files:
'
'   ThisWorkbook.cls
'   bootstrap.bas
'
Public Sub UpdateCoreModules()
    UpdateCoreModuleFrom "ThisWorkbook.cls", vbCrLf & "Private Sub"
    UpdateCoreModuleFrom "bootstrap.bas", "Option Explicit"

    Debug.Print ThisWorkbook.Name & " -- after reviewing the changes, " & _
                                        "save the file"
End Sub

' Update the source code of a core module from its corresponding code file.
Private Sub UpdateCoreModuleFrom(source_file_name As String, _
                                 initial_code_text As String)
    Dim source_file_path As String
    source_file_path = ThisWorkbook.Path & Application.PathSeparator & _
                       source_file_name

    Dim source_code As String
    Dim file_number As Integer
    file_number = FreeFile()
    Open source_file_path For Input As #file_number
    source_code = Input$(LOF(file_number), file_number)
    Close file_number

    ' Remove the leading non-code line(s) from the source code
    Dim code_len As Long
    code_len = Len(source_code) - InStr(source_code, initial_code_text) + 1
    source_code = Right(source_code, code_len)

    ' Trim the last line terminator so when the module is exported, the
    ' exported code will match the contents in the file.
    Dim line_terminator_len As Long
    If Right(source_code, 2) = vbCrLf Then
        line_terminator_len = 2
    Else
        line_terminator_len = 1
    End If
    source_code = Left(source_code, Len(source_code) - line_terminator_len)

    Dim module_name As String
    module_name = Split(source_file_name, ".")(0)

    With ThisWorkbook.VBProject.VBComponents(module_name).CodeModule
        .DeleteLines 1, .CountOfLines
        .InsertLines 1, source_code

        ' For an unknown reason (hey, it's VBA!), an extra blank line is
        ' appended onto the bootstrap module.  Delete this last code line so
        ' that the module will round-trip correctly (i.e., exported code will
        ' match the file contents that were imported).
        If module_name = "bootstrap" Then
            If .Lines(.CountOfLines, 1) = "" Then
                .DeleteLines .CountOfLines, 1
            End If
        End If
    End With

    Debug.Print ThisWorkbook.Name & " -- core module updated: " & module_name
End Sub
