Attribute VB_Name = "toolkit"
Option Explicit

' This file is part of the Minnesota Population Center's VBA libraries project.
' For copyright and licensing information, see the NOTICE and LICENSE files
' in this project's top-level directory, and also on-line at:
'   https://github.com/mnpopcenter/vba-libs


' Module dependencies:
'   conf
'   initialization   (ToolkitMode)
'   menu_definition
'   menu

' The mode that the toolkit is currently running in
Public CurrentMode As ToolkitMode

' Date when the production version was built
Public Const BUILD_DATE_FORMAT = "mmmm d, yyyy"
Public BuildDate As String

' Cell where build date & time of production version is stored.
' Even add-ins have at least 1 worksheet
Public Const BUILT_WHEN_CELL = "$A$1"

Public Sub Initialize(mode As ToolkitMode)
    CurrentMode = mode
    If mode = ToolkitMode.Development Then
        toolkit.BuildDate = "Development version"
    Else
        Dim builtWhen As Date
        builtWhen = ThisWorkbook.Worksheets(1).Range(BUILT_WHEN_CELL).Value
        toolkit.BuildDate = Format(builtWhen, BUILD_DATE_FORMAT)
    End If

    If conf.ADDITIONAL_INIT_MACRO <> "" Then
        Application.Run conf.ADDITIONAL_INIT_MACRO
    End If

    Dim menu_defn() As String
    If menu_definition.LoadIntoArray(menu_defn) Then
        CreateToolkitMenu mode, menu_defn
    End If
End Sub

Public Sub Auto_Close()
    RemoveToolkitMenu
End Sub

Public Sub StoreBuildDateTime(date_time As Date)
    ThisWorkbook.Worksheets(1).Range(BUILT_WHEN_CELL).Value = date_time
End Sub

Public Sub DisplayVersion()
    MsgBox BuildDate & " (" & conf.VERSION_STR & ").", vbOKOnly, _
           conf.TOOLKIT_NAME
End Sub
