Attribute VB_Name = "menu"
Option Explicit

' This file is part of the Minnesota Population Center's VBA libraries project.
' For copyright and licensing information, see the NOTICE and LICENSE files
' in this project's top-level directory, and also on-line at:
'   https://github.com/mnpopcenter/vba-libs


' Module dependencies:
'   bootstrap      (CurrentMode, ToolkitMode, CurrentEdition, ToolkitEdition)
'   conf
'   Excel_version
'   menu_library

Private myMenuName As String

Public Sub CreateToolkitMenu(ByRef definition() As String)
    myMenuName = conf.MENU_NAME
    If CurrentMode = ToolkitMode.Development Then
        myMenuName = myMenuName & " (dev)"
    ElseIf CurrentEdition = ToolkitEdition.BuiltProduction Then
        myMenuName = myMenuName & " (prod)"
    End If

    If CurrentMode = ToolkitMode.Development Then
        EnableDevelopersMenu definition
    End If

    ' Workaround for bug with menus not being removed when an add-in closes.
    ' https://support.microsoft.com/en-us/kb/2761240
    If ExcelVersionIs(ExcelWin, 2013) Then
        ' Remove any menus left over from the previous Excel session
        RemoveToolkitMenu
    End If
    AddCustomMenu myMenuName, definition
End Sub

Public Sub RemoveToolkitMenu()
    RemoveCustomMenu myMenuName
End Sub

Sub EnableDevelopersMenu(definition() As String)
    ' Remove any "#dev>" markers to enable the developers menu
    Dim i As Integer
    For i = LBound(definition) To UBound(definition)
        definition(i) = Replace(definition(i), "#dev>", "")
    Next i
End Sub
