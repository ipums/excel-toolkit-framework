Attribute VB_Name = "menu_definition"
Option Explicit

' This file is part of the Minnesota Population Center's VBA libraries project.
' For copyright and licensing information, see the NOTICE and LICENSE files
' in this project's top-level directory, and also on-line at:
'   https://github.com/mnpopcenter/vba-libs


Const MENU_DEFINITION_STR = _
             "Foo | FooMacro" _
    & vbLf & "Bar | BarMacro" _
    & vbLf & "---------" _
    & vbLf & "Compression ==>" _
    & vbLf & "     Normal      | CompressData ""Normal""" _
    & vbLf & "     Fast        | CompressData ""Fast""" _
    & vbLf & "     Best        | CompressData ""Best""" _
    & vbLf & "" _
    & vbLf & "-------" _
    & vbLf & "Version  |  DisplayVersion" _
    & vbLf & "" _
    & vbLf & "" _
    & vbLf & "#    (enabled only in Development mode)" _
    & vbLf & "#dev>----------------------------------" _
    & vbLf & "#dev>Developer Tools  ==>" _
    & vbLf & "#dev>    Export VBA code           |  ExportVbaCode" _
    & vbLf & "#dev>    ------------------------" _
    & vbLf & "#dev>    Build Production version  |  BuildProductionVersion"

Public Function LoadIntoArray(ByRef definition() As String) As Boolean
    definition = Split(MENU_DEFINITION_STR, vbLf)
    LoadIntoArray = True
End Function
