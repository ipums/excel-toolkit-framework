Attribute VB_Name = "menu_definition"
Option Explicit

' This file is part of the Minnesota Population Center's VBA libraries project.
' For copyright and licensing information, see the NOTICE and LICENSE files
' in this project's top-level directory, and also on-line at:
'   https://github.com/mnpopcenter/vba-libs


Const MENU_DEFINITION = _
    "Foo | FooMacro"                                               & vbLf & _
    "Bar | BarMacro"                                               & vbLf & _
    "---------"                                                    & vbLf & _
    "Compression ==>"                                              & vbLf & _
    "     Normal      | CompressData ""Normal"""                   & vbLf & _
    "     Fast        | CompressData ""Fast"""                     & vbLf & _
    "     Best        | CompressData ""Best"""                     & vbLf & _
    ""                                                             & vbLf & _
    "-------"                                                      & vbLf & _
    "Version  |  DisplayVersion"                                   & vbLf & _
    ""                                                             & vbLf & _
    ""                                                             & vbLf & _
    "#    (enabled only in Development mode)"                      & vbLf & _
    "#dev>----------------------------------"                      & vbLf & _
    "#dev>Developer Tools  ==>"                                    & vbLf & _
    "#dev>    Export VBA code           |  ExportVbaCode"          & vbLf & _
    "#dev>    ------------------------"                            & vbLf & _
    "#dev>    Build Production version  |  BuildProductionVersion"

Public Function LoadIntoArray(ByRef definition() As String) As Boolean
    definition = Split(MENU_DEFINITION, vbLf)
    LoadIntoArray = True
End Function
