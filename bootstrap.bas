Attribute VB_Name = "bootstrap"
Option Explicit

' This file is part of the Minnesota Population Center's VBA libraries project.
' For copyright and licensing information, see the NOTICE and LICENSE files
' in this project's top-level directory, and also on-line at:
'   https://github.com/mnpopcenter/vba-libs

Public Const MODULE_FILENAME = "bootstrap.bas"

' This module is NOT imported into the development version of the add-in.
' It is exported to MODULE_FILENAME so a copy of its code is under version
' control.  Changes to its code must be made in the Visual Basic editor.  To
' make those changes, the development add-in must be opened with
' macros disabled.  That will allow the add-in to be saved with just this
' bootstrap module.

' Called by ThisWorkbook.Workbook_Open event procedure (as a workaround
' for this issue: http://stackoverflow.com/q/34498794/1258514)
Public Sub InitializeAddIn()
    If ThisWorkbook.Name Like "*DEV*" Then
        Dim conf_module_path As String
        conf_module_path = Replace(ThisWorkbook.FullName, "DEV.xlam", _
                                                          "conf.bas")
        Dim init_module_path As String
        init_module_path = ThisWorkbook.Path & Application.PathSeparator _
                                             & "initialization.bas"
        Dim conf_module
        Dim init_module
        With ThisWorkbook.VBProject.VBComponents
            Set conf_module = .Import(conf_module_path)
            Set init_module = .Import(init_module_path)
        End With
        Application.Run "InitializeDevelopmentMode", conf_module.Name, _
                                                     conf_module_path, _
                                                     init_module.Name, _
                                                     init_module_path
    Else
        Application.Run "InitializeProductionMode"
    End If
End Sub
