Attribute VB_Name = "initialization"
Option Explicit

' This file is part of the Minnesota Population Center's VBA libraries project.
' For copyright and licensing information, see the NOTICE and LICENSE files
' in this project's top-level directory, and also on-line at:
'   https://github.com/mnpopcenter/vba-libs


' Module dependencies:
'   bootstrap
'   conf

' The list of full paths to all the modules in the add-in.  They are in the
' order that they are imported into the development version of the add-in.
' Each path is also indexed by the module's name (i.e., what's specified by
' its Attribute VB_Name).
Public ModulePaths As Collection

Public Sub InitializeDevelopmentMode()
    InitializeModulePaths
    LoadModules
    Application.Run "toolkit.Initialize"
End Sub

Public Sub InitializeProductionMode()
    Application.Run "toolkit.Initialize"
End Sub

' Initialize ModulePaths with the 3 modules that are initially in the add-in
Private Sub InitializeModulePaths()
    Set ModulePaths = New Collection
    Const STANDARD_MODULE_TYPE = 1
    Dim component
    For Each component In ThisWorkbook.VBProject.VBComponents
        If component.Type = STANDARD_MODULE_TYPE Then
            Dim module_path As String
            If component.Name = bootstrap.ConfModule_Name Then
                module_path = bootstrap.ConfModule_Path
            ElseIf component.Name = bootstrap.InitModule_Name Then
                module_path = bootstrap.InitModule_Path
            Else
                ' Only other module is the bootstrap one
                module_path = PathInThisWorkbookDir(bootstrap.MODULE_FILENAME)
            End If
            ModulePaths.Add module_path, component.Name
        End If
    Next component
End Sub

' Import all the other modules in the toolkit
Private Sub LoadModules()
    Dim module_file_names() As String
    module_file_names = Split(conf.MODULE_FILENAMES, "|")
    With ThisWorkbook.VBProject.VBComponents
        Dim i As Integer
        For i = LBound(module_file_names) To UBound(module_file_names)
            Dim module_path As String
            module_path = PathInThisWorkbookDir(module_file_names(i))
            Dim imported_module
            Set imported_module = .Import(module_path)
            ModulePaths.Add module_path, imported_module.Name
        Next i
    End With
End Sub

Public Function PathInThisWorkbookDir(file_name As String) As String
    PathInThisWorkbookDir = ThisWorkbook.Path & Application.PathSeparator _
                                              & file_name
End Function
