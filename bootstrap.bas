Attribute VB_Name = "bootstrap"
Option Explicit

' This file is part of the Minnesota Population Center's VBA libraries project.
' For copyright and licensing information, see the NOTICE and LICENSE files
' in this project's top-level directory, and also on-line at:
'   https://github.com/mnpopcenter/vba-libs

Public Const MODULE_FILENAME = "bootstrap.bas"

' The module that's used to update the toolkit's core modules
Private Const MODULE_FOR_CORE_UPDATES = "update_core"

Public Enum ToolkitMode
    Unknown = 0   ' So an uninitialized variable will have this value
    Development
    Production
End Enum

' The mode that the add-in is currently running in.
Public CurrentMode As ToolkitMode

Public Enum ToolkitEdition
    Unknown = 0          ' So an uninitialized variable will have this value
    Development          ' (toolkit base name)_DEV.xlam
    BuiltProduction      ' (toolkit base name)_PROD.xlam
    InstalledProduction  ' (toolkit base name).xlam
End Enum

' The current edition that the add-in represents
Public CurrentEdition As ToolkitEdition

' The name and file path for the configuration and loader modules that
' are imported in Development mode.
Public ConfModule_Name As String
Public ConfModule_Path As String
Public LoaderModule_Name As String
Public LoaderModule_Path As String

' Can this toolkit edition be saved?
' By default, no.  Only the development edition can be saved when building
' the production edition.
Public AllowToolkitSave As Boolean

' This module is NOT imported into the development version of the add-in.
' It is exported to MODULE_FILENAME so a copy of its code is under version
' control.  Changes to its code must be made in the Visual Basic editor.  To
' make those changes, the development add-in must be opened with
' macros disabled.  That will allow the add-in to be saved with just this
' bootstrap module.

' Called by ThisWorkbook.Workbook_Open event procedure (as a workaround
' for this issue: http://stackoverflow.com/q/34498794/1258514)
Public Sub InitializeAddIn()
    If ThisWorkbook.Name Like "*NO-LOAD*" Then
        AllowToolkitSave = True
        ' Do not import any modules so developer can change file properties.
        Exit Sub
    End If

    AllowToolkitSave = False
    If ThisWorkbook.Name Like "*DEV*" Then
        CurrentMode = ToolkitMode.Development
        CurrentEdition = ToolkitEdition.Development
        ConfModule_Path = Replace(ThisWorkbook.FullName, "DEV.xlam", _
                                                          "conf.bas")
        LoaderModule_Path = ThisWorkbook.Path & Application.PathSeparator _
                                              & "loader.bas"
        With ThisWorkbook.VBProject.VBComponents
            ConfModule_Name = .Import(ConfModule_Path).Name
            LoaderModule_Name = .Import(LoaderModule_Path).Name
        End With
        Application.Run "loader.LoadToolkitModules"
    Else
        CurrentMode = ToolkitMode.Production
        If ThisWorkbook.Name Like "*PROD*" Then
            CurrentEdition = ToolkitEdition.BuiltProduction
        Else
            CurrentEdition = ToolkitEdition.InstalledProduction
        End If
    End If
    Application.Run "toolkit.Initialize"
End Sub

' Determine whether to allow the current toolkit edition to be saved.
'
' Called from ThisWorkbook.Workbook_BeforeSave event handler.
Public Sub BeforeToolkitSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    ' Don't allow NO-LOAD edition to save if the module for updating the core
    ' modules hasn't been unloaded yet.
    If ThisWorkbook.Name Like "*_NO-LOAD.xlam" Then
        If IsModuleLoaded(MODULE_FOR_CORE_UPDATES) Then
            Cancel = True
            MsgBox "The toolkit cannot be saved with the " & _
                   MODULE_FOR_CORE_UPDATES & " module loaded.  Please" & _
                   " run the macro to unload it first.", vbExclamation
        End If
        ' Don't check the AllowToolkitSave module variable because when the
        ' bootstrap module's source code is changed, that variable is reset
        ' (i.e., which for Boolean variables means it's = False).
        Exit Sub
    End If
    If Not AllowToolkitSave Then
        Cancel = True
        If CurrentMode = ToolkitEdition.Development Then
            TellUser_SavingDisabled
        End If
    End If
End Sub

Private Sub TellUser_SavingDisabled()
    MsgBox "Saving the Development edition of this toolkit" & vbCr & _
           "(" & ThisWorkbook.Name & ") from the VB editor is not" & vbCr & _
           "allowed.  Instead, select this action in its menu:" & vbCr & _
           vbCr & _
           "    Developer Tools --> Export VBA code", _
           vbCritical
End Sub

' ---------------------------------------------------------------------------
' Subprocedures for updating the toolkit's core modules when needed

Private Sub LoadModuleForCoreUpdates()
    If Not ThisWorkbook.Name Like "*_NO-LOAD.xlam" Then
        MsgBox "Error: the LoadModuleForCoreUpdates macro should only be" _
               & " run in a toolkit's NO-LOAD edition.", vbCritical
        Exit Sub
    End If

    If IsModuleLoaded(MODULE_FOR_CORE_UPDATES) Then
        ReportModuleStatus "is already loaded"
    Else
        Dim module_path As String
        module_path = ThisWorkbook.Path & Application.PathSeparator & _
                      MODULE_FOR_CORE_UPDATES & ".bas"
        ThisWorkbook.VBProject.VBComponents.Import (module_path)
        ReportModuleStatus "loaded"
    End If
End Sub

Private Sub UnloadModuleForCoreUpdates()
    If IsModuleLoaded(MODULE_FOR_CORE_UPDATES) Then
        With ThisWorkbook.VBProject
            .VBComponents.Remove .VBComponents(MODULE_FOR_CORE_UPDATES)
        End With
        ReportModuleStatus "unloaded"
    Else
        ReportModuleStatus "is not loaded"
    End If
End Sub

Private Sub ReportModuleStatus(mod_status As String)
    Debug.Print ThisWorkbook.Name & " -- " & MODULE_FOR_CORE_UPDATES & _
                " module " & mod_status
End Sub

Private Function IsModuleLoaded(module_name) As Boolean
    On Error GoTo NoSuchModule
    With ThisWorkbook.VBProject.VBComponents(module_name)
        ' Do nothing
    End With
    IsModuleLoaded = True
    Exit Function
NoSuchModule:
    IsModuleLoaded = False
End Function
