Attribute VB_Name = "developer_tools"
Option Explicit

' This file is part of the Minnesota Population Center's VBA libraries project.
' For copyright and licensing information, see the NOTICE and LICENSE files
' in this project's top-level directory, and also on-line at:
'   https://github.com/mnpopcenter/vba-libs


' Module dependencies:
'   conf
'   loader
'   toolkit


' Full path to the production version of this add-in
Private ProductionAddInPath As String

' Names of macros to run before and after building the production version
' (default = "", which is a no-op).
Public PreBuildMacro As String    ' Sub()
Public PostBuildMacro As String   ' Function(build_message As String) As String
                                  '    returns an updated build message or ""
                                  '    if an error occurred and was displayed
                                  '    to the developer.

Public Sub ExportVbaCode()
    Dim message As String
    message = "Exported the VBA code into these files:" & vbCr

    Dim component
    For Each component In ThisWorkbook.VBProject.VBComponents
        Dim module_path As String
        If component.Name = "ThisWorkbook" Or component.Name = "Sheet1" Then
            module_path = PathInThisWorkbookDir(component.Name & ".cls")
        Else
            module_path = loader.ModulePaths(component.Name)
        End If
        component.Export module_path
        Dim file_name As String
        file_name = Replace(module_path, PathInThisWorkbookDir(""), "")
        message = message & vbCr & _
                  "    " & component.Name & "  -->  " & file_name
    Next component

    MsgBox message & vbCr & _
           vbCr & _
           "The files are located in this directory:" & vbCr & _
           vbCr & _
           ThisWorkbook.Path, vbOKOnly, "Exported VBA Code"
End Sub

Public Sub BuildProductionVersion()
    Dim prod_add_in_name As String
    prod_add_in_name = Replace(ThisWorkbook.Name, "_DEV", "_PROD")
    ProductionAddInPath = PathInThisWorkbookDir(prod_add_in_name)

    Dim message As String
    Dim window_title As String
    If FileExists(ProductionAddInPath) Then
        message = "The add-in """ & prod_add_in_name & """ already exists in" & _
                  " the folder:" & _
                  vbCr & vbCr & _
                  ThisWorkbook.Path & _
                  vbCr & vbCr & _
                  "Do you want to rebuild it?"
        Dim Answer As Integer
        Answer = MsgBox(message, vbYesNo, "Rebuild the Add-in?")
        If Answer = vbNo Then
            Exit Sub
        End If
        Kill ProductionAddInPath
        window_title = "Rebuilt the Production Add-in"
    Else
        window_title = "Built the Production Add-in"
    End If

    If PreBuildMacro <> "" Then Application.Run PreBuildMacro
    MakeProductionAddIn
    message = "Created the add-in """ & prod_add_in_name & """ in the folder:" & _
              vbCr & vbCr & _
              ThisWorkbook.Path
    If PostBuildMacro <> "" Then
        message = Application.Run(PostBuildMacro, message)
        If message = "" Then
            ' Error occurred and the developer was notified, so just exit
            Exit Sub
        End If
    End If
    MsgBox message, vbOKOnly, window_title
End Sub

Private Sub MakeProductionAddIn()
    toolkit.StoreBuildDateTime Now()

    ' Set the add-in's Title and Comments (description) for the production
    ' version
    Dim titleDev As String
    titleDev = ThisWorkbook.Title
    Dim commentsDev As String
    commentsDev = ThisWorkbook.Comments
    With ThisWorkbook
        .Title = Replace(titleDev, " (dev)", "")
        .Comments = Replace(commentsDev, "development version", _
                                         "version " & conf.VERSION_STR)
        .SaveAs ProductionAddInPath, xlOpenXMLAddIn
    End With

    ' Restore the development version of Title and Comments
    With ThisWorkbook
        .Title = titleDev
        .Comments = commentsDev
    End With
End Sub
