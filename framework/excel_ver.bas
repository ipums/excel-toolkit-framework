Attribute VB_Name = "Excel_version"
Option Explicit

' This file is part of the Minnesota Population Center's VBA libraries project.
' For copyright and licensing information, see the NOTICE and LICENSE files
' in this project's top-level directory, and also on-line at:
'   https://github.com/mnpopcenter/vba-libs


' Constants for values returned by Application.Version

Public Const EXCEL_WIN_2003 = 11
Public Const EXCEL_WIN_2007 = 12
Public Const EXCEL_WIN_2010 = 14
Public Const EXCEL_WIN_2013 = 15
Public Const EXCEL_WIN_2016 = 16

Public Const EXCEL_MAC_2004 = 11
Public Const EXCEL_MAC_2008 = 12   ' Note: did not have support for VBA
Public Const EXCEL_MAC_2011 = 14
Public Const EXCEL_MAC_2016 = 15

' Enumerated type for platform-specific Excel
Public Enum ExcelPlatform
    ExcelUnknown
    ExcelWin
    ExcelMac
End Enum

' How to compare versions
Public Enum VersionComparison
    Exact      ' only the specified version matches
    OrLater    ' matches the specified version or a later (newer) one
    OrEarlier  ' matches the specified version or an earlier (older) one
End Enum

' Indicates that any year for a platform-specific Excel is acceptable
Private Const ANY_YEAR = -999

Public Function ExcelVersionIs( _
    expected_platform As ExcelPlatform, _
    Optional expected_year As Integer = ANY_YEAR, _
    Optional version_compare As VersionComparison = Exact) _
As Boolean
    Dim platform_app As ExcelPlatform
    #If Mac Then
        platform_app = ExcelMac
    #Else
        platform_app = ExcelWin
    #End If
    If platform_app <> expected_platform Then
        ExcelVersionIs = False
    ElseIf expected_year = ANY_YEAR Then
        ExcelVersionIs = True
    Else
        Dim major_version As Integer
        major_version = Int(Val(Application.Version))
        Dim expected_version As Integer
        expected_version = GetAppVersion(platform_app, expected_year)
        Select Case version_compare
            Case Exact
                ExcelVersionIs = (major_version = expected_version)
            Case OrLater
                ExcelVersionIs = (major_version >= expected_version)
            Case OrEarlier
                ExcelVersionIs = (major_version <= expected_version)
            Case Else
                ExcelVersionIs = False
        End Select
    End If
End Function

Public Function GetAppVersion(platform_app As ExcelPlatform, _
                              app_year As Integer) As Integer
    If platform_app = ExcelWin Then
        Select Case app_year
            Case 2003:  GetAppVersion = EXCEL_WIN_2003
            Case 2007:  GetAppVersion = EXCEL_WIN_2007
            Case 2010:  GetAppVersion = EXCEL_WIN_2010
            Case 2013:  GetAppVersion = EXCEL_WIN_2013
            Case 2016:  GetAppVersion = EXCEL_WIN_2016
            Case Else:  GetAppVersion = 0
        End Select
    ElseIf platform_app = ExcelMac Then
        Select Case app_year
            Case 2004:  GetAppVersion = EXCEL_MAC_2004
            Case 2008:  GetAppVersion = EXCEL_MAC_2008
            Case 2011:  GetAppVersion = EXCEL_MAC_2011
            Case 2016:  GetAppVersion = EXCEL_MAC_2016
            Case Else:  GetAppVersion = 0
        End Select
    Else
        GetAppVersion = 0
    End If
End Function
