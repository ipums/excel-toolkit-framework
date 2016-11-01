Attribute VB_Name = "conf"
Option Explicit

' This file is part of the Minnesota Population Center's VBA libraries project.
' For copyright and licensing information, see the NOTICE and LICENSE files
' in this project's top-level directory, and also on-line at:
'   https://github.com/mnpopcenter/vba-libs


' Used as the title of message boxes
Public Const TOOLKIT_NAME = "Simple Tools"

Public Const VERSION_STR = "1.2.0"

Public Const MENU_NAME = "&Simple Tools"

' The name of the macro that performs additional initialization for the
' toolkit.  If not blank, it is run before the toolkit's menu is created.
Public Const ADDITIONAL_INIT_MACRO = ""

' The order of this list determines the order that the modules are imported.
' Therefore, if module FOO depends on modules BAR and QUX, then make sure
' FOO is after BAR and QUX in the list.
Public Const MODULE_FILENAMES = _
       "excel_ver.bas" _
    & "|file_utils.bas" _
    & "|menu_lib.bas" _
    & "|menu_defn_in_code.bas" _
    & "|menu.bas" _
    & "|toolkit.bas" _
    & "|tools.bas" _
    & "|dev_tools.bas"
