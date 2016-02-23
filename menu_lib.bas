Attribute VB_Name = "menu_library"
Option Explicit

' This file is part of the Minnesota Population Center's VBA libraries project.
' For copyright and licensing information, see the NOTICE and LICENSE files
' in this project's top-level directory, and also on-line at:
'   https://github.com/mnpopcenter/vba-libs

' Add a custom menu to the worksheet menu bar.  The definition of the menu
' is a sequence of strings; for example:
'
'     # Menu definition
'     Foo | FooMacro
'     Bar | BarMacro
'     ---------
'     Compression ==>
'         Normal      | CompressData "Normal"
'         Fast        | CompressData "Fast"
'         Best        | CompressData "Best"
'
'     -------
'     Version  |  DisplayVersion
'
' Blank lines are ignored.  A comment line has "#" as the first non-whitespace
' character.  Comment lines are also ignored.  A submenu of the custom menu is
' is denoted with a "==>" at the end of line.
'
' A menu item (for the cusotm menu or one of its submenus) has the format:
'
'     menu item caption | action
'
' The action is the value assigned to the menu item's OnAction property.  It
' is the name of the macro to execute, along with any necessary arguments.
' Menu items for the custom menu are not indented.  The items for submenus
' must be indented at least 4 spaces.
'
' A separator in a menu (custom or submenu) is represented by a line with at
' least 4 "-" (hyphens).  Extra hyphens can be used for readability.  A
' submenu separator must be indented at least 4 spaces.
'
Sub AddCustomMenu(menuName As String, definition() As String, _
                  Optional insertBefore As String = "Help")
    Dim helpMenuIndex As Integer
    Dim customMenu As CommandBarControl
    Dim mainMenuBar As CommandBar
  
    Set mainMenuBar = Application.CommandBars("Worksheet Menu Bar")
    helpMenuIndex = mainMenuBar.Controls(insertBefore).Index
    Set customMenu = mainMenuBar.Controls.Add(Type:=msoControlPopup, _
                                              Before:=helpMenuIndex)
    customMenu.Caption = menuName

    Dim line As String
    Dim i As Long
    Dim currentSubMenu As CommandBarControl
    Dim addSeparator As Boolean
    addSeparator = False
    
    For i = LBound(definition) To UBound(definition)
        line = definition(i)
        If IsBlank(line) Or IsComment(line) Then
            ' Ignore blank lines and comment lines
        ElseIf Right(line, 3) = "==>" Then
            ' Submenu ends with "==>"
            With customMenu.Controls
                Set currentSubMenu = .Add(Type:=msoControlPopup)
            End With
            With currentSubMenu
                .Caption = Trim(Replace(line, "==>", ""))
                .BeginGroup = addSeparator
            End With
            addSeparator = False
        ElseIf Left(LTrim(line), 4) = "----" Then
            ' Add separator above the next menu item in the definition
            addSeparator = True
        Else
            ' New menu item for either current submenu or the custom menu
            Dim isSubmenuItem As Boolean
            isSubmenuItem = StartsWith(line, "    ")
            Dim menu As CommandBarControl
            Set menu = IIf(isSubmenuItem, currentSubMenu, customMenu)
            Dim menuItem As CommandBarControl
            Set menuItem = menu.Controls.Add(Type:=msoControlButton)

            ' line format =  menu item caption | action
            Dim fields() As String
            fields = Split(Trim(line), "|")
            Dim itemCaption As String
            itemCaption = Trim(fields(0))
            Dim itemAction As String
            itemAction = Trim(fields(1))
            With menuItem
                .Caption = itemCaption
                .OnAction = "'" & itemAction & "'"
                .BeginGroup = addSeparator
            End With
            addSeparator = False
        End If
    Next
End Sub

Function StartsWith(str_ As String, prefix As String) As Boolean
    StartsWith = Left(str_, Len(prefix)) = prefix
End Function

Function IsBlank(line As String) As Boolean
   IsBlank = RTrim(line) = ""
End Function

Function IsComment(line As String) As Boolean
   IsComment = Left(LTrim(line), 1) = "#"
End Function

Public Sub RemoveCustomMenu(menuName As String)
    With Application.CommandBars("Worksheet Menu Bar")
       Dim ctrl As CommandBarControl
       For Each ctrl In .Controls
           If ctrl.Caption = menuName Then
               ctrl.Delete
           End If
       Next ctrl
    End With
End Sub
