Attribute VB_Name = "tools"
Option Explicit

' Module dependencies:
'    conf

Public Sub FooMacro()
    ShowRunningToolMessage "FooMacro"
End Sub

Public Sub BarMacro()
    ShowRunningToolMessage "BarMacro"
End Sub

Public Sub CompressData(level As String)
    ShowRunningToolMessage "CompressData", "level = " & level
End Sub

Private Sub ShowRunningToolMessage(tool_name As String, _
                                   Optional extra_info As String = "")
    Dim message As String
    message = "Running the " & tool_name & " tool..."
    If extra_info <> "" Then
        message = message & vbCr & _
                  extra_info
    End If
    MsgBox message, vbOKOnly, conf.TOOLKIT_NAME
End Sub
