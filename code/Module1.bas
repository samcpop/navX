Attribute VB_Name = "Module1"
Public IsInternalCall As Boolean
Sub showHeadings()
    UserForm1.show vbModeless
End Sub

Sub z_help()
    If IsInternalCall Then
        On Error Resume Next
        UserForm1.Controls("fraFocusHolder").SetFocus
        DoEvents
        UserForm1.NavListBox.ListIndex = -1
        On Error GoTo 0
        
        IsInternalCall = False
        Exit Sub
    End If
    
    MsgBox "github.com/samcpop/navX"
End Sub
