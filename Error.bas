Attribute VB_Name = "Error"


Function View(ByVal What As String) As Boolean

TestModus = False

Select Case TestModus
Case True

    Select Case What
    'View error message
    Case "Errr"
    View = False
    
    'View screen update
    Case "Updte"
    View = True
    
    'View Alerts
    Case "Alrt"
    View = True
    
    Case ""
    View = False
    
    MsgBox "View function is empty, now set to: " & View
    
    End Select

Case False

    Select Case What
    'View error message
    Case "Errr"
    View = True
    
    'View screen update
    Case "Updte"
    View = False
    
    'View Alerts
    Case "Alrt"
    View = False
    
    Case ""
    View = False
    
    MsgBox "View function is empty, now set to: " & View
    
    End Select
End Select
End Function



Function SeeText(SubName As String)

    Msg = "Error # " & Str(Err.Number) & Chr(13) _
    & SubName & " genarated a error. Source: " & Err.Source & Chr(13) _
    & "Error Line: " & Erl & Chr(13) _
    & Err.Description
    
    Answer = MsgBox(Msg, vbQuestion + vbOKCancel, "Error", Err.HelpFile, Err.HelpContext)
    
    If Answer = vbCancel Then
    End
    
    End If
    
End Function
