Attribute VB_Name = "Sorteren"
Sub sorteerboekingen()
Attribute sorteerboekingen.VB_Description = "De macro is opgenomen op 16-10-2005 door Michel Oltheten."
Attribute sorteerboekingen.VB_ProcData.VB_Invoke_Func = " \n14"

SubName = "'SorteerBoekingen'"
If View("Errr") = True Then On Error GoTo ErrorText:
Application.ScreenUpdating = View("Updte")
Application.DisplayAlerts = View("Alrt")

Error.DebugTekst Tekst:="Start", _
                        FunctionName:=SubName
'----Start

    EindeAfdruk = Range("A22").End(xlDown).Row
    
    Range("A22:N" & EindeAfdruk).Select

    Selection.Sort Key1:=Range("B22"), Order1:=xlAscending, Header:=xlGuess, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom

 '--------End Function
Error.DebugTekst Tekst:="Finish", FunctionName:=SubName
Exit Sub

ErrorText:
If Err.Number <> 0 Then SeeText (SubName)
Resume Next

End Sub
