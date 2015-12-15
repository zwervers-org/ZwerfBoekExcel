Attribute VB_Name = "Sorteren"
Sub sorteerboekingen()
Attribute sorteerboekingen.VB_Description = "De macro is opgenomen op 16-10-2005 door Michel Oltheten."
Attribute sorteerboekingen.VB_ProcData.VB_Invoke_Func = " \n14"

    EindeAfdruk = Range("A22").End(xlDown).Row
    
    Range("A22:N" & EindeAfdruk).Select

    Selection.Sort Key1:=Range("B22"), Order1:=xlAscending, Header:=xlGuess, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom

End Sub
