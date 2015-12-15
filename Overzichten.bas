Attribute VB_Name = "Overzichten"
Sub GenereerAfdrukBoeking()
    
SubName = "'GenereerAfdrukBoeking'"
If View("Errr") = True Then On Error GoTo ErrorText:
Application.ScreenUpdating = View("Updte")
Application.DisplayAlerts = View("Alrt")

Error.DebugTekst Tekst:="Start", _
                        FunctionName:=SubName
'----Start
    
    
    EindeBoeking = Sheets("Boekingslijst").Range("C4").End(xlDown).Row
    
    Sheets("Boekingslijst").Range("B3:O1048576").AdvancedFilter Action:=xlFilterCopy _
        , CriteriaRange:=Sheets("Afdruk boekingen").Range("A5:N17"), CopyToRange:= _
        Sheets("Afdruk boekingen").Range("A21:N1048576"), Unique:=False
    
    EindeAfdruk = Sheets("Afdruk Boekingen").Range("A22").End(xlDown).Row
    
    With Range("A21:N" & EindeAfdruk)
        .Interior.ColorIndex = xlNone
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        .Borders(xlEdgeLeft).LineStyle = xlNone
        .Borders(xlEdgeTop).LineStyle = xlNone
        .Borders(xlEdgeBottom).LineStyle = xlNone
        .Borders(xlEdgeRight).LineStyle = xlNone
        .Borders(xlInsideVertical).LineStyle = xlNone
        .Borders(xlInsideHorizontal).LineStyle = xlNone
        .Font.ColorIndex = 0
    End With
    
    Range("G19").Value = WorksheetFunction.Sum(Range("G22:G" & EindeAfdruk)) 'totaal inkomsten
    Range("H19").FormulaR1C1 = WorksheetFunction.Sum(Range("H22:H" & EindeAfdruk)) 'Totaal uitgaven
    Range("J19").FormulaR1C1 = WorksheetFunction.Sum(Range("J22:J" & EindeAfdruk)) 'Totaal omzetbelasting
    Range("K19").FormulaR1C1 = WorksheetFunction.Sum(Range("K22:K" & EindeAfdruk)) 'Totaal voorheffing
    Range("L19").FormulaR1C1 = WorksheetFunction.Sum(Range("L22:L" & EindeAfdruk)) 'Totaal netto inkomsten
    Range("M19").FormulaR1C1 = WorksheetFunction.Sum(Range("M22:M" & EindeAfdruk)) 'Totaal netto uitgaven
   
 '--------End Function
Error.DebugTekst Tekst:="Finish", FunctionName:=SubName
Exit Sub

ErrorText:
If Err.Number <> 0 Then SeeText (SubName)
Resume Next
  
End Sub
