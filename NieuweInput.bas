Attribute VB_Name = "NieuweInput"
Sub NieuwDeb()
Attribute NieuwDeb.VB_Description = "De macro is opgenomen op 21-10-2005 door Michel Oltheten."
Attribute NieuwDeb.VB_ProcData.VB_Invoke_Func = " \n14"

SubName = "'NieuwDeb'"
If View("Errr") = True Then On Error GoTo ErrorText:
Application.ScreenUpdating = View("Updte")
Application.DisplayAlerts = View("Alrt")

Error.DebugTekst Tekst:="Start", _
                        FunctionName:=SubName
'----Start

1 'Beveiliging verwijderen
If Admin.Bewerkbaar("Debiteuren") = False Then GoTo EindSub


2 Admin.ShowOneSheet ("Debiteuren")

'Check of er nog vreemde input staat en die opruimen
If NieuweInput.CheckNwInput(ActiveSheet.Name) = False Then
    MsgBox "Problem with CheckNwInput, system has to end"
    Admin.ShowOneSheet ("Factuur invoer")
    Exit Sub
End If

With Sheets("Debiteuren")
3 'nieuwe invoer onderaan de pagina toevoegen
    Einde1 = .Range("C2").End(xlDown).Row
    Einde2 = .Range("D2").End(xlDown).Row
    Einde3 = .Range("A2500").End(xlUp).Row
    
    If Einde1 > Einde2 Then Einde = Einde1 Else Einde = Einde2
    If Einde3 > Einde Then Einde = Einde3

4 'Waarden uit de invoer plakken
    Sheets("Factuur invoer").Range("O2:O10").Copy

    .Range("C" & Einde + 1).PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
    
    NieuweDeb = NwDebNr(Einde + 1)
    If NieuweDeb = False Then GoTo EindSub
    
10
    With Sheets("Debiteuren").Columns(1)
        Set DebRw = .Find(What:=NieuweDeb, _
                                After:=.Cells(1), _
                                LookIn:=xlValues, _
                                LookAt:=xlWhole, _
                                SearchOrder:=xlByRows, _
                                SearchDirection:=xlNext, _
                                MatchCase:=False)
    End With
    NieuweDeb = DebRw.Row - 2
End With

6   Admin.ShowOneSheet ("Factuur invoer")
    
    With Sheets("Factuur invoer")
        Range("V1").Value = NieuweDeb 'nieuwe invoer toewijzen
8       Range("O2:O14").ClearContents 'velden weer leeg maken
        Range("O7").FormulaR1C1 = "=IF(ISBLANK(R6C15),"""",""Nederland"")"
    End With

'--------End Function
Error.DebugTekst Tekst:="Finish", FunctionName:=SubName

Exit Sub
EindSub:

Error.DebugTekst "Bewerkbaarheid wil niet wijzigingen", SubName
Exit Sub

ErrorText:
If Err.Number <> 0 Then SeeText (SubName)
Resume Next

End Sub

Private Function NwDebNr(ActiveRow As Integer) As String

SubName = "'NwDebNr'"
If View("Errr") = True Then On Error GoTo ErrorText:
Application.ScreenUpdating = View("Updte")
Application.DisplayAlerts = View("Alrt")

Error.DebugTekst Tekst:="Start input:" & vbNewLine _
                    & "ActiveRow: " & ActiveRow, _
                        FunctionName:=SubName
'----Start

1 'nieuwe toevoeging opslaan om later toe te wijzen
With Sheets("Debiteuren")
    .Range("A" & ActiveRow).Value = Application.WorksheetFunction.Max(.Range("A:A")) + 1
    .Range("B" & ActiveRow).Value = .Range("D" & ActiveRow).Value & " " & .Range("C" & ActiveRow).Value
    
    NwDebNr = .Range("A" & ActiveRow).Value

    Error.DebugTekst "Nw debiteur: " & .Range("A" & ActiveRow).Value, SubName

5 'debuteuren sorteren op alfabetische volgorde (achternaam)
    .Range("A4:K" & ActiveRow).Sort Key1:=Range("C4"), Order1:=xlAscending, Header:=xlGuess, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
End With

'--------End Function
Error.DebugTekst Tekst:="Finish", FunctionName:=SubName
Exit Function

ErrorText:
If Err.Number <> 0 Then SeeText (SubName)
Resume Next

End Function

Function CheckNwInput(ActSht As String) As Boolean

SubName = "'CheckNwInput'"
If View("Errr") = True Then On Error GoTo ErrorText:
Application.ScreenUpdating = View("Updte")
Application.DisplayAlerts = View("Alrt")

Error.DebugTekst Tekst:="Start input:" & vbNewLine _
                    & "ActiveRow: " & ActiveRow, _
                        FunctionName:=SubName
'----Start

1 'nieuwe toevoeging opslaan om later toe te wijzen
With Sheets(ActSht)
    'Check of er nieuwe input is gegeven
    MaxRw = 104586
    cEnd = .Range("C" & MaxRw).End(xlUp).Row
    dEnd = .Range("D" & MaxRw).End(xlUp).Row
    SearchEnd = .Range("B3").End(xlDown).Row
    NrEnd = .Range("A3").End(xlDown).Row
    
    MaxEnd = WorksheetFunction.Max(cEnd, dEnd, SearchEnd, NrEnd)
    
    Select Case MaxEnd
        Case SearchEnd, NrEnd
            Error.DebugTekst Tekst:="No new input > SearchEnd or NrEnd are max value"
            CheckNwInputOut = True
            GoTo EindeFunctie
        Case cEnd, dEnd
            For i = NrEnd + 1 To WorksheetFunction.Max(cEnd, dEnd)
                NieuweInput.NwDebNr (i)
            Next
            CheckNwInputOut = True
            GoTo EindeFunctie
        Case Else
            Error.DebugTekst Tekst:="MaxEnd gives error"
            CheckNwDebOut = False
            GoTo EindeFunctie
    End Select
End With

'--------End Function
EindeFunctie:
Error.DebugTekst Tekst:="Finish with " & CheckNwInputOut, FunctionName:=SubName
CheckNwInput = CheckNwInputOut
Exit Function

ErrorText:
If Err.Number <> 0 Then SeeText (SubName)
Resume Next

End Function

Sub NieuwArt()

SubName = "'NieuwArt'"
If View("Errr") = True Then On Error GoTo ErrorText:
Application.ScreenUpdating = View("Updte")
Application.DisplayAlerts = View("Alrt")

Error.DebugTekst Tekst:="Start", _
                        FunctionName:=SubName
'----Start

If Admin.Bewerkbaar("Artikelen") = False Then GoTo EindSub

'Wanneer er een nieuw artikel wordt ingevoerd
    NieuweArt = Sheets("Factuur invoer").Range("O20")

2   Admin.ShowOneSheet ("Artikelen")

'Check of er nog vreemde input staat en die opruimen
If NieuweInput.CheckNwInput(ActiveSheet.Name) = False Then
    MsgBox "Problem with CheckNwInput, system has to end"
    Admin.ShowOneSheet ("Factuur invoer")
    Exit Sub
End If

With Sheets("Artikelen")
3 'nieuwe invoer onderaan de pagina toevoegen
    Einde1 = .Range("C2").End(xlDown).Row
    Einde2 = .Range("D2").End(xlDown).Row
    Einde3 = .Range("C2500").End(xlUp).Row 'Check of er nog handmatig artikelen zijn toegevoegd
    
    If Einde1 > Einde2 Then Einde = Einde1 Else Einde = Einde2
    If Einde3 > Einde Then Einde = Einde3
    
4 'Waarden uit de invoer plakken
        Sheets("Factuur invoer").Range("O20:O24").Copy
        .Range("C" & Einde + 1).PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
            False, Transpose:=True
        If NwArtCode(Einde + 1) = "" Then GoTo EindSub

End With
'Wanneer er een nieuw artikel wordt ingevoerd
7       Admin.ShowOneSheet ("Factuur invoer")
        
8       Range("O20:O28").ClearContents 'velden weer leeg maken
    
        'artikel onderaan de lijst met artikelen toevoegen
        ArtikelCode = Range("A21").End(xlUp).Row + 1
        ArtikelOmsch = Range("C21").End(xlUp).Row + 1
        
        If ArtikelCode Or ArtikelOmsch > 9 Then
            If ArtikelCode > ArtikelOmsch Then
13              Range("C" & ArtikelCode).Value = NieuweArt
            Else
15              Range("C" & ArtikelOmsch).Value = NieuweArt
            End If
        Else
18          Range("C9").Value = NieuweArt
        End If

'--------End Function
Error.DebugTekst Tekst:="Finish", FunctionName:=SubName

Exit Sub
EindSub:

Error.DebugTekst "Probleem met verwerken nieuwe input. Regel: " & Erl, SubName
Exit Sub

ErrorText:
If Err.Number <> 0 Then SeeText (SubName)
Resume Next
   
End Sub

Private Function NwArtCode(ActiveRow As Integer) As String

SubName = "'NwArtCode'"
If View("Errr") = True Then On Error GoTo ErrorText:
Application.ScreenUpdating = View("Updte")
Application.DisplayAlerts = View("Alrt")

Error.DebugTekst Tekst:="Start input:" & vbNewLine _
                    & "ActiveRow: " & ActiveRow, _
                        FunctionName:=SubName
'----Start

1 'Nieuw artikelcode aanmaken
    '(opbouw artikelcode: 1ste 2letters, laatste letter, nr.)
With Sheets("Artikelen")
    ArtNrMax = Application.WorksheetFunction.Max(.Range("A:A")) + 1
    'ArtNr = ActiveRow - 1
    ArtCode = Left(.Range("C" & ActiveRow).Value, 2)
    ArtCode = UCase(ArtCode & Right(.Range("C" & ActiveRow).Value, 1))
    ArtCode = ArtCode & ArtNrMax
    
    Error.DebugTekst "ArtNr: " & ArtNrMax, SubName
    Error.DebugTekst "ArtCode: " & ArtCode, SubName
    
    .Range("A" & ActiveRow).Value = ArtNrMax
    .Range("B" & ActiveRow).Value = ArtCode

2 'Sorteren op Artikelcode
    .Range("A4:G" & ActiveRow).Sort Key1:=Range("B4"), Order1:=xlAscending, Header:=xlGuess, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
End With

NwArtCode = ArtCode

'--------End Function
Error.DebugTekst Tekst:="Finish", FunctionName:=SubName
Exit Function

ErrorText:
If Err.Number <> 0 Then SeeText (SubName)
Resume Next

End Function


Sub NieuwBoekjaarAanmaken()

SubName "'NieuwBoekjaarAanmaken'"
If View("Errr") = True Then On Error GoTo ErrorText:
Application.ScreenUpdating = View("Updte")
Application.DisplayAlerts = View("Alrt")

Error.DebugTekst Tekst:="Start", _
                        FunctionName:=SubName
'----Start

AfsluitCheck = MsgBox("Wilt u een nieuw boekjaar aanmaken en het huidige sluiten?", vbYesNo, "Boekjaar sluiten?")

If AfsluitCheck = vbYes Then
    OpslaanBackup.Backup
    
    Sheets("Basisgeg.").Range("B8").Value = Sheets("Basisgeg.").Range("B8").Value + 1
    
    'Boekingslijst leegmaken
    With Sheets("Boekingslijst")
        LaatsteRij = .Range("C4").End(xlDown).Row
        .Range("C4:I" & LaatsteRij).ClearContents
    End With
    
    'Factuurlijst LeegMaken
    With Sheets("Factuurlijst")

    LaatsteRij = .Range("D2").End(xlDown).Row - 1

OpnieuwFactuur:
            If LaatsteRij >= 2 Then
                LaatsteRij = LaatsteRij - 1 'ga elke keer een rij terug
            Else
                MsgBox "Kan laatste rij niet vinden in 'Factuurlijst'" & vbNewLine _
                        & "Deze moet handmatig geleegd worden"
                GoTo BoekingLegen
            End If
        
        'Check of de juiste rij is gevonden
        If .Range("A" & LaatsteRij).Value = 0 Then GoTo OpnieuwFactuur
        
        .Range("A2:A" & LaatsteRij).EntireRow.Delete
    End With

BoekingLegen:
    'Afdruk boekingen leegmaken
    With Sheets("Afdruk boekingen")
        LaatsteRij = .Range("A22").End(xlDown).Row
        .Range("A22:N" & LaatsteRij).ClearContents
    End With
    
    'Buffer leegmaken
    With Sheets("Buffer")
        LaatsteRij = .Range("A1").End(xlDown).Row
        .Range("A1:A" & LaatsteRij).EntireRow.Delete
    End With
End If

'--------End Function
Error.DebugTekst Tekst:="Finish", FunctionName:=SubName
Exit Sub

ErrorText:
If Err.Number <> 0 Then SeeText (SubName)
Resume Next

End Sub
