Attribute VB_Name = "NieuweInput"
Sub NieuwDeb()
Attribute NieuwDeb.VB_Description = "De macro is opgenomen op 21-10-2005 door Michel Oltheten."
Attribute NieuwDeb.VB_ProcData.VB_Invoke_Func = " \n14"

SubName = "'NieuwDeb'"
If View("Errr") = True Then
    On Error GoTo ErrorText:
End If


Application.ScreenUpdating = View("Updte")
Application.DisplayAlerts = View("Alrt")

    Range("O2:O10").Copy
    
2 Admin.ShowOneSheet ("Debiteuren")
    
    Einde = Range("C2").End(xlDown).Row 'nieuwe invoer onderaan de pagina toevoegen
    
4    Range("C" & Einde + 1).PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
        
    NieuweDeb = Range("A" & Einde + 1) 'nieuwe toevoeging opslaan om later toe te wijzen
    
    'niet sorteren want anders is de debiteurcode niet meer juist
    'Range("C3:G1000").Sort Key1:=Range("B3"), Order1:=xlAscending, Header:=xlGuess, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom 'debuteuren sorteren op alfabetische volgorde

6 Admin.ShowOneSheet ("Factuur invoer")
    
    With Sheets("Factuur invoer")
        Range("D2").Value = NieuweDeb 'nieuwe invoer toewijzen
8       Range("O2:O14").ClearContents 'velden weer leeg maken
        Range("O7").FormulaR1C1 = "=IF(ISBLANK(R6C15),"""",""Nederland"")"
    End With
    
Exit Sub
ErrorText:
If Err.Number <> 0 Then
    SeeText (SubName)
    End If
    Resume Next

End Sub

Sub NieuwArt()

SubName = "'NiewArt'"
If View("Errr") = True Then
    On Error GoTo ErrorText:
End If

Application.ScreenUpdating = View("Updte")
Application.DisplayAlerts = View("Alrt")

    NieuweArt = Range("O20")
    Range("O20:O24").Copy

3   Admin.ShowOneSheet ("Artikelen")

    Range("C2").End(xlDown).Select
5   ActiveCell.Offset(1, 0).PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
    
    'niet sorteren want anders is de artikelcode niet meer juist _
        (opbouw artikelcode: 1ste 2letters, laatste letter, nr.)
    'Range("C3:G1000").Sort Key1:=Range("C3"), Order1:=xlAscending, Header:=xlGuess, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
        
7   Admin.ShowOneSheet ("Factuur invoer")
    
8   Range("O20:O28").ClearContents 'velden weer leeg maken
    
    'artikel onderaan de lijst met artikelen toevoegen
    ArtikelCode = Range("A21").End(xlUp).Row + 1
    ArtikelOmsch = Range("C21").End(xlUp).Row + 1
    
    If ArtikelCode Or ArtikelOmsch > 9 Then
        If ArtikelCode > ArtikelOmsch Then
13          Range("C" & ArtikelCode).Value = NieuweArt
        Else
15          Range("C" & ArtikelOmsch).Value = NieuweArt
        End If
    Else
18      Range("C9").Value = NieuweArt
    End If
   
Exit Sub
ErrorText:
If Err.Number <> 0 Then
    SeeText (SubName)
    End If
    Resume Next
   
End Sub

Sub NieuwBoekjaarAanmaken()

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
End Sub
