Attribute VB_Name = "Verwerken"
Sub FactuurVerwerken()
Attribute FactuurVerwerken.VB_Description = "De macro is opgenomen op 25-8-2005 door Michel Oltheten."
Attribute FactuurVerwerken.VB_ProcData.VB_Invoke_Func = " \n14"

SubName = "'FactuurVerwerken'"
If View("Errr") = True Then
    On Error GoTo ErrorText:
End If

Application.ScreenUpdating = View("Updte")
Application.DisplayAlerts = View("Alrt")

Admin.ShowAllSheets

HuidigScherm = ActiveSheet.Name

If Range("G24").Value = "Blok" Then
    MsgBox ("Deze factuur is al verwerkt")
    End
End If

'Kijken of er ook een voorbeeld moet worden weergegeven voor het opslaan
If BackgroundFunction.InArray("CheckSave", Sheets("Basisgeg.").Range("C20").Value, _
                                ArrayList:=Array("Altijd", "Verwerken", "Printen|Verwerken", "Verwerken|Opslaan")) Then
    ActiveSh = ActiveSheet.Name
    Application.ScreenUpdating = True
    
    Sheets("Factuur").Select
    InvoiceGood = MsgBox("Is het factuur goed?", vbYesNo, "Factuur goed?")
    
    Application.ScreenUpdating = View("Updte")
    If InvoiceGood = vbNo Then
        Sheets(ActiveSh).Select
        End
    End If
End If

Admin.Bewerkbaar ("Factuurlijst")
    FactuurlijstUitvoeren = Verwerken.Factuurlijst 'uitvoeren van de factuurboeking in het factuur archief
Admin.NietBewerkbaar ("Factuurlijst")

If FactuurlijstUitvoeren = True Then 'checken of het boekingsproces correct is afgesloten
    Admin.Bewerkbaar ("Boekingslijst")
        BoekingslijstUivoeren = Verwerken.Boekingslijst 'uitvoeren van de factuurboeking in de boekingslijst
    Admin.NietBewerkbaar ("Boekingslijst")
    
    If BoekingslijstUivoeren = True Then 'checken of het boekingsproces correct is afgesloten
        'verwerken is voltooid
        VervolgActie = MsgBox(Prompt:="De factuur is verwerkt!" _
                        & vbNewLine & vbNewLine & "Is het nog nodig om nog een actie uit te voeren?" _
                        & vbNewLine & "Bijvoorbeeld: Verzenden als PDF, Emailen, Afdrukken?" _
                        & vbNewLine & vbNewLine & vbNewLine & "!!!! HET IS HIERNA NIET MEER MOGELIJK OM DIT TE DOEN !!!!", _
                        Buttons:=vbYesNo, Title:="Factuur verwerkt!")
        
        If VervolgActie = vbNo Then
            LeegMaken.FactuurInvoerLeeg 'Factuurinvoer leegmaken
        Else 'Factuurinvoer voor wijzigingen beveiligen, alleen de knoppen zijn nog te gebruiken
            Admin.Bewerkbaar ("Factuur invoer")
            
            With Sheets("Factuur invoer")
                .EnableSelection = xlNoSelection
                .Range("D2").Locked = True 'Blokkeren dat klant veranderd kan worden
                .Range("G24").Value = "Blok"  'Knop /verwerken\ blokkeren
            End With
            
            Admin.NietBewerkbaar ("Factuur invoer")
            
            MsgBox "De factuur invoer is beveiligd tegen bewerking! Het is alleen nog mogelijk om de knoppen te gebruiken" _
                    & vbNewLine & vbNewLine & "Wanneer je een nieuw factuur wil beginnen klik dan op de knop 'OVERNIEUW'."
        End If

    Else
        MsgBox "Er is iets fout gegaan bij het verwerken van het factuur in de boekingslijst" & vbNewLine _
            & "de foutcode is: " & BoekingslijstUivoeren & vbNewLine & vbNewLine _
            & "Neem contact op met de software programeur: " & Sheets("Basisgeg.").Range("H26")
    End If
Else
    MsgBox "Er is iets fout gegaan bij het verwerken van het factuur in het factuurarchief" & vbNewLine _
        & "de foutcode is: " & FactuurlijstUitvoeren & vbNewLine & vbNewLine _
        & "Neem contact op met de software programeur: " & Sheets("Basisgeg.").Range("H26")
End If

Admin.ShowOneSheet (HuidigScherm)

Exit Sub

ErrorText:
If Err.Number <> 0 Then
    SeeText (SubName)
    End If
    Resume Next

End Sub

Function Factuurlijst() As String

SubName = "'Factuurlijst'"
If View("Errr") = True Then
    On Error GoTo ErrorText:
End If


Application.ScreenUpdating = View("Updte")
Application.DisplayAlerts = View("Alrt")

'------------------Verwerken van facturen in de factuurlijst

'Controleer factuurnr berekening
1    Sheets("Factuur invoer").Select
    FactuurNummer = Range("H2").Value
    Range("H2:H3").Copy
    
2    Sheets("Factuurlijst").Select
    LaatsteRij = Range("A1").End(xlDown).Row
    
    With Sheets("Factuurlijst")
3    If .Range("A" & LaatsteRij).Value = -1 Then
4        .Range("B" & LaatsteRij).PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
            False, Transpose:=True 'Kopiëer factuurnr + datum in laatste rij (rij -1)
5    'Factuurnummer berekening invoeren in rij 0
        .Range("B" & LaatsteRij - 1).FormulaR1C1 = BackgroundFunction.FormuleProvider("FactuurNrLijst")
        .Range("C" & LaatsteRij - 1).FormulaR1C1 = "=IF(R[1]C="""","""",R[1]C)"
6
    'Factuurnummer opbouw opnieuw inzetten
        .Range("C" & LaatsteRij + 18).FormulaR1C1 = BackgroundFunction.FormuleProvider("FactuurVolgNr")
        .Range("C" & LaatsteRij + 19).FormulaR1C1 = BackgroundFunction.FormuleProvider("Voorloop 1ste-0")
        .Range("C" & LaatsteRij + 20).FormulaR1C1 = BackgroundFunction.FormuleProvider("Voorloop 2de-0")
        .Range("C" & LaatsteRij + 21).FormulaR1C1 = BackgroundFunction.FormuleProvider("Voorloop 0en")
        
        
7        If .Range("B" & LaatsteRij).Value = Range("B" & LaatsteRij - 1).Value Then 'Controle op factuurnummers overeenkomen
            .Range("B" & LaatsteRij - 1, "C" & LaatsteRij).ClearContents 'Controle cellen weer leeg maken

        'Maak nieuwe regel aan in Factuurlijst
8           .Rows("2:2").Insert Shift:=xlDown 'Maak een nieuwe regel aan
            
            Sheets("Factuur invoer").Select
            Range("T1:T80").Copy 'Kopiëer alle data
                        
            Sheets("Factuurlijst").Select
            .Range("C2").PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
                False, Transpose:=True 'Plak alle data in de nieuwe rij
            .Range("B2").Value = FactuurNummer 'factuurnummer vastzetten
            .Range("A2").Value = Range("A3").Value + 1 'volgnr berekenen
            
            'opmaak netjes maken
9            With Rows("2:2")
                .Font.Bold = False
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlBottom
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = False
            End With
            '---
        
10        Else 'Factuurnummer berekening klopt niet
            MsgBox "Factuurnummering berekening klopt niet meer" & vbNewLine _
                & "Het factuur kan niet worden opgeslagen en verzonden" & vbNewLine & vbNewLine _
                & "Neem contact op met de software programeur: " & Sheets("Basisgeg.").Range("H26")
            Range("B" & LaatsteRij - 1, "C" & LaatsteRij).ClearContents 'Controle cellen weer leeg maken
            
            Factuurlijst = "FacNrControlFalse"
            
            Exit Function
        End If
11    Else
        MsgBox "Het programma kan de juiste regel niet vinden voor een factuurnummer controle" & vbNewLine & vbNewLine _
            & "Neem contact op met de software programeur: " & Sheets("Basisgeg.").Range("H26")
        Factuurlijst = "LastRowNotFound"
        Exit Function
    End If
    End With

Factuurlijst = True
Exit Function

ErrorText:
If Err.Number <> 0 Then
    SeeText (SubName)
    End If
    Resume Next

End Function

Function Boekingslijst() As String

SubName = "'Boekingslijst'"
If View("Errr") = True Then
    On Error GoTo ErrorText:
End If

Application.ScreenUpdating = View("Updte")
Application.DisplayAlerts = View("Alrt")

'------------------Verwerken van facturen in de boekingslijst

'Bufferen van data
1 Sheets("Factuurlijst").Select
Boekingsdatum = Range("C2").Value 'Boekingsdatum

'Achternaam + Factuurnummer voor omschrijving
Klantnr = Range("D2").Value

2 With Sheets("Debiteuren").Columns(1)
    Set AchternaamRij = .Find(What:=Klantnr, _
                                After:=.Cells(1), _
                                LookIn:=xlValues, _
                                LookAt:=xlWhole, _
                                SearchOrder:=xlByRows, _
                                SearchDirection:=xlNext, _
                                MatchCase:=False)
    Achternaam = Sheets("Debiteuren").Range("C" & AchternaamRij.Row).Value
End With

FactuurNummer = Range("B2").Value

Omschrijving = Achternaam & "-" & FactuurNummer

Categorie = Range("F2").Value 'Categorie

'BTW zien of alles gelijk is (dan 1 regel aanmaken, anders per tarief 1 regel)
Artikel = Array("P2", "V2", "AB2", "AH2", "AN2", "AT2", "AZ2", _
                "BF2", "BL2", "BR2", "BX2", "CD2") 'Alle BTW tarieven

art = 0 'teller op 0

BTWHoog = False
BTWLaag = False
BTWNul = False

3 For art = 0 To UBound(Artikel) 'check de verschillende BTW tarieven
    If Range(Artikel(art)).Value <> "" Then
        If Range(Artikel(art)).Value = Sheets("Basisgeg.").Range("B10").Value Then
            BTWHoog = True
        ElseIf Range(Artikel(art)).Value = Sheets("Basisgeg.").Range("B11").Value Then
            BTWLaag = True
        ElseIf Range(Artikel(art)).Value = Sheets("Basisgeg.").Range("B12").Value Then
            BTWNul = True
        Else
            Boekingslijst = "Fout BTW tarief bepaling"
            Exit Function
        End If
    End If
Next art

BTW = 0 'set de hoeveelheid BTW tarieven op 1

'Tel de hoeveelheid tarieven
4 If BTWHoog Then
    BTW = BTW + 1
End If
If BTWLaag Then
    BTW = BTW + 1
End If
If BTWNul Then
    BTW = BTW + 1
End If

'Bedrag (afhankelijk van BTW tarieven)
5 If BTW = 1 Then
    Bedrag = Range("H2").Value 'Totaal bedrag
    BTWTarief = Range(Artikel(0)).Value
6 ElseIf BTW > 1 Then
    Dim ArtikelPrijs(0 To 11) As Double
    Dim KortingPercentage As Double
    Dim KortingBedrag As Double
    Dim ArtikelPrijsEx As Double
    Dim BTWbedragHoog As Double
    Dim BTWbedragLaag As Double
    Dim BTWbedragNul As Double
    
7    For art = 0 To UBound(Artikel) 'reken het Artikelbedrag per artikel (incl korting)
        If Range(Artikel(art)).Value <> "" Then
            If Range("J2").Value = "Bedrag" Then
                ArtikelPrijsEx = Cells(Range(Artikel(art)).Row, Range(Artikel(art)).Column - 1).Value * Cells(Range(Artikel(art)).Row, Range(Artikel(art)).Column - 2).Value
                KortingBedrag = Left(Cells(Range(Artikel(art)).Row, Range(Artikel(art)).Column + 1).Value, Len(Cells(Range(Artikel(art)).Row, Range(Artikel(art)).Column + 1)) - 1)
                ArtikelPrijs(art) = ArtikelPrijsEx - KortingBedrag
            ElseIf Range("J2").Value = "Percentage" Then
                ArtikelPrijsEx = Cells(Range(Artikel(art)).Row, Range(Artikel(art)).Column - 1).Value * Cells(Range(Artikel(art)).Row, Range(Artikel(art)).Column - 2).Value
                KortingPercentage = Left(Cells(Range(Artikel(art)).Row, Range(Artikel(art)).Column + 1).Value, Len(Cells(Range(Artikel(art)).Row, Range(Artikel(art)).Column + 1)) - 1)
                ArtikelPrijs(art) = ArtikelPrijsEx - (ArtikelPrijsEx * (KortingPercentage / 100))
            Else
                Boekingslijst = "Korting niet gesnapt"
                Exit Function
            End If
        End If
    Next art
    
8    For art = 0 To UBound(ArtikelPrijs) 'reken het BTWbedrag per tarief
        If Range(Artikel(art)).Value <> "" Then
            'bedrag per tarief berekeken
            If Range(Artikel(art)).Value = Sheets("Basisgeg.").Range("B10").Value Then
                BTWbedragHoog = BTWbedragHoog + ArtikelPrijs(art) 'exlusief BTW
            ElseIf Range(Artikel(art)).Value = Sheets("Basisgeg.").Range("B11").Value Then
                BTWbedragLaag = BTWbedragLaag + ArtikelPrijs(art) 'exlusief BTW
            ElseIf Range(Artikel(art)).Value = Sheets("Basisgeg.").Range("B12").Value Then
                BTWbedragNul = BTWbedragNul + ArtikelPrijs(art) 'exlusief BTW
            End If
        End If
    Next art

  'BTW berekening checker
    BTWbedragHoog = BTWbedragHoog * (1 + Sheets("Basisgeg.").Range("B10").Value) 'inclusief BTW
    BTWbedragLaag = BTWbedragLaag * (1 + Sheets("Basisgeg.").Range("B11").Value) 'inclusief BTW
    BTWbedragNul = BTWbedragNul * (1 + Sheets("Basisgeg.").Range("B12").Value) 'inclusief BTW
    
    'Controle BTW berekeningen
9    If Range("H2").Value = (BTWbedragHoog + BTWbedragLaag + BTWbedragNul + Range("K2").Value) Then
10    Else
        MsgBox "BTW berekening is niet juist verlopen" & vbNewLine _
                & "Verschil is: " & Range("H2").Value - (BTWbedragHoog + BTWbedragLaag + BTWbedragNul)
        Boekingslijst = "Fout BTW bedrag"
        Exit Function
    End If

11 Else
    Boekingslijst = "Fout BTW tarief"
    Exit Function
End If

'-------------------------------------------Buffer compleet

'Alle informatie op de juiste plek zetten
Sheets("Boekingslijst").Select

BeginRij = 2

Opnieuw: 'opnieuw data wegschrijven
LaatsteRij = Range("D" & BeginRij).End(xlDown).Row 'laatst gevulde rij bepalen
Rij = 1 'Rij teller voor verschillende BTW tarieven

12 If IsEmpty(Range("C" & LaatsteRij + 1).Value) Then 'checken of rij echt leeg is
    Range("C" & LaatsteRij + 1).Value = Boekingsdatum
    Range("D" & LaatsteRij + 1).Value = Omschrijving
    Range("E" & LaatsteRij + 1).Value = Categorie
    
    If BTW = 1 Then 'alles is 1 BTW tarief
        Range("G" & LaatsteRij + 1).Value = BTWTarief
        Range("H" & LaatsteRij + 1).Value = Bedrag
        
    Else 'er zijn meer BTW tarieven, nu per tarief een nieuwe regel aanmaken
        If BTWHoog Then
            OmschrijvingHoog = Omschrijving & " | " & Sheets("Basisgeg.").Range("B10").Value * 100 & "% BTW"
            
            Range("D" & LaatsteRij + Rij).Value = OmschrijvingHoog
            Range("G" & LaatsteRij + Rij).Value = Sheets("Basisgeg.").Range("B10").Value
            Range("H" & LaatsteRij + Rij).Value = BTWbedragHoog
        
            Rij = Rij + 1
        End If
        If BTWLaag Then
            OmschrijvingLaag = Omschrijving & " | " & Sheets("Basisgeg.").Range("B11").Value * 100 & "% BTW"
            If IsEmpty(Range("D" & LaatsteRij + Rij).Value) Then 'checken of de volgende rij leeg is
                Range("C" & LaatsteRij + Rij).Value = Boekingsdatum
                Range("D" & LaatsteRij + Rij).Value = OmschrijvingLaag
                Range("E" & LaatsteRij + Rij).Value = Categorie
                Range("G" & LaatsteRij + Rij).Value = Sheets("Basisgeg.").Range("B11").Value
                Range("H" & LaatsteRij + Rij).Value = BTWbedragLaag
            Else 'volgende rij is al gevuld
                LaatsteRij = Range("D" & LaatsteRij).End(xlDown).Row 'laatst gevulde rij bepalen
                
                If LaatsteRij = Range("C4").End(xlDown).Row Then 'check of volgende rij ook leeg is
                    Range("C" & LaatsteRij + Rij).Value = Boekingsdatum
                    Range("D" & LaatsteRij + Rij).Value = OmschrijvingLaag
                    Range("E" & LaatsteRij + Rij).Value = Categorie
                    Range("G" & LaatsteRij + Rij).Value = Sheets("Basisgeg.").Range("B11").Value
                    Range("H" & LaatsteRij + Rij).Value = BTWbedragLaag
                Else
                    If LaatsteRij > Range("C4").End(xlDown).Row Then 'omschrijving is vergeten
                        Range("C" & LaatsteRij + Rij).Value = Boekingsdatum
                        Range("D" & LaatsteRij + Rij).Value = OmschrijvingLaag
                        Range("E" & LaatsteRij + Rij).Value = Categorie
                        Range("G" & LaatsteRij + Rij).Value = Sheets("Basisgeg.").Range("B11").Value
                        Range("H" & LaatsteRij + Rij).Value = BTWbedragLaag
                    Else 'datum is vergeten
                        Range("C" & LaatsteRij + Rij).Value = Boekingsdatum
                        Range("D" & LaatsteRij + Rij).Value = OmschrijvingLaag
                        Range("E" & LaatsteRij + Rij).Value = Categorie
                        Range("G" & LaatsteRij + Rij).Value = Sheets("Basisgeg.").Range("B11").Value
                        Range("H" & LaatsteRij + Rij).Value = BTWbedragLaag
                    End If
                End If
            End If
            
            Rij = Rij + 1
        End If
        
        If BTWNul Then
            OmschrijvingNul = Omschrijving & " | " & Sheets("Basisgeg.").Range("B12").Value & "% BTW"
            If Rij > 1 Then 'BTWHoog of BTWLaag is ook ingevuld
                If IsEmpty(Range("C" & LaatsteRij + Rij).Value) Then 'checken of de volgende rij leeg is
                    Range("C" & LaatsteRij + Rij).Value = Boekingsdatum
                    Range("D" & LaatsteRij + Rij).Value = OmschrijvingNul
                    Range("E" & LaatsteRij + Rij).Value = Categorie
                    Range("G" & LaatsteRij + Rij).Value = Sheets("Basisgeg.").Range("B12").Value
                    Range("H" & LaatsteRij + Rij).Value = BTWbedragNul
                Else 'volgende rij is al gevuld
                    LaatsteRij = Range("D" & LaatsteRij).End(xlDown).Row 'laatst gevulde rij bepalen
                    
                    If LaatsteRij = Range("C4").End(xlDown).Row Then 'check of volgende rij ook leeg is
                        Range("C" & LaatsteRij + Rij).Value = Boekingsdatum
                        Range("D" & LaatsteRij + Rij).Value = OmschrijvingNul
                        Range("E" & LaatsteRij + Rij).Value = Categorie
                        Range("G" & LaatsteRij + Rij).Value = Sheets("Basisgeg.").Range("B12").Value
                        Range("H" & LaatsteRij + Rij).Value = BTWbedragNul
                    Else
                        If LaatsteRij > Range("C4").End(xlDown).Row Then 'omschrijving is vergeten
                            Range("C" & LaatsteRij + Rij).Value = Boekingsdatum
                            Range("D" & LaatsteRij + Rij).Value = OmschrijvingNul
                            Range("E" & LaatsteRij + Rij).Value = Categorie
                            Range("G" & LaatsteRij + Rij).Value = Sheets("Basisgeg.").Range("B12").Value
                            Range("H" & LaatsteRij + Rij).Value = BTWbedragNul
                        Else 'datum is vergeten
                            Range("C" & LaatsteRij + Rij).Value = Boekingsdatum
                            Range("D" & LaatsteRij + Rij).Value = OmschrijvingNul
                            Range("E" & LaatsteRij + Rij).Value = Categorie
                            Range("G" & LaatsteRij + Rij).Value = Sheets("Basisgeg.").Range("B12").Value
                            Range("H" & LaatsteRij + Rij).Value = BTWbedragNul
                        End If
                    End If
                End If
            Else 'BTWHoog of BTWLaag is niet geweest
                Boekingslijst = "Fout invullen meedere BTW tarieven"
            End If
        End If
    End If
Else 'overnieuw een lege plek vinden
    BeginRij = LaatsteRij + 1
    GoTo Opnieuw
End If
'Klaar

Boekingslijst = True
Exit Function

ErrorText:
If Err.Number <> 0 Then
    SeeText (SubName)
    End If
    Resume Next

End Function

Sub EmailFactuur()

SubName = "'EmailFactuur'"
If View("Errr") = True Then
    On Error GoTo ErrorText:
End If

Dim FactuurNummer As String

10
'Krijg de factuurnummer
FactuurNummer = Sheets("Factuur").Range("H17").Value
'check of de factuur al verwerkt is anders eerst verwerken
If BackgroundFunction.FactuurCheck(FactuurNummer) = False Then
    If FactuurNummer = "" Or IsEmpty(FactuurNummer) Then
        MsgBox "Er is geen factuurnummer die gebruikt kan worden"
    Else
        Verwerken.FactuurVerwerken
    End If
End If

20
AttachedInvoice = SavePDF()

Dim iMsg As Object
Dim iConf As Object
Dim strbody As String
Dim Flds As Variant

30
'Get path and name of logo
If IsEmpty(Sheets("Basisgeg.").Range("C26")) Then BackgroundFunction.GetFile ("Logo")

LogoPath = Sheets("Basisgeg.").Range("C26").Value

40
Set iMsg = CreateObject("CDO.Message")
Set iConf = CreateObject("CDO.Configuration")

Dim objBP As Object

iConf.Load -1    ' CDO Source Defaults
Set Flds = iConf.Fields
With Flds
    .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
    .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") _
                   = "mail.lieskebethke.nl"
    .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 587
    .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
    .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "info@lieskebethke.nl"
    .Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "Levi2307"
    .Update
End With

'Bij de testmodus alleen een BCC laten verzenden naar de beheerder
If BackgroundFunction.InArray("Modus", Sheets("Basisgeg.").Range("O1").Value) Then
    EmailTo = ""
    EmailCopy = ""
    EmailBCC = "webbeheerder@lieskebethke.nl"
50
Else
    'Voor test doeleinden de eerste 100 emails naar de maker sturen
    If Sheets("Basisgeg.").Range("O2").Value < 100 Then
        EmailBCC = "webbeheerder@lieskebethke.nl"
        Sheets("Basisgeg.").Range("O2").Value = Sheets("Basisgeg.").Range("O2").Value + 1
    Else
        EmailBCC = ""
    End If

    If Sheets("Factuur invoer").Range("G5").Value <> "" Then
        EmailTo = Sheets("Factuur invoer").Range("G5").Value
    Else
        GiveEmail = InputBox("Er is geen e-mailadres bij deze relatie opgeslagen" & vbNewLine _
                            & "Hier kan je het adres alsnog opgeven:", "Geef e-mailadres", "e-mailadres")
        
        If GiveEmail = "e-mailadres" Then
            MsgBox "Geen adres opgegeven, er kan geen email worden verzonden"
            Exit Sub
        Else
            'Admin.ShowOneSheet ("Debiteuren")
            
            'Vind de debiteur en voeg daar het emailadres aan toe.
            EmailTo = GiveEmail
        End If
    End If
    
60
    If (Sheets("Basisgeg.").Range("C23").Value = "Ja") Then
        If Sheets("Basisgeg.").Range("E9").Value = "" Then
           GiveEmail = InputBox("Er is geen e-mailadres opgeslagen om een kopie aan te sturen" & vbNewLine _
                            & "Hier kan je het adres alsnog opgeven:", "Geef e-mailadres", "e-mailadres")
        
            If GiveEmail = "e-mailadres" Then
                MsgBox "Geen adres opgegeven, er kan geen kopie worden verzonden"
                EmailCopy = ""
            Else
                EmailTo = GiveEmail
                
                Sheets("Basisgeg.").Range("E9").Value = GiveEmail
            End If
        Else
            EmailCopy = Sheets("Basisgeg.").Range("E9").Value
        End If
    Else
        EmailCopy = ""
    End If
End If

70
With iMsg
    Set .Configuration = iConf
    .To = EmailTo
    .CC = EmailCopy
    .BCC = EmailBCC
    .From = """Factuur """ & Sheets("Basisgeg.").Range("B2").Value & " < " & Sheets("Basisgeg.").Range("E9").Value & ">"
    .Subject = "Nieuw factuur van """ & Sheets("Basisgeg.").Range("B2").Value & Sheets("Factuur").Range("H17").Value
    .HTMLBody = GetBody()
    .AddAttachment AttachedInvoice
    .Send
End With

MsgBox "Email is verzonden"

Exit Sub

ErrorText:
If Err.Number <> 0 Then
    SeeText (SubName)
    End If
    Resume Next

End Sub

Function GetBody()

    Dim StrBodyOpen As String 'opening text of the email
    Dim StrBodyClose As String 'end text of the email
    Dim rngHtml As Range 'Range for the changing body info (factuur)
    Dim rngLogo As Range 'Range of the header
   
SubName = "'GetBody'"
If View("Errr") = True Then
    On Error GoTo ErrorText:
End If

Application.ScreenUpdating = View("Updte")
Application.DisplayAlerts = View("Alrt")
    
100    StrBodyOpen = "Beste " & Sheets("Factuur").Range("D10").Value _
                    & "<br><br>Er is een nieuwe factuur voor je klaargemaakt. Hieronder wordt een HTML-versie weergegeven." _
                    & " Wanneer je deze niet kunt weergeven is er in de bijlage een pdf-bestand toegevoegd." _
                    & "<br>Wanneer er verder vragen zijn kan je deze e-mail beantwoorden, dan komt je vraag bij mij terecht." _
                    & "<br>Hartelijke groet Liesbeth"
'haal beveiliging van het werkblad
If BackgroundFunction.Bewerkbaar("Factuur") Then
110  Set rngHtml = Sheets("Factuur").Range("B1", "K53").SpecialCells(xlCellTypeVisible)
     Set rngLogo = Sheets("Factuur").Range("B1:K5")
     
     If BackgroundFunction.NietBewerkbaar("Factuur") Then
     Else
        MsgBox "Kan Werkblad: 'Factuur' niet opnieuw beveiligen. Programma maakt een critieke stop! Code:GetBody110"
        End
     End If
Else
111  GetBody = False
     If BackgroundFunction.NietBewerkbaar("Factuur") Then
        Exit Function
     Else
        MsgBox "Kan Werkblad: 'Factuur' niet opnieuw beveiligen. Programma maakt een critieke stop! Code:GetBody111"
        End
     End If
End If

120    StrBodyClose = "<br>"

Maken:
    
900    GetBody = StrBodyOpen & RangetoHTML(rngHtml, rngLogo) & StrBodyClose
    
Exit Function

ErrorText:
If Err.Number <> 0 Then
    SeeText (SubName)
    End If
    Resume Next

End Function

Function RangetoHTML(text As Range, logo As Range)

    Dim fso As Object
    Dim ts As Object
    Dim TempFile As String
    Dim TempWB As Workbook
    Dim BaseFile As Workbook

SubName = "'RangeToHtml'"
If View("Errr") = True Then
    On Error GoTo ErrorText:
End If


Application.ScreenUpdating = View("Updte")
Application.DisplayAlerts = View("Alrt")
    
1    TempFile = Environ$("temp") & "\" & Format(Now, "dd-mm-yy hh-mm-ss") & ".htm"
     Set BaseFile = ThisWorkbook
2    Set TempWB = Workbooks.Add(1)

     Windows(BaseFile.Name).Activate
4    text.Copy
     
     Windows(TempWB.Name).Activate
10    With TempWB.Sheets(1)
        .Cells(1).PasteSpecial Paste:=8
        .Cells(1).PasteSpecial xlPasteValues, , False, False
        .Cells(1).PasteSpecial xlPasteFormats, , False, False
        .Cells(1).Select
        Application.CutCopyMode = False
        On Error Resume Next
        .DrawingObjects.Visible = True
        .DrawingObjects.Delete
        
        If View("Errr") = True Then
            On Error GoTo ErrorText:
        End If
19    End With

'Add logo
     Windows(BaseFile.Name).Activate
3    logo.Copy
     Windows(TempWB.Name).Activate
     
     'Paste the header range for the logo
     Range("A1").Select
     ActiveSheet.Paste
     
    'Publish the sheet to a htm file
20    With TempWB.PublishObjects.Add( _
         SourceType:=xlSourceRange, _
         Filename:=TempFile, _
         Sheet:=TempWB.Sheets(1).Name, _
         Source:=TempWB.Sheets(1).UsedRange.Address, _
         HtmlType:=xlHtmlStatic)
        .Publish (True)
29    End With

    'Read all data from the htm file into RangetoHTML
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(TempFile).OpenAsTextStream(1, -2)
30    RangetoHTML = ts.readall
31    ts.Close
32    RangetoHTML = Replace(RangetoHTML, "align=center x:publishsource=", _
                          "align=left x:publishsource=")

    'Close TempWB
33    TempWB.Close savechanges:=False

    'Delete the htm file we used in this function
34    Kill TempFile

    Set ts = Nothing
    Set fso = Nothing
    Set TempWB = Nothing

Exit Function
ErrorText:
If Err.Number <> 0 Then
    SeeText (SubName)
    End If
    Resume Next
End Function

