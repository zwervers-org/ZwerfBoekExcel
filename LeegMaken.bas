Attribute VB_Name = "LeegMaken"
Sub FactuurInvoerLeeg()

SubName = "'FactuurInvoerLeeg'"
If View("Errr") = True Then On Error GoTo ErrorText:
Application.ScreenUpdating = View("Updte")
Application.DisplayAlerts = View("Alrt")

Error.DebugTekst Tekst:="Start", _
                        FunctionName:=SubName
'----Start

Admin.ShowOneSheet ("Factuur invoer")

Admin.Bewerkbaar ("Factuur invoer")

With Sheets("Factuur invoer")
    
    .Range("V1").Locked = False 'De-blokkeren dat klant veranderd kan worden
    .Range("G24").Value = "" 'Knop /verwerken\ de-blokkeren
    .Range("I2").Value = "" 'Backload factuurnummer
        
    .Range("V1").Value = "" 'debiteur
    .Range("D3").Value = BackgroundFunction.FormuleProvider("Naam") 'Naam Formule
    .Range("D4").Value = BackgroundFunction.FormuleProvider("Adres") 'Straat Formule
    .Range("D5").Value = BackgroundFunction.FormuleProvider("PC_Plaats1") 'PC+Plaat zonder , Formule
    .Range("E5").Value = BackgroundFunction.FormuleProvider("PC_Plaats") 'PC+Plaats Formule
    .Range("F4").Value = BackgroundFunction.FormuleProvider("LandNm") 'Land label Formule
    .Range("F5").Value = BackgroundFunction.FormuleProvider("EmailNm") 'Email label Formule
    .Range("F6").Value = BackgroundFunction.FormuleProvider("TelefoonNm") 'Telefoon label Formule
    .Range("G4").Value = BackgroundFunction.FormuleProvider("Land") 'Land Formule
    .Range("G5").Value = BackgroundFunction.FormuleProvider("Email") 'Email Formule
    .Range("G6").Value = BackgroundFunction.FormuleProvider("Telefoon") 'Telefoon Formule
    .Range("K3").Value = BackgroundFunction.FormuleProvider("OpmerkingNm") 'Opmerking label Formule
    .Range("K4").Value = BackgroundFunction.FormuleProvider("Opmerking") 'Opmerking Formule
    
    .Range("D6").ClearContents 'datum
    
    .Range("D7").ClearContents 'Categorie
    
    .Range("A9:A20").ClearContents 'Artikelen en beschrijving ed
    .Range("C9:E20").ClearContents 'Artikelen en beschrijving ed
    
    .Range("H9:I20").ClearContents 'Eigen prijs en korting
    
    .Range("D21").ClearContents 'Verzendkosten
    
    .Range("D23:D24").ClearContents 'Totaal Korting
    
    With .Range("O2:O14") 'Nieuwe klant invoer
        .Locked = False
        .ClearContents
    End With
    .Range("O7").FormulaR1C1 = "=IF(ISBLANK(R6C15),"""",""Nederland"")"
    
    With .Range("O20:O28") 'Nieuwe artikel invoer
        .Locked = False
        .ClearContents
    End With
        
    
    .Range("D31").ClearContents 'Korting berekenen
    
    'Factuurnummer formule opnieuw instellen
    .Range("H2").FormulaR1C1 = BackgroundFunction.FormuleProvider("FactuurNrInvoer")
    .Range("V9").FormulaR1C1 = BackgroundFunction.FormuleProvider("FactuurVolgNr")
    .Range("V10").FormulaR1C1 = BackgroundFunction.FormuleProvider("Voorloop 1ste-0")
    .Range("V11").FormulaR1C1 = BackgroundFunction.FormuleProvider("Voorloop 2de-0")
    .Range("V12").FormulaR1C1 = BackgroundFunction.FormuleProvider("Voorloop 0en")
    
    .EnableSelection = xlUnlockedCells
End With

Admin.NietBewerkbaar ("Factuur invoer")

'--------End Function
Error.DebugTekst Tekst:="Finish", FunctionName:=SubName
Exit Sub

ErrorText:
If Err.Number <> 0 Then SeeText (SubName)
Resume Next
End Sub

Sub BoekhoudingLeegMaken()

SubName = "'BoekhoudingLeegMaken'"
If View("Errr") = True Then On Error GoTo ErrorText:
Application.ScreenUpdating = View("Updte")
Application.DisplayAlerts = View("Alrt")

Error.DebugTekst Tekst:="Start", _
                        FunctionName:=SubName
'----Start

20
With Sheets("Boekingslijst")
    Einde = .Range("C2").End(xlDown).Row + 10
    .Range("C4:K" & Einde).ClearContents
End With

30
With Sheets("Factuurlijst")
    Einde = .Range("A1").End(xlDown).Row - 2
    If Einde > 1 Then .Range("A2:A" & Einde).EntireRow.Delete
End With

40
With Sheets("Factuur")
    Dim Shp As Shape
    Dim fName As String
    .PageSetup.RightHeaderPicture.FileName = ""
    .PageSetup.RightHeader = ""
    .Range("S2:S7").ClearContents
    .Range("S2").Value = "Ja"
    Admin.Bewerkbaar ("Factuur")
    On Error Resume Next
        For Each Sh In .Shapes
           If Not Application.Intersect(Sh.TopLeftCell, .Range("B1:K5")) Is Nothing Then
                Sh.Delete
           End If
        Next Sh
    On Error GoTo ErrorText
    fName = BackgroundFunction.GetFile("Logo")
    BackgroundFunction.InsertPictureInRange PictureFileName:=fName, TargetCells:=Range("K5: K5"), TargetSheet:=Sheets("Factuur")
    Admin.NietBewerkbaar ("Factuur")
End With

10
With Sheets("Basisgeg.")
    .Range("B2:B9,E2:E9,C14:C16,D14:D17,C20,C21:D21,C22:C27,A37:B100,E37:F100").ClearContents
    .Range("A37:B37,E37:F37").Value = "Voorbeeld"
End With

50
With Sheets("Artikelen")
    Einde = .Range("C2").End(xlDown).Row + 10
    .Range("C4:G" & Einde).ClearContents
End With

60
With Sheets("Debiteuren")
    Einde = .Range("C2").End(xlDown).Row + 10
    .Range("C4:K" & Einde).ClearContents
End With

70
With Sheets("Maandoverzicht")
    .Range("D9").ClearContents
    .PageSetup.RightHeaderPicture.FileName = ""
    .PageSetup.RightHeader = ""
End With

80
With Sheets("Kwartaaloverzicht")
    .Range("D9").ClearContents
    .PageSetup.RightHeaderPicture.FileName = ""
    .PageSetup.RightHeader = ""
End With

90
With Sheets("Jaaroverzicht")
    .Range("C15:C24,F15:F24").ClearContents
    .PageSetup.RightHeaderPicture.FileName = ""
    .PageSetup.RightHeader = ""
End With

100
With Sheets("Afdruk boekingen")
End With

110
With Sheets("Buffer")
End With

120
LeegMaken.FactuurInvoerLeeg

With Sheets("Basisgeg.")
    .Select
    .Range("O1").Value = "Leeg"
    .Range("B2").Select
End With

'--------End Function
Error.DebugTekst Tekst:="Finish", FunctionName:=SubName
Exit Sub

ErrorText:
If Err.Number <> 0 Then SeeText (SubName)
Resume Next
End Sub
