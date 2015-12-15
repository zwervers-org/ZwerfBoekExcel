Attribute VB_Name = "LeegMaken"
Sub FactuurInvoerLeeg()

SubName = "'FactuurInvoerLeeg'"
If View("Errr") = True Then
    On Error GoTo ErrorText:
End If

Application.ScreenUpdating = View("Updte")
Application.DisplayAlerts = View("Alrt")

Admin.ShowOneSheet ("Factuur invoer")

Admin.Bewerkbaar ("Factuur invoer")

With Sheets("Factuur invoer")
    
    .Range("D2").Locked = False 'De-blokkeren dat klant veranderd kan worden
    .Range("G24").Value = "" 'Knop /verwerken\ de-blokkeren
    
    .Range("D2").Value = "" 'debiteur
    
    .Range("D6").ClearContents 'datum
    
    .Range("D7").ClearContents 'Categorie
    
    .Range("A9:A20").ClearContents 'Artikelen en beschrijving ed
    .Range("C9:E20").ClearContents 'Artikelen en beschrijving ed
    
    .Range("H9:I20").ClearContents 'Eigen prijs en korting
    
    .Range("D21").ClearContents 'Verzendkosten
    
    .Range("D23:D24").ClearContents 'Totaal Korting
    
    .Range("O2:O14").ClearContents 'Nieuwe klant invoer
    .Range("O7").FormulaR1C1 = "=IF(ISBLANK(R6C15),"""",""Nederland"")"
    
    .Range("O20:O28").ClearContents 'Nieuwe artikel invoer
    
    .Range("D30").ClearContents 'Korting berekenen
    
    'Factuurnummer formule opnieuw instellen
    .Range("H2").FormulaR1C1 = BackgroundFunction.FormuleProvider("FactuurNrInvoer")
    .Range("V9").FormulaR1C1 = BackgroundFunction.FormuleProvider("FactuurVolgNr")
    .Range("V10").FormulaR1C1 = BackgroundFunction.FormuleProvider("Voorloop 1ste-0")
    .Range("V11").FormulaR1C1 = BackgroundFunction.FormuleProvider("Voorloop 2de-0")
    .Range("V12").FormulaR1C1 = BackgroundFunction.FormuleProvider("Voorloop 0en")
    
    .EnableSelection = xlUnlockedCells
End With

Admin.NietBewerkbaar ("Factuur invoer")

Exit Sub
ErrorText:
If Err.Number <> 0 Then
    SeeText (SubName)
    End If
    Resume Next

End Sub

Sub BoekhoudingLeegMaken()
SubName = "'BoekhoudingLeegMaken'"
If View("Errr") = True Then
    On Error GoTo ErrorText:
End If

Application.ScreenUpdating = View("Updte")
Application.DisplayAlerts = View("Alrt")

20
With Sheets("Boekingslijst")
    Einde = .Range("C2").End(xlDown).Row + 10
    .Range("C4:I" & Einde).ClearContents
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
    .PageSetup.RightHeaderPicture.Filename = ""
    .PageSetup.RightHeader = ""
    .Range("S2:S7").ClearContents
    .Range("S2").Value = "Ja"
    Admin.Bewerkbaar ("Factuur")
    On Error Resume Next
        For Each Shp In .Shapes
            Shp.Delete
        Next Shp
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
    .PageSetup.RightHeaderPicture.Filename = ""
    .PageSetup.RightHeader = ""
End With

80
With Sheets("Kwartaaloverzicht")
    .Range("D9").ClearContents
    .PageSetup.RightHeaderPicture.Filename = ""
    .PageSetup.RightHeader = ""
End With

90
With Sheets("Jaaroverzicht")
    .Range("C15:C24,F15:F24").ClearContents
    .PageSetup.RightHeaderPicture.Filename = ""
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

Exit Sub
ErrorText:
If Err.Number <> 0 Then
    SeeText (SubName)
    End If
    Resume Next
End Sub
