Attribute VB_Name = "Afdrukken"
Sub Printfactuur()

SubName = "'PrintFactuur'"
If View("Errr") = True Then On Error GoTo ErrorText:

Application.ScreenUpdating = View("Updte")
Application.DisplayAlerts = View("Alrt")

Error.DebugTekst Tekst:="Start", FunctionName:=SubName
'----Start

'Krijg het factuurnummer
FactuurNummer = Sheets("Factuur").Range("H17").Value

'check of de factuur al verwerkt is anders eerst verwerken
If BackgroundFunction.FactuurCheck(FactuurNummer) = False Then
    If FactuurNummer = "" Or IsEmpty(FactuurNummer) Then
        MsgBox "Er is geen factuurnummer die gebruikt kan worden"
    Else
        SaveFactuur = True
        Verwerken.FactuurVerwerken
        OpslaanBackup.PDFOpslaan
    End If
End If

If SaveFactuur = False Then
        'Kijken of er ook een voorbeeld moet worden weergegeven voor het opslaan
    If BackgroundFunction.InArray("CheckSave", Sheets("Basisgeg.").Range("C20").Value, _
                                    ArrayList:=Array("Altijd", "Printen", "Printen|Opslaan", "Printen|Verwerken")) Then
        ActiveSh = ActiveSheet.Name
        Admin.ShowOneSheet ("Factuur")
        InvoiceGood = MsgBox("Is het factuur goed?", vbYesNo, "Factuur goed?")
        
        If InvoiceGood = vbNo Then
            Admin.ShowOneSheet (ActiveSh)
            Exit Sub
        End If
    End If
End If
    
    Admin.ShowOneSheet ("Factuur")
    ActiveWindow.SelectedSheets.PrintOut copies:=1, collate:=True
    Admin.ShowOneSheet (ActiveSh)
   
'---Finish
Error.DebugTekst Tekst:="Finish", FunctionName:=SubName
    
Exit Sub
ErrorText:
If Err.Number <> 0 Then SeeText (SubName)
 Resume Next
End Sub

Sub OverzichtAfdrukken()

SubName = "'OverzichtAfdrukken'"
If View("Errr") = True Then On Error GoTo ErrorText:

Application.ScreenUpdating = View("Updte")
Application.DisplayAlerts = View("Alrt")

Error.DebugTekst Tekst:="Start", FunctionName:=SubName
'----Start

Dim AantalAfdrukken As Double

TypeOverzicht = ActiveSheet.Name

If Sheets(TypeOverzicht).PageSetup.RightHeader = "&G" Then 'check of het bedrijfslogo bovenaan de pagina staat
    If Sheets(TypeOverzicht).PageSetup.RightHeaderPicture.FileName <> Sheets("Basisgeg.").Range("C26").Value Then _
        Sheets(TypeOverzicht).PageSetup.RightHeader = ""
End If

'Selecteer het bedrijfslogo voor op het overzicht
If Sheets(TypeOverzicht).PageSetup.RightHeader <> "&G" Then 'check of er een afbeelding staat in de kopregel-rechts
    If IsEmpty(Sheets("Basisgeg.").Range("C26").Value) Then BackgroundFunction.GetFile ("Logo") 'Bedrijfslogo bekend?

    'Zet alle pagina instellingen voor de eerste keer  erin
    With Sheets(TypeOverzicht).PageSetup
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeaderPicture.FileName = Application.ActiveWorkbook.path & "\" & Sheets("Basisgeg.").Range("C26").Value
        .RightHeader = "&G"
        .LeftFooter = ""
        .CenterFooter = "Afgedrukt op: &D | &T"
        .RightFooter = "Pagina &P / &N"
        .LeftMargin = Application.InchesToPoints(0.25)
        .RightMargin = Application.InchesToPoints(0.25)
        .TopMargin = Application.InchesToPoints(2)
        .BottomMargin = Application.InchesToPoints(1)
        .HeaderMargin = Application.InchesToPoints(0.5)
        .FooterMargin = Application.InchesToPoints(0.5)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .CenterHorizontally = True
        .CenterVertically = False
        .Orientation = xlPortrait
        .Draft = False
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .PrintErrors = xlPrintErrorsDisplayed
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = False
        .EvenPage.LeftHeader.text = ""
        .EvenPage.CenterHeader.text = ""
        .EvenPage.RightHeader.text = ""
        .EvenPage.LeftFooter.text = ""
        .EvenPage.CenterFooter.text = ""
        .EvenPage.RightFooter.text = ""
        .FirstPage.LeftHeader.text = ""
        .FirstPage.CenterHeader.text = ""
        .FirstPage.RightHeader.text = ""
        .FirstPage.LeftFooter.text = ""
        .FirstPage.CenterFooter.text = ""
        .FirstPage.RightFooter.text = ""
    End With
End If

Select Case TypeOverzicht
    Case "Jaaroverzicht"
        Sheets(TypeOverzicht).PageSetup.PrintArea = "$B$2:$L$27"
        
    Case "Kwartaaloverzicht"
        Sheets(TypeOverzicht).PageSetup.PrintArea = "$B$2:$L$23"
        Inputfield = "D9"
        
    Case "Maandoverzicht"
        Sheets(TypeOverzicht).PageSetup.PrintArea = "$B$2:$L$18"
        Inputfield = "D9"
    
    Case "Afdruk boekingen"
        LaatsteRij = Range("A22").End(xlDown).Row
        Sheets(TypeOverzicht).PageSetup.PrintArea = "$A$22:$N$" & LaatsteRij
        With Sheets(TypeOverzicht).PageSetup
            .Orientation = xlLandscape
            .FitToPagesTall = False
        End With
    
    Case "Boekingslijst"
        LaatsteRij = Range("F4").End(xlDown).Row
        Sheets(TypeOverzicht).PageSetup.PrintArea = "$B$1:$P$" & LaatsteRij
        With Sheets(TypeOverzicht).PageSetup
            .LeftMargin = Application.InchesToPoints(0.25)
            .RightMargin = Application.InchesToPoints(0.25)
            .TopMargin = Application.InchesToPoints(1.18)
            .BottomMargin = Application.InchesToPoints(0.39)
            .HeaderMargin = Application.InchesToPoints(0.12)
            .FooterMargin = Application.InchesToPoints(0.12)
            .Orientation = xlLandscape
            .FitToPagesTall = False
        End With
    
    Case Else
        MsgBox "Er zijn geen instellingen voor dit overzicht, het afdruk bereik wordt in een voorbeeld weergegeven."

End Select

If Not IsEmpty(Inputfield) Then
    'Invulvakje verbergen
    Admin.Bewerkbaar (TypeOverzicht)
    With Range(Inputfield)
        .Interior.Pattern = xlSolid
        .Interior.PatternColorIndex = xlAutomatic
        .Interior.ThemeColor = xlThemeColorDark1
        .Interior.TintAndShade = 0
        .Interior.PatternTintAndShade = 0
        
        .Font.ThemeColor = xlThemeColorDark1
        .Font.TintAndShade = 0
    End With
    Admin.NietBewerkbaar (TypeOverzicht)
End If

'Aantal afdrukken op laten geven
AantalAfdrukken = Application.InputBox(Prompt:="Hoeveel afdrukken zijn er nodig?", _
        Title:="Aantal afdrukken?", Default:=1, Type:=1)
    
    If AantalAfdrukken = 0 Then Exit Sub

    If AantalAfdrukken > 10 Then MoreTen = MsgBox(Prompt:="Meer dan 10x afdrukken?", _
            Buttons:=vbYesNo, Title:="Meer dan 10!?")

    If MoreTen = vbNo Then Exit Sub

Sheets(TypeOverzicht).PrintPreview (True)
Sheets(TypeOverzicht).PrintOut copies:=AantalAfdrukken, collate:=True, IgnorePrintAreas:=False

If Not IsEmpty(Inputfield) Then
    'Invulvakje zichtbaar maken
    Admin.Bewerkbaar (TypeOverzicht)
    With Range(Inputfield)
        .Interior.Pattern = xlSolid
        .Interior.PatternColorIndex = xlAutomatic
        .Interior.Color = 65535
        .Interior.TintAndShade = 0
        .Interior.PatternTintAndShade = 0
        
        .Font.ThemeColor = xlThemeColorLight1
        .Font.TintAndShade = 0
    End With
    Admin.NietBewerkbaar (TypeOverzicht)
End If

'---Finish
Error.DebugTekst Tekst:="Finish", FunctionName:=SubName

Exit Sub

ErrorText:
If Err.Number <> 0 Then SeeText (SubName)
 Resume Next
End Sub


