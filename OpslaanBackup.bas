Attribute VB_Name = "OpslaanBackup"
Sub PDFOpslaan()

SubName = "'PDFOpslaan'"
If View("Errr") = True Then On Error GoTo ErrorText:
Application.ScreenUpdating = View("Updte")
Application.DisplayAlerts = View("Alrt")

Error.DebugTekst Tekst:="Start", _
                        FunctionName:=SubName
'----Start

'Krijg de factuurnummer
FactuurNummer = Sheets("Factuur").Range("H17").Value

'check of de factuur al verwerkt is anders eerst verwerken
If BackgroundFunction.FactuurCheck(FactuurNummer) = False Then
    If FactuurNummer = "" Or IsEmpty(FactuurNummer) Then
        MsgBox "Er is geen factuurnummer die gebruikt kan worden"
        Exit Sub
    Else
        Verwerken.FactuurVerwerken
    End If
End If

Bestandspad = SavePDF

If Not IsEmpty(Bestandspad) Then MsgBox "Opgeslagen in " & Bestandspad

Admin.ShowOneSheet ("Factuur invoer")

'--------End Function
Error.DebugTekst Tekst:="Finish", FunctionName:=SubName
Exit Sub

ErrorText:
If Err.Number <> 0 Then SeeText (SubName)
Resume Next

End Sub
Function SavePDF(Optional Automatic As Boolean) As String

SubName = "'SavePDF'"
If View("Errr") = True Then On Error GoTo ErrorText:
Application.ScreenUpdating = View("Updte")
Application.DisplayAlerts = View("Alrt")
1
Error.DebugTekst Tekst:="Start", _
                        FunctionName:=SubName
'----Start
2
Dim FileSlct As String
Dim fDialog As Office.FileDialog
Dim PathNow As String
Dim FileSort As String
Dim FileName As String
Dim FilePath As String

ThisSheet = ActiveSheet.Name
3
'Krijg de factuurnummer
FactuurNr = Sheets("Factuur").Range("H17").Value
4
'check of de factuur al verwerkt is anders eerst verwerken
If FactuurNr = "" Or IsEmpty(FactuurNr) Then
   MsgBox "Er is geen factuurnummer die gebruikt kan worden"
   Error.DebugTekst "Geen factuurnummer"
   End
End If

5
If Automatic = True Then GoTo SaveDirectly 'Als Automatic is true dan wordt deze uit een andere functie aangeroepen
If BackgroundFunction.FactuurCheck(FactuurNr) = False Then
    'Kijken of er ook een voorbeeld moet worden weergegeven voor het opslaan
    If BackgroundFunction.InArray("CheckSave", Sheets("Basisgeg.").Range("C20").Value, _
                                    ArrayList:=Array("Altijd", "Verwerken", "Printen|Verwerken", "Verwerken|Opslaan")) Then
        ActiveSh = ActiveSheet.Name
       
        GaNaar.bekijkfactuur
        
        InvoiceGood = MsgBox("Is het factuur goed?", vbYesNo, "Factuur goed?")
        
        If InvoiceGood = vbNo Then
            Admin.ShowOneSheet (ActiveSh)
            End
        End If
    End If
End If

SaveDirectly:
10
'Selecteer opslag locatie en laat deze opslaan
PathNow = Application.ActiveWorkbook.path

20
'Kijken of er al een pad is opgeslagen om op te slaan.
If IsEmpty(Sheets("Basisgeg.").Range("C25")) Then
    FilePath = BackgroundFunction.GetFolder(PathNow, "PDF")
Else
    FilePath = Sheets("Basisgeg.").Range("C25").Value
    If BackgroundFunction.TestFolderExist(FilePath) = False Then FilePath = BackgroundFunction.GetFolder(PathNow, "PDF")
    
    If Left(FilePath, 1) = "\" Then FilePath = PathNow & FilePath
End If

'Vind achternaam voor de PDF bestandsnaam
Klantnr = Sheets("Factuur invoer").Range("D2").Value

30
'Admin.ShowOneSheet ("Debiteuren")
'Admin.Bewerkbaar ("Debiteuren")
With Sheets("Debiteuren").Columns(1)
    Set AchternaamRij = .Find(What:=Klantnr, _
                                After:=.Cells(1), _
                                LookIn:=xlValues, _
                                LookAt:=xlWhole, _
                                SearchOrder:=xlByRows, _
                                SearchDirection:=xlNext, _
                                MatchCase:=False)
    Achternaam = Sheets("Debiteuren").Range("C" & AchternaamRij.Row).Value
End With
'Admin.NietBewerkbaar ("Debiteuren")
'Admin.ShowOneSheet (ThisSheet)

FileName = FactuurNr & "_" & Achternaam

50
Admin.ShowOneSheet ("Factuur")
With Sheets("Factuur")
    .Range("B1", "K50").Select
    
    ChDir (FilePath)
        
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:= _
        FilePath & "\" & FileName & ".pdf" _
        , Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
        :=False, OpenAfterPublish:=False
    
    .Range("B1").Select
End With

SavePDF = FilePath & "\" & FileName & ".pdf"

Error.DebugTekst Tekst:="Finish", FunctionName:=SubName

'--------End Function
Error.DebugTekst Tekst:="Finish", FunctionName:=SubName
Exit Function

ErrorText:
Debug.Print Err.Description
If Err.Number <> 0 Then SeeText (SubName)
Resume Next

End Function

Sub Backup()

SubName = "'BackUp'"
If View("Errr") = True Then On Error GoTo ErrorText:
Application.ScreenUpdating = View("Updte")
Application.DisplayAlerts = View("Alrt")

Error.DebugTekst Tekst:="Start", _
                        FunctionName:=SubName
'----Start

Dim FileSlct As String
Dim fDialog As Office.FileDialog
Dim PathNow As String
Dim FileSort As String

'Selecteer opslag locatie en laat deze opslaan
PathNow = Application.ActiveWorkbook.path

'Kijken of er al een pad is opgeslagen om op te slaan.
If IsEmpty(Sheets("Basisgeg.").Range("C24")) Then
    FilePath = BackgroundFunction.GetFolder(PathNow, "Backup")
Else
    FilePath = Sheets("Basisgeg.").Range("C24").Value
    If Left(FilePath, 1) = "\" Then FilePath = PathNow & FilePath
    If BackgroundFunction.TestFolderExist(FilePath) = False Then BackgroundFunction.GetFolder PathNow, "PDF"
End If

If FilePath = "" Then End 'wanneer er geen pad is geselecteerd

DateNow = Format(Now, "ddMMMyyyy-hhmm")

If Right(FilePath, 1) <> "\" Then FilePath = FilePath & "\"

ActiveWorkbook.SaveCopyAs (FilePath & DateNow & "-backup.xlsm")

ActiveSht = ActiveSheet.Name

If ActiveSht <> "Basisgeg." Then Admin.ShowOneSheet ("Basisgeg.")

Admin.Bewerkbaar ("Basisgeg.")
Sheets("Basisgeg.").Range("O10").Value = DateNow
Admin.NietBewerkbaar ("Basisgeg.")

If ActiveSht <> "Basisgeg." Then Admin.ShowOneSheet (ActiveSht)

If BackgroundFunction.IsFile(FilePath & "\" & DateNow & "-backup.xlsm") Then
    Error.DebugTekst ("Backup successfull")
Else
    Error.DebugTekst ("NO BACKUP MADE")
End If

'--------End Function
Error.DebugTekst Tekst:="Finish", FunctionName:=SubName
Exit Sub

ErrorText:
If Err.Number <> 0 Then SeeText (SubName)
Resume Next
End Sub

