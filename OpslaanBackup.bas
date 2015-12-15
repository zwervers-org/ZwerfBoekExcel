Attribute VB_Name = "OpslaanBackup"
Function SavePDF() As String

SubName = "'SavePDF'"
If View("Errr") = True Then
    On Error GoTo ErrorText:
End If

Application.ScreenUpdating = View("Updte")
Application.DisplayAlerts = View("Alrt")

Dim FileSlct As String
Dim fDialog As Office.FileDialog
Dim PathNow As String
Dim FileSort As String

'Kijken of er ook een voorbeeld moet worden weergegeven voor het opslaan
If BackgroundFunction.InArray("CheckSave", Sheets("Basisgeg.").Range("C20").Value, _
                                Values:=Array("Altijd", "Opslaan", "Printen|Opslaan", "Verwerken|Opslaan")) Then
    ActiveSh = ActiveSheet
    Sheets("Factuur").Select
    InvoiceGood = MsgBox("Is het factuur goed?", vbYesNo, "Factuur goed?")
    
    If InvoiceGood = vbNo Then
        ActiveSh.Select
        End
    End If
End If

'Selecteer opslag locatie en laat deze opslaan
PathNow = Application.ActiveWorkbook.path

'Kijken of er al een pad is opgeslagen om op te slaan.
If IsEmpty(Sheets("Basisgeg.").Range("C25")) Then
    FilePath = BackgroundFunction.GetFolder(PathNow, "PDF")
Else
    FilePath = Sheets("Basisgeg.").Range("C25").Value
End If

'Vind achternaam voor de PDF bestandsnaam
Klantnr = Range("D2").Value

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

FactuurNr = Sheets("Factuur").Range("H17").Value

With Sheets("Factuur")
    .Range("B1", "K52"). _
        ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
        FilePath & Achternaam & " " & FactuurNr & ".pdf", Quality:= _
        xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, _
        OpenAfterPublish:=False
End With

SavePDF = FilePath & "\" & Achternaam & " " & FactuurNr & ".pdf"

Exit Function

ErrorText:
If Err.Number <> 0 Then
    SeeText (SubName)
    End If
    Resume Next

End Function

Sub Backup()

Dim FileSlct As String
Dim fDialog As Office.FileDialog
Dim PathNow As String
Dim FileSort As String

'Selecteer opslag locatie en laat deze opslaan
PathNow = Application.ActiveWorkbook.path

'Kijken of er al een pad is opgeslagen om op te slaan.
If IsEmpty(Sheets("Basisgeg.").Range("C24")) Then
    FilePath = BackgroundFunction.GetFolder(PathNow, "PDF")
Else
    FilePath = Sheets("Basisgeg.").Range("C25").Value
End If

If FilePath = "" Then End 'wanneer er geen pad is geselecteerd

DateNow = Format(Now, "ddMMMyyyy-hhmm")

ActiveWorkbook.SaveCopyAs (FilePath & "\" & DateNow & "-backup.xlsm")

End Sub

