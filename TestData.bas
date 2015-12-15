Attribute VB_Name = "TestData"
Sub AddTestData()

SubName = "'AddTestData'"
If View("Errr") = True Then On Error GoTo ErrorText:
Application.ScreenUpdating = View("Updte")
Application.DisplayAlerts = View("Alrt")

Error.DebugTekst Tekst:="Start", _
                        FunctionName:=SubName
'----Start
Dim StartPos As Range
Dim Fvalue As String

1
Sheets("Basisgeg.").Range("O1").Value = "TestData"
LeegMaken.BoekhoudingLeegMaken
Error.DebugTekst "Boekhouding inhoud gewist"
2
Admin.ShowAllSheets

FindArray = Array("Basisgeg.", "Boekingslijst", "Factuurlijst", "Factuur invoer", "Artikelen", "Debiteuren")
On Error Resume Next
'Get all starting positions
For i = 0 To UBound(FindArray)
    Fvalue = FindArray(i)
    With Sheets("TestData").Columns(1) 'beginnen in kolom 1
                Set StartPos = .Find(What:=Fvalue, _
                                After:=.Cells(1), _
                                LookIn:=xlValues, _
                                LookAt:=xlWhole, _
                                SearchOrder:=xlByRows, _
                                SearchDirection:=xlNext, _
                                MatchCase:=False)
    End With
    If Not StartPos Is Nothing Then
        FindArray(i) = FindArray(i) & "|" & StartPos.Row + 1
        Error.DebugTekst "Start postition address of " & Fvalue & " = " _
                            & StartPos.Address & ". Saved start row = " & StartPos.Row + 1
    End If
Next i

If View("Errr") = True Then On Error GoTo ErrorText:
10
With Sheets("Basisgeg.")
    X = FindArray(0)
    Y = InStr(X, "|")
    Z = Len(X)
    ZA = Left(X, Z - (Z - Y + 1))
    StartRow = Right(X, Z - Y)
    Error.DebugTekst "StartRow of " & ZA & " = " & StartRow

    'Eerste rij gegevens kolommen (bedrijfs gegevens ed)
    Sheets("TestData").Range("C" & StartRow & ":C" & StartRow + 10).Copy
        .Range("B2").PasteSpecial (xlPasteValues)
    
    'Tweede rij gegevens kolommen (adres gegevens ed)
    Sheets("TestData").Range("C" & StartRow + 11 & ":C" & StartRow + 21).Copy
        .Range("E2").PasteSpecial (xlPasteValues)
        
    'Belasting tarievengroepen
    Sheets("TestData").Range("D" & StartRow + 23 & ":D" & StartRow + 25).Copy
        .Range("C14").PasteSpecial (xlPasteValues)
    Sheets("TestData").Range("E" & StartRow + 23 & ":E" & StartRow + 26).Copy
        .Range("D14").PasteSpecial (xlPasteValues)
    
     'Instellingen
    Sheets("TestData").Range("C" & StartRow + 27 & ":D" & StartRow + 27).Copy
        .Range("C21").PasteSpecial (xlPasteValues)
    Sheets("TestData").Range("C" & StartRow + 28).Copy
        .Range("C22").PasteSpecial (xlPasteValues)
End With
Error.DebugTekst Tekst:="Basisgegevens leeg gehaald en voorbeeld tekst ingevoerd"

20
With Sheets("Boekingslijst")
    X = FindArray(1)
    Y = InStr(X, "|")
    Z = Len(X)
    ZA = Left(X, Z - (Z - Y + 1))
    StartRow = Right(X, Z - Y)
    Error.DebugTekst "StartRow of " & ZA & " = " & StartRow

    Sheets("TestData").Range("A" & StartRow & ":I" & StartRow + 9).Copy
        .Range("C4").PasteSpecial (xlPasteValues)
End With
Error.DebugTekst Tekst:="Boekingslijst is gevuld"

30
With Sheets("Factuurlijst")
    X = FindArray(1)
    Y = InStr(X, "|")
    ZA = Left(X, Z - (Z - Y + 1))
    StartRow = Right(X, Z - Y)
    Error.DebugTekst "StartRow of " & ZA & " = " & StartRow

    Sheets("TestData").Range("A" & StartRow & ":CE" & StartRow + 9).Copy
        .Range("A2").Insert Shift:=xlDown
    Application.CutCopyMode = False
End With
Error.DebugTekst Tekst:="Factuurlijst is gevuld"

60
With Sheets("Artikelen")
    X = FindArray(1)
    Y = InStr(X, "|")
    ZA = Left(X, Z - (Z - Y + 1))
    StartRow = Right(X, Z - Y)
    Error.DebugTekst "StartRow of " & ZA & " = " & StartRow

    Sheets("TestData").Range("A" & StartRow & ":G" & StartRow + 9).Copy
        .Range("A4").PasteSpecial (xlPasteValues)
End With
Error.DebugTekst Tekst:="Artikelen ingevoerd"


70
With Sheets("Debiteuren")
    X = FindArray(1)
    Y = InStr(X, "|")
    ZA = Left(X, Z - (Z - Y + 1))
    StartRow = Right(X, Z - Y)
    Error.DebugTekst "StartRow of " & ZA & " = " & StartRow

    Sheets("TestData").Range("A" & StartRow & ":K" & StartRow + 9).Copy
        .Range("A4").PasteSpecial (xlPasteValues)
End With
Error.DebugTekst Tekst:="Debiteuren ingevoerd"

Admin.ActivateWorkModus

'--------End Function
Error.DebugTekst Tekst:="Finish", FunctionName:=SubName
Exit Sub

ErrorText:
If Err.Number <> 0 Then SeeText (SubName)

Resume Next
End Sub

