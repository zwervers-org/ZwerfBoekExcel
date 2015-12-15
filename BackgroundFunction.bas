Attribute VB_Name = "BackgroundFunction"
Function InArray(WitchArray, strValue, Optional ArrayList)

SubName = "'InArray'"
If View("Errr") = True Then
    On Error GoTo ErrorText:
End If

Application.ScreenUpdating = View("Updte")
Application.DisplayAlerts = View("Alrt")

Dim j
  
Select Case WitchArray
Case "Sheets"
10  ArrayVals = Array()

Case "NotAv"
    ArrayVals = Array()
Case "CheckSave"
    ArrayVals = ArrayList
Case "Database"
    ArrayVals = Array("Factuurlijst", "Boekingslijst", "Artikelen", "Debiteuren", "Afdruk boekingen")
Case "SchuivenHorizontaal"
    ArrayVals = Array()
Case "SchuivenVerticaal"
    ArrayVals = Array("Factuur", "Artikelen", "Boekingslijst", "Debiteuren", "Afdruk boekingen")
Case "Modus"
    ArrayVals = Array("Test modus", "Test modus beveiligd")
Case "ModusBeveiliging"
    ArrayVals = Array("Test modus")
End Select

20  For j = 0 To UBound(ArrayVals)
21    If ArrayVals(j) = CStr(strValue) Then
22      InArray = True
      Exit Function
    End If
  Next
25  InArray = False
Exit Function

ErrorText:
If Err.Number <> 0 Then
    SeeText (SubName)
    End If
    Resume Next
    
End Function

Function ColLett(Col As Integer) As String
     
    If Col > 26 Then
        ColLett = ColLett((Col - (Col Mod 26)) / 26) + Chr(Col Mod 26 + 64)
    Else
        ColLett = Chr(Col + 64)
    End If
     
End Function

Function PrintArray(StrArray As Variant)
  
  For k = 0 To UBound(StrArray)
    txt = txt & k & ": " & StrArray(k) & vbCrLf
  Next k
  
  MsgBox txt

End Function

Function FactuurCheck(FactuurNr As String) As Boolean

SubName = "'FactuurCheck'"
If View("Errr") = True Then
    On Error GoTo ErrorText:
End If

Application.ScreenUpdating = View("Updte")
Application.DisplayAlerts = View("Alrt")

Dim rng As Range

If IsEmpty(FactuurNr) Or FactuurNr = "" Then
    FactuurCheck = False
    Exit Function
End If

With Sheets("Factuurlijst").Columns(2) 'beginnen in kolom 2 en sla de titelrij over
    Set FactuurVinden = .Find(What:=FactuurNr, _
                                After:=.Cells(2), _
                                LookIn:=xlValues, _
                                LookAt:=xlWhole, _
                                SearchOrder:=xlByRows, _
                                SearchDirection:=xlNext, _
                                MatchCase:=False)
    
End With

If Not FactuurVinden Is Nothing Then
    FactuurCheck = True
    
Else
    FactuurCheck = False
End If

Exit Function

ErrorText:
If Err.Number <> 0 Then
    SeeText (SubName)
    End If
    Resume Next
    
End Function

Function GetFolder(strPath As String, FileSort As String) As String

SubName = "'GetFolder'"
If View("Errr") = True Then
    On Error GoTo ErrorText:
End If

Application.ScreenUpdating = View("Updte")
Application.DisplayAlerts = View("Alrt")


Dim fldr As FileDialog
Dim sItem As String

Set fldr = Application.FileDialog(msoFileDialogFolderPicker)

CutFromNr = InStrRev(strPath, "\")

strPath = Left(strPath, CutFromNr)

With fldr
    .Title = "Selecteer de folder voor het opslaan van de" & FileSort
    .AllowMultiSelect = False
    .InitialFileName = strPath
    
    If .Show <> -1 Then End
    
    sItem = .SelectedItems(1)
End With



GetFolder = sItem

Set fldr = Nothing

'Pad opslaan voor later gebruik?
SavePath = MsgBox("Pad opslaan voor later gebruik?", vbYesNo, "Pad opslaan?")

With Sheets("Basisgeg.")
    If SavePath = 6 Then
        Select Case FileSort
            Case "PDF"
                .Range("C25").Value = GetFolder
            Case "Backup"
                .Range("C24").Value = GetFolder
            Case Else
                MsgBox "Er is iets fout gegaan bij het verwerken van de folder gegevens." & vbNewLine _
                    & "De foutcode is: FileSort" & FileSort & "CaseElse" & vbNewLine & vbNewLine _
                    & "Neem contact op met de software programeur: " & Sheets("Basisgeg.").Range("H26")
        End Select
    End If
End With

Exit Function

ErrorText:
If Err.Number <> 0 Then
    SeeText (SubName)
    End If
    Resume Next
    
End Function

Function GetFile(FileSort As String) As String

SubName = "'GetFile'"
If View("Errr") = True Then
    On Error GoTo ErrorText:
End If

Application.ScreenUpdating = View("Updte")
Application.DisplayAlerts = View("Alrt")

Dim fName As String

FilePath = ActiveWorkbook.path

With Application.FileDialog(msoFileDialogOpen)

    .InitialFileName = FilePath

    Select Case FileSort
        Case "Logo"
            .Filters.Clear
            .Filters.Add "JPEGS", "*.jpg; *.jpeg"
            .Filters.Add "GIF", "*.GIF"
            .Filters.Add "Bitmaps", "*.bmp"
    
            .AllowMultiSelect = False
    End Select

    If .Show <> -1 Then End

    GetFile = .SelectedItems(1)

End With

'Pad opslaan voor later gebruik
With Sheets("Basisgeg.")
    Select Case FileSort
        Case "Logo"
            .Range("C26").Value = GetFile
    End Select
End With

Exit Function

ErrorText:
If Err.Number <> 0 Then
    SeeText (SubName)
    End If
    Resume Next
    
End Function

Function InsertPictureInRange(PictureFileName As String, TargetCells As Range, TargetSheet As Worksheet)

SubName = "'InsertPictureInRange'"
If View("Errr") = True Then
    On Error GoTo ErrorText:
End If

Application.ScreenUpdating = View("Updte")
Application.DisplayAlerts = View("Alrt")

' inserts a picture and resizes it to fit the TargetCells range
Dim p As Object, t As Double, l As Double, w As Double, h As Double, b As Double, r As Double

1
If TypeName(ActiveSheet) <> "Worksheet" Then Exit Function

If ActiveSheet.Name <> TargetSheet.Name Then Exit Function

If Dir(PictureFileName) = "" Then Exit Function

10
    ' import picture
Set p = TargetSheet.Pictures.Insert(PictureFileName)

20
If TargetSheet.Name = "Factuur" Then
    If p.Height > 75 Then ScalePicH = True 'check if logo is to big
    If p.Width > 528.75 Then ScalPicW = True
    p.Name = "Bedrijfslogo"
End If

PicHoogte = p.Height

30
    ' determine positions
With TargetCells
    't = .Top
    'l = .Left
    b = .Offset(1, 1).Top
    r = .Offset(0, 1).Left
    
    t = b - p.Height
    l = r - p.Width
    
    w = p.Width
    h = p.Height
End With

40
If ScalePicH Then
    t = 0
    h = 75
    w = p.Width - (p.Height - 75)
    l = r - w
End If

50
If ScalePicW Then
    MsgBox "De afbeelding mag niet langer zijn dan 88,5cm."
    Sheets("Basisgeg.").Range("C26").ClearContents
End If

60
    ' position picture
With p
    .Top = t
    .Left = l
    .Width = w
    .Height = h
End With

70
Set p = Nothing

Exit Function

ErrorText:
If Err.Number <> 0 Then
    SeeText (SubName)
    End If
    Resume Next

End Function

Function FormuleProvider(WhatFormule As String) As String

SubName = "'FormuleProvider'"
If View("Errr") = True Then
    On Error GoTo ErrorText:
End If

Application.ScreenUpdating = View("Updte")
Application.DisplayAlerts = View("Alrt")

Select Case WhatFormule
    Case "FactuurVolgNr"
        FormuleProvider = "=IF(Factuurlijst!R2C1 < 9,""0"","""") &R[3]C& Factuurlijst!R2C1 +1"
        'FormuleProvider = "=IF(Factuurlijst!R[-7]C[-21] < 9,""0"","""") &R[3]C& Factuurlijst!R[-7]C[-21] +1"
    
    Case "Voorloop 1ste-0"
        FormuleProvider = "=IF(OR(Basisgeg.!R21C3='Factuur invoer'!R8C21,Basisgeg.!R21C3='Factuur invoer'!R7C21,Basisgeg.!R21C3='Factuur invoer'!R6C21,Basisgeg.!R21C3='Factuur invoer'!R3C21),Factuurlijst!R2C1*100,FALSE)"
    
    Case "Voorloop 2de-0"
        FormuleProvider = "=IF(OR(Basisgeg.!R21C4='Factuur invoer'!R8C21,Basisgeg.!R21C4='Factuur invoer'!R7C21,Basisgeg.!R21C4='Factuur invoer'!R6C21,Basisgeg.!R21C4='Factuur invoer'!R3C21),Factuurlijst!R2C1*100,FALSE)"
    
    Case "Voorloop 0en"
        FormuleProvider = "=CONCATENATE(IF(ISNUMBER(R[-2]C),RIGHT(R[-2]C,LEN(10-Factuurlijst!R2C1+1)),""""),IF(ISNUMBER(R[-1]C),RIGHT(R[-1]C,LEN(10-Factuurlijst!R2C1+1)),""""))"
    
    Case "FactuurNrInvoer"
        FormuleProvider = "=IF(R6C4="""","""",CONCATENATE(VLOOKUP(Basisgeg.!R21C3,'Factuur invoer'!R2C21:R8C22,2,FALSE),VLOOKUP(Basisgeg.!R21C4,'Factuur invoer'!R2C21:R8C22,2,FALSE),R9C22))"
    
    Case "FactuurNrLijst"
        FormuleProvider = "=IF(R3C1="""","""",IF(RC3="""","""",CONCATENATE(VLOOKUP(Basisgeg.!R21C3,R[12]C:R[18]C[1],2,FALSE),VLOOKUP(Basisgeg.!R21C4,R[12]C:R[18]C[1],2,FALSE),R[19]C[1])))"
        'FormuleProvider = "=IF(R3C1="""","""",IF(RC3="""","""",CONCATENATE(VLOOKUP(Basisgeg.!R21C3,R[12]C:R[18]C[1],2,FALSE),VLOOKUP(Basisgeg.!R21C4,R[12]C:R[18]C[1],2,FALSE),""0"",R2C1+1)))"
    
    
    Case Else
        MsgBox "Geen formule ingesteld"
        End
End Select

Exit Function

ErrorText:
If Err.Number <> 0 Then
    SeeText (SubName)
    End If
    Resume Next

End Function
