Attribute VB_Name = "BackgroundFunction"
Function DocVersie() As String

SubName = "'DocVersie'"

Application.ScreenUpdating = View("Updte")
Application.DisplayAlerts = View("Alrt")

Error.DebugTekst Tekst:="Start", FunctionName:=SubName
'----Start

On Error GoTo FixVersion

DocName = ThisWorkbook.Name

vStart = InStr(DocName, "-v") + 2
vStop = InStr(DocName, ".xl")

DocVersie = Mid(DocName, vStart, vStop - vStart)

Error.DebugTekst Tekst:="Finish - DocVersie: " & DocVersie, FunctionName:=SubName

'---Finish
Exit Function
FixVersion:

DocVersie = "1-1"
Error.DebugTekst "Fixed:" & DocVersie, SubName

Exit Function

ErrorText:
If Err.Number <> 0 Then SeeText (SubName)
 Resume Next

End Function

Function InArray(WitchArray, strValue, Optional ArrayList)

SubName = "'InArray'"
If View("Errr") = True Then On Error GoTo ErrorText:

Application.ScreenUpdating = View("Updte")
Application.DisplayAlerts = View("Alrt")

Error.DebugTekst Tekst:="Start input:" & vbNewLine _
                        & "--- WithcArray: " & WitchArray & vbNewLine _
                        & "--- strValue: " & strValue & vbNewLine _
                        & "--- ArrayList: Optional", _
                    FunctionName:=SubName
'----Start

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
Case "Doornummeren"
    ArrayVals = Array("Maand", "Onderneming", "Afk. onderneming", "Niets", "")
End Select

20  For j = 0 To UBound(ArrayVals)
21    If ArrayVals(j) = CStr(strValue) Then
22      InArrayOut = True
        GoTo EndFunction
    End If
  Next
25  InArrayOut = False

EndFunction:
'---Finish
Error.DebugTekst Tekst:="Finish: " & InArrayOut, FunctionName:=SubName

InArray = InArrayOut

Exit Function

ErrorText:
If Err.Number <> 0 Then SeeText (SubName)
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

Function FactuurCheck(ByRef FactuurNr) As Boolean

SubName = "'FactuurCheck'"
If View("Errr") = True Then On Error GoTo ErrorText:

Application.ScreenUpdating = View("Updte")
Application.DisplayAlerts = View("Alrt")

Error.DebugTekst Tekst:="Start input: " & vbNewLine _
                        & "FactuurNr" & FactuurNr, _
                        FunctionName:=SubName
'----Start

Dim rng As Range

If IsEmpty(FactuurNr) Or FactuurNr = "" Then
    FactuurCheckOut = False
    GoTo EndFunction
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
    FactuurCheckOut = True
    
Else
    FactuurCheckOut = False
End If

EndFunction:
'---Finish
Error.DebugTekst Tekst:="FactuurCheck = " & FactuurCheckOut, FunctionName:=SubName

FactuurCheck = FactuurCheckOut

Exit Function

ErrorText:
If Err.Number <> 0 Then SeeText (SubName)
 Resume Next
    
End Function

Function GetFolder(strPath As String, FileSort As String) As String

SubName = "'GetFolder'"
If View("Errr") = True Then On Error GoTo ErrorText:

Application.ScreenUpdating = View("Updte")
Application.DisplayAlerts = View("Alrt")

Error.DebugTekst Tekst:="Start input: " & vbNewLine _
                        & "strPath: " & strPath & vbNewLine _
                        & "FileSort: " & FileSort, _
                        FunctionName:=SubName
'----Start

Dim fldr As FileDialog
Dim sItem As String
Dim ProblemTxt As String
Dim GetFolderOut As String

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

GetFolderOut = sItem

'Pad opslaan voor later gebruik?
SavePath = MsgBox("Pad opslaan voor later gebruik?", vbYesNo, "Pad opslaan?")

With Sheets("Basisgeg.")
    If SavePath = 6 Then
        SameBaseFolder = InStr(GetFolderOut, Application.ActiveWorkbook.path)
        If SameBaseFolder = 1 Then
            GetFolderOut = Mid(GetFolderOut, Len(Application.ActiveWorkbook.path) - 1, Len(GetFolderOut) - Len(Application.ActiveWorkbook.path))
        End If
        Select Case FileSort
            Case "PDF"
                .Range("C25").Value = GetFolderOut
            Case "Backup"
                .Range("C24").Value = GetFolderOut
            Case Else
                ProblemTxt = "Er is iets fout gegaan bij het verwerken van de folder gegevens." & vbNewLine _
                    & "De foutcode is: FileSort-" & FileSort & "-CaseElse" & vbNewLine & vbNewLine _
                    & "Neem contact op met de software programeur: " & Sheets("Basisgeg.").Range("H9")
                'Error.FunctionProblem FunctionName:="SaveFolder", _
                    Problem:=ProblemTxt
                Error.DebugTekst Tekst:=ProblemTxt, FunctionName:=SubName
        End Select
    End If
End With

Set fldr = Nothing

'Check if folder exits
If GetFolderOut <> "" Then
    If BackgroundFunction.TestFolderExist(GetFolderOut) = False Then
        ProblemTxt = "De folder: " & GetFolderOut & " bestaad niet." & vbNewLine _
                        & "Controleer of de sub-folders wel bestaan of selecteer een nieuwe map."
        Error.DebugTekst Tekst:=ProblemTxt, FunctionName:=SubName
        'SelectNewFolder = MsgBox("Een nieuwe map selecteren?", vbYesNo, "Nieuwe map selecteren")
        'If SelectNewFolder = vbYes Then BackgroundFunction.GetFolder
    End If
End If
    


'---Finish
Error.DebugTekst Tekst:="Finish: " & GetFolderOut, FunctionName:=SubName

GetFolder = GetFolderOut

Exit Function

ErrorText:
If Err.Number <> 0 Then SeeText (SubName)
 Resume Next
 
End Function

Function GetFile(FileSort As String) As String

SubName = "'GetFile'"
If View("Errr") = True Then On Error GoTo ErrorText:

Application.ScreenUpdating = View("Updte")
Application.DisplayAlerts = View("Alrt")

Error.DebugTekst Tekst:="Start input: " & vbNewLine _
                        & "FileSort: " & FileSort, _
                        FunctionName:=SubName
'----Start

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
    
    GetFileOut = .SelectedItems(1)
    Error.DebugTekst "Selected File: " & GetFileOut
    Error.DebugTekst "Application map: " & Application.ActiveWorkbook.path
    If InStr(GetFileOut, Application.ActiveWorkbook.path) Then _
        GetFileOut = Mid(GetFileOut, Len(Application.ActiveWorkbook.path) + 1, Len(GetFileOut) - Len(Application.ActiveWorkbook.path))
    
    Error.DebugTekst "Get File: " & GetFileOut

End With

'Pad opslaan voor later gebruik
With Sheets("Basisgeg.")
    Select Case FileSort
        Case "Logo"
            .Range("C26").Value = GetFileOut
    End Select
End With

'---Finish
Error.DebugTekst Tekst:="Finish: " & GetFileOut, FunctionName:=SubName
   
GetFile = GetFileOut

Exit Function

ErrorText:
If Err.Number <> 0 Then SeeText (SubName)
 Resume Next
    
End Function

Function InsertPictureInRange(PictureFileName As String, TargetCells As Range, TargetSheet As Worksheet) As Boolean

SubName = "'InsertPictureInRange'"
If View("Errr") = True Then On Error GoTo ErrorText:

Application.ScreenUpdating = View("Updte")
Application.DisplayAlerts = View("Alrt")

Error.DebugTekst Tekst:="Start input: " & vbNewLine _
                        & "--- PictureFileName: " & PictureFileName & vbNewLine _
                        & "--- TargetCells: " & TargetCells.Address & vbNewLine _
                        & "--- TargetSheet: " & TargetSheet.Name & vbNewLine, _
                        FunctionName:=SubName
'----Start

Admin.Bewerkbaar (TargetSheet.Name)

' inserts a picture and resizes it to fit the TargetCells range
Dim p As Object, t As Double, l As Double, w As Double, h As Double, b As Double, r As Double

1
If TypeName(ActiveSheet) <> "Worksheet" Then Exit Function

If ActiveSheet.Name <> TargetSheet.Name Then Exit Function

If Left(PictureFileName, 1) = "\" Then PictureFileName = Application.ActiveWorkbook.path & PictureFileName
If Dir(PictureFileName) = "" Then Exit Function
Error.DebugTekst "PictureFileName = " & PictureFileName
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
    InsertPictureInRange = False
    Exit Function
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

Admin.NietBewerkbaar (TargetSheet.Name)

'---Finish
Error.DebugTekst Tekst:="Finish", FunctionName:=SubName
InsertPictureInRange = True
Exit Function

ErrorText:
If Err.Number <> 0 Then SeeText (SubName)
 Resume Next

End Function

Function FormuleProvider(WhatFormule As String) As String

SubName = "'FormuleProvider'"
If View("Errr") = True Then On Error GoTo ErrorText:

Application.ScreenUpdating = View("Updte")
Application.DisplayAlerts = View("Alrt")

Error.DebugTekst Tekst:="Start input: " & vbNewLine _
                        & "WhatFormule: " & WhatFormule, _
                        FunctionName:=SubName
'----Start

Dim FormuleProviderOut As String

Select Case WhatFormule
    Case "FactuurVolgNr"
        ActSht = ActiveSheet.Name
        Admin.ShowOneSheet ("Factuurlijst")
        With Sheets("Factuurlijst")
            MonthCounterEnd = .Range("G14000").End(xlUp).Row
        End With
        Admin.ShowOneSheet (ActSht)
        FormuleProviderOut = "=IF(R[3]C="""",IF(VLOOKUP(MONTH('Factuur invoer'!R3C8),Factuurlijst!R" & MonthCounterEnd - 12 & "C7:R" & MonthCounterEnd & _
                                "C8,2,FALSE)=0,""01"",IF(VLOOKUP(MONTH('Factuur invoer'!R3C8),Factuurlijst!R" & MonthCounterEnd - 12 & "C7:R" & MonthCounterEnd & _
                                "C8,2,FALSE)<9,""0"","""")&VLOOKUP(MONTH('Factuur invoer'!R3C8),Factuurlijst!R" & MonthCounterEnd - 12 & "C7:R" & MonthCounterEnd & _
                                "C8,2,FALSE)+1),R[3]C&IF(VLOOKUP(MONTH('Factuur invoer'!R3C8),Factuurlijst!R" & MonthCounterEnd - 12 & "C7:R" & MonthCounterEnd & _
                                "C8,2,FALSE)<9,""0"","""")&VLOOKUP(MONTH('Factuur invoer'!R3C8),Factuurlijst!R" & MonthCounterEnd - 12 & "C7:R" & MonthCounterEnd & "C8,2,FALSE)+1)"
        'FormuleProviderOut = "=IF(R[3]C="""",IF(MONTH('Factuur invoer'!R1C20)=MONTH(Factuurlijst!R2C3),IF(Factuurlijst!R2C1<9,""0"","""")&Factuurlijst!R2C1+1,""01""),R[3]C&IF(Factuurlijst!R2C1<9,""0"","""")&Factuurlijst!R2C1+1)"
        'FormuleProviderOut = "=IF(Factuurlijst!R2C1 < 9,""0"","""") &R[3]C& Factuurlijst!R2C1 +1"
        'FormuleProviderOut = "=IF(Factuurlijst!R[-7]C[-21] < 9,""0"","""") &R[3]C& Factuurlijst!R[-7]C[-21] +1"
    
    Case "Voorloop 1ste-0"
        FormuleProviderOut = "=IF(OR(Basisgeg.!R21C3='Factuur invoer'!R8C21,Basisgeg.!R21C3='Factuur invoer'!R7C21,Basisgeg.!R21C3='Factuur invoer'!R6C21,Basisgeg.!R21C3='Factuur invoer'!R3C21),Factuurlijst!R2C1*100,FALSE)"
    
    Case "Voorloop 2de-0"
        FormuleProviderOut = "=IF(OR(Basisgeg.!R21C4='Factuur invoer'!R8C21,Basisgeg.!R21C4='Factuur invoer'!R7C21,Basisgeg.!R21C4='Factuur invoer'!R6C21,Basisgeg.!R21C4='Factuur invoer'!R3C21),Factuurlijst!R2C1*100,FALSE)"
    
    Case "Voorloop 0en"
        FormuleProviderOut = "=CONCATENATE(IF(ISNUMBER(R[-2]C),RIGHT(R[-2]C,LEN(10-Factuurlijst!R2C1+1)),""""),IF(ISNUMBER(R[-1]C),RIGHT(R[-1]C,LEN(10-Factuurlijst!R2C1+1)),""""))"
    
    Case "FactuurNrInvoer"
        FormuleProviderOut = "=IF(ISBLANK(R2C9),IF(R6C4="""","""",CONCATENATE(VLOOKUP(Basisgeg.!R21C3,'Factuur invoer'!R2C21:R8C22,2,FALSE),VLOOKUP(Basisgeg.!R21C4,'Factuur invoer'!R2C21:R8C22,2,FALSE),R9C22)),R2C9)"
        'FormuleProviderOut = "=IF(R6C4="""","""",CONCATENATE(VLOOKUP(Basisgeg.!R21C3,'Factuur invoer'!R2C21:R8C22,2,FALSE),VLOOKUP(Basisgeg.!R21C4,'Factuur invoer'!R2C21:R8C22,2,FALSE),R9C22))"
    
    Case "FactuurNrLijst"
        FormuleProviderOut = "=IF(R3C1="""","""",IF(RC3="""","""",CONCATENATE(VLOOKUP(Basisgeg.!R21C3,R[12]C:R[18]C[1],2,FALSE),VLOOKUP(Basisgeg.!R21C4,R[12]C:R[18]C[1],2,FALSE),R[19]C[1])))"
        'FormuleProviderOut = "=IF(R3C1="""","""",IF(RC3="""","""",CONCATENATE(VLOOKUP(Basisgeg.!R21C3,R[12]C:R[18]C[1],2,FALSE),VLOOKUP(Basisgeg.!R21C4,R[12]C:R[18]C[1],2,FALSE),""0"",R2C1+1)))"
    
    Case "PC_Plaats"
        FormuleProviderOut = "=IF(ISBLANK(R[-4]C[17]),"""",IF(VLOOKUP(R2C4,Debiteuren!R[-2]C[-4]:R[992]C[7],7,FALSE)="""","""",CONCATENATE(VLOOKUP(R[-3]C[-1],Debiteuren!R[-2]C[-4]:R[992]C[7],6,FALSE),"", "",VLOOKUP(R[-3]C[-1],Debiteuren!R[-2]C[-4]:R[992]C[7],7,FALSE))))"
    
    Case "PC_Plaats1"
        FormuleProviderOut = "=IF(RC[1]="""","""",IF(SEARCH("","",RC[1])=1,RIGHT(RC[1],LEN(RC[1])-2),RC[1]))"
    
    Case "Adres"
        FormuleProviderOut = "=IF(ISBLANK(R[-3]C[18]),"""",IF(VLOOKUP(R[-2]C,Debiteuren!R[-1]C[-3]:R[993]C[5],5,FALSE)="""","""",VLOOKUP(R[-2]C,Debiteuren!R[-1]C[-3]:R[993]C[5],5,FALSE)))"
    
    Case "Naam"
        FormuleProviderOut = "=IF(ISBLANK(R[-2]C[18]),"""",IF(VLOOKUP(R[-1]C,Debiteuren!RC[-3]:R[994]C[5],2,FALSE)="""","""",VLOOKUP(R[-1]C,Debiteuren!RC[-3]:R[994]C[5],2,FALSE)))"
        
    Case "LandNm"
        FormuleProviderOut = "=IF(ISERROR(VLOOKUP(R[-2]C[-2],Debiteuren!R[-1]C[-5]:R[993]C[6],8,FALSE)),"""",IF(VLOOKUP(R[-2]C[-2],Debiteuren!R[-1]C[-5]:R[993]C[6],8,FALSE)=""Nederland"","""",""Land""))"
    
    Case "EmailNm"
        FormuleProviderOut = "=IF(ISERROR(VLOOKUP(R[-3]C[-2],Debiteuren!R[-2]C[-5]:R[992]C[6],9,FALSE)),"""",IF(VLOOKUP(R[-3]C[-2],Debiteuren!R[-2]C[-5]:R[992]C[6],9,FALSE)="""","""",""Email""))"
    
    Case "TelefoonNm"
        FormuleProviderOut = "=IF(ISERROR(VLOOKUP(R[-4]C[-2],Debiteuren!R[-3]C[-5]:R[991]C[6],10,FALSE)),"""",IF(VLOOKUP(R[-4]C[-2],Debiteuren!R[-3]C[-5]:R[991]C[6],10,FALSE)="""","""",""Telefoon""))"
    
    Case "Land"
        FormuleProviderOut = "=IF(ISERROR(VLOOKUP(R[-2]C[-3],Debiteuren!R[-1]C[-6]:R[993]C[5],8,FALSE)),"""",IF(VLOOKUP(R[-2]C[-3],Debiteuren!R[-1]C[-6]:R[993]C[5],8,FALSE)=""Nederland"","""",UPPER(VLOOKUP(R[-2]C[-3],Debiteuren!R[-1]C[-6]:R[993]C[5],8,FALSE))))"
    
    Case "Email"
        FormuleProviderOut = "=IF(ISERROR(VLOOKUP(R[-3]C[-3],Debiteuren!R[-2]C[-6]:R[992]C[5],9,FALSE)),"""",IF(VLOOKUP(R[-3]C[-3],Debiteuren!R[-2]C[-6]:R[992]C[5],9,FALSE)="""","""",VLOOKUP(R[-3]C[-3],Debiteuren!R[-2]C[-6]:R[992]C[5],9,FALSE)))"
    
    Case "Telefoon"
        FormuleProviderOut = "=IF(ISERROR(VLOOKUP(R[-4]C[-3],Debiteuren!R[-3]C[-6]:R[991]C[5],10,FALSE)),"""",IF(VLOOKUP(R[-4]C[-3],Debiteuren!R[-3]C[-6]:R[991]C[5],10,FALSE)="""","""",VLOOKUP(R[-4]C[-3],Debiteuren!R[-3]C[-6]:R[991]C[5],10,FALSE)))"
    
    Case "OpmerkingNm"
        FormuleProviderOut = "=IF(ISERROR(VLOOKUP(R[-1]C[-7],Debiteuren!RC[-10]:R[994]C[1],11,FALSE)),"""",IF(VLOOKUP(R[-1]C[-7],Debiteuren!RC[-10]:R[994]C[1],11,FALSE)="""","""",""Opmerking:""))"
    
    Case "Opmerking"
        FormuleProviderOut = "=IF(ISERROR(VLOOKUP(R[-2]C[-7],Debiteuren!R[-1]C[-10]:R[993]C[1],11,FALSE)),"""",IF(VLOOKUP(R[-2]C[-7],Debiteuren!R[-1]C[-10]:R[993]C[1],11,FALSE)="""","""",VLOOKUP(R[-2]C[-7],Debiteuren!R[-1]C[-10]:R[993]C[1],11,FALSE)))"

    Case Else
        MsgBox "Geen formule ingesteld"
        End
End Select

EndFunction:

'---Finish
Error.DebugTekst Tekst:="Finish FormuleProvider: " & FormuleProviderOut, FunctionName:=SubName

FormuleProvider = FormuleProviderOut

Exit Function

ErrorText:
If Err.Number <> 0 Then SeeText (SubName)
 Resume Next

End Function

Function IsFile(fName As String) As Boolean

SubName = "'IsFile'"

Application.ScreenUpdating = View("Updte")
Application.DisplayAlerts = View("Alrt")

Error.DebugTekst Tekst:="Start input: " & vbNewLine _
                        & "fName: " & fName, _
                        FunctionName:=SubName
'----Start

'Returns TRUE if the provided name points to an existing file.
'Returns FALSE if not existing, or if it's a folder
    On Error Resume Next
    IsFileOut = ((GetAttr(fName) And vbDirectory) <> vbDirectory)

'---Finish
Error.DebugTekst Tekst:="Finish: " & IsFileOut, FunctionName:=SubName

IsFile = IsFileOut

Exit Function

End Function

Function TestFolderExist(FolderPath As String) As Boolean

SubName = "'TestFolderExist'"
If View("Errr") = True Then On Error GoTo ErrorText:

Application.ScreenUpdating = View("Updte")
Application.DisplayAlerts = View("Alrt")

Error.DebugTekst Tekst:="Start input: " & vbNewLine _
                        & "FolderPath: " & FolderPath, _
                        FunctionName:=SubName
'----Start
Dim TestStr As String

If Right(FolderPath, 1) <> "\" Then FolderPath = FolderPath & "\"
If Left(FolderPath, 1) = "\" Then FolderPath = Application.ActiveWorkbook.path & FolderPath

Error.DebugTekst ("After changes folderpath= " & FolderPath)

TestStr = ""
On Error Resume Next
TestStr = Dir(FolderPath)

If View("Errr") = True Then On Error GoTo ErrorText:

If TestStr = "" Then
    TestFolderExistOut = False
Else
    TestFolderExistOut = True
End If

'---Finish
Error.DebugTekst Tekst:="Finish with: " & TestFolderExistOut, FunctionName:=SubName

TestFolderExist = TestFolderExistOut

Exit Function

ErrorText:
If Err.Number <> 0 Then SeeText (SubName)
 Resume Next
End Function
