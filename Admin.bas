Attribute VB_Name = "Admin"

Sub ShowAllSheets()

ActSht = ActiveSheet.Name

For i = 1 To Sheets.count
        Sheets(i).Visible = xlSheetVisible
Next i

Sheets(ActSht).Select

End Sub
Sub ShowOneSheet(Sht As String)

ActSht = ActiveSheet.Name

If ActSht <> Sht Then
    Sheets(Sht).Visible = xlSheetVisible
    Sheets(Sht).Select
    If BackgroundFunction.InArray("Modus", Sheets("Basisgeg.").Range("O1").Value) = False Then _
        Sheets(ActSht).Visible = xlSheetHidden
End If

For i = 1 To Sheets.count
    If Sheets(i).Name <> Sht Then
        If BackgroundFunction.InArray("Modus", Sheets("Basisgeg.").Range("O1").Value) = False Then _
            Sheets(i).Visible = xlSheetHidden
    End If
Next i

End Sub

Sub HideAllSheets()

If BackgroundFunction.InArray("Modus", Sheets("Basisgeg.").Range("O1").Value) Then Exit Sub

ActSht = "Factuur invoer"

For i = 1 To Sheets.count
    If ActSht <> Sheets(i).Name Then
        If Sheets(i).Visible <> xlSheetHidden Then
            Sheets(i).Visible = xlSheetHidden
        End If
    End If
Next i

Sheets(ActSht).Range("D6").Select

End Sub

Sub ActivateWorkModus()

Admin.ShowAllSheets

With Sheets("Basisgeg.")
    'Verplichte velden checken
    If .Range("B8").Value = "" And .Range("B2").Value = "" And _
        .Range("B5").Value = "" And .Range("B6").Value = "" Then
            MsgBox "Bestand wordt opgeslagen als 'LEEG'"
            .Range("O1").Value = "Leeg"
    End If
End With

For i = 1 To Sheets.count
    Admin.NietBewerkbaar (Sheets(i).Name)
    
    Sheets(i).Select
    If BackgroundFunction.InArray("Database", Sheets(i).Name) = False Then _
        ActiveWindow.DisplayGridlines = False 'Rasterlijnen
    
    With Sheets(i)
        Select Case .Name
            Case "Basisgeg."
                .Range("O1").Value = "Work modus"
                .ScrollArea = "A1:H100"
            Case "Boekingslijst"
                .ScrollArea = "A1:N10000"
            Case "Factuurlijst"
                .ScrollArea = "A1:CE10000"
            Case "Factuur invoer"
                .ScrollArea = "A1:P39"
            Case "Factuur"
                .ScrollArea = "A1:S53"
            Case "Artikelen"
                .ScrollArea = "A1:G10000"
            Case "Debiteuren"
                .ScrollArea = "A1:K10000"
            Case "Maandoverzicht"
                .ScrollArea = "A1:M22"
            Case "Kwartaaloverzicht"
                .ScrollArea = "A1:M22"
            Case "Jaaroverzicht"
                .ScrollArea = "A1:M32"
            Case "Afdruk boekingen"
                .ScrollArea = "A1:O10000"
        End Select
    End With
    
    With ActiveWindow
        .DisplayHeadings = False 'Kolom en rij koppen
    End With
Next i

With ActiveWindow
    .DisplayWorkbookTabs = False 'Werkblad tabs
    .DisplayHorizontalScrollBar = False 'Horizontaal scrollen uit
    .DisplayVerticalScrollBar = False 'Verticaal scrollen uit
End With
With Application
    .DisplayFormulaBar = False
    .DisplayFullScreen = True 'Volledig scherm
    .DisplayFormulaBar = False
End With

Admin.HideAllSheets

End Sub

Sub DeActivateWorkModus()

Admin.ShowAllSheets

For i = 1 To Sheets.count
    Admin.Bewerkbaar (Sheets(i).Name)
    
    Sheets(i).Select
    With ActiveWindow
        .DisplayGridlines = True 'Rasterlijnen
        .DisplayHeadings = True 'Kolom en rij koppen
        .DisplayWorkbookTabs = True 'Werkblad tabs
        .DisplayHorizontalScrollBar = True
        .DisplayVerticalScrollBar = True
    End With
    
    Sheets(i).ScrollArea = ""
    
Next i

With Application
    .DisplayFormulaBar = True
    .DisplayFullScreen = False 'Volledig scherm
    .DisplayFormulaBar = True
End With

Admin.ShowAllSheets

Beveiliging = MsgBox("Schrijfbeveiliging op de bladen aan zetten?", vbYesNo, "Schrijfbeveiliging?")
If Beveiliging = vbYes Then
    Sheets("Basisgeg.").Select
    Sheets("Basisgeg.").Range("O1").Value = "Test modus beveiligd"
    
    For i = 1 To Sheets.count
        Admin.NietBewerkbaar (Sheets(i).Name)
    Next i
Else
    Sheets("Basisgeg.").Select
    Sheets("Basisgeg.").Range("O1").Value = "Test modus"
End If

End Sub

Sub HideOneSheet(Sht As String)

ActSht = ActiveSheet.Name

If ActSht = Sht Or ActSht = "" Then ActSht = "Basisgeg."

    Sheets(ActSht).Visible = xlSheetVisible
    Sheets(ActSht).Select
    If BackgroundFunction.InArray("Modus", Sheets("Basisgeg.").Range("O1").Value) = False Then _
        Sheets(Sht).Visible = xlSheetHidden

End Sub

Function Bewerkbaar(Sht As String) As Boolean

SubName = "'Bewerkbaar'"
If View("Errr") = True Then
    On Error GoTo ErrorText:
End If

Application.ScreenUpdating = View("Updte")
Application.DisplayAlerts = View("Alrt")

ThisSheet = ActiveSheet.Name
Sheets(Sht).Select
ActiveSheet.Unprotect Password:=PassWordChanger()
Sheets(ThisSheet).Select

Bewerkbaar = True
Exit Function

ErrorText:
If Err.Number <> 0 Then
    SeeText (SubName)
    End If
    Resume Next
    
End Function

Function NietBewerkbaar(Sht As String) As Boolean

SubName = "'NietBewerkbaar'"
If View("Errr") = True Then
    On Error GoTo ErrorText:
End If

Application.ScreenUpdating = View("Updte")
Application.DisplayAlerts = View("Alrt")

If BackgroundFunction.InArray("ModusBeveiliging", Sheets("Basisgeg.").Range("O1").Value) Then Exit Function

10 ThisSheet = ActiveSheet.Name

100 Sheets(Sht).Select

111     ActiveSheet.Protect Password:=PassWordChanger()

900 Sheets(ThisSheet).Select

999 NietBewerkbaar = True

Exit Function

ErrorText:
If Err.Number <> 0 Then
    SeeText (SubName)
    End If
    Resume Next

End Function

Private Function PassWordChanger() As String

PassWordChanger = "Freedom1945"

End Function

Sub Afsluiten()

With ThisWorkbook
    .Save
End With

Application.Quit

End Sub

Public Sub ExportVisualBasicCode()

SubName = "'ExportVisualBasicCode'"
If View("Errr") = True Then
    On Error GoTo ErrorText:
End If

Application.ScreenUpdating = View("Updte")
Application.DisplayAlerts = View("Alrt")

'Error.DebugTekst Tekst:="Start", FunctionName:=SubName

' Excel macro to export all VBA source code in this project to text files for proper source control versioning
' Requires enabling the Excel setting in Options/Trust Center/Trust Center Settings/Macro Settings/Trust access to the VBA project object model
    Const Module = 1
    Const ClassModule = 2
    Const Form = 3
    Const Document = 100
    Const Padding = 24
    
    Dim VBComp
    Dim count As Integer
    Dim path As String
    Dim directory As String
    Dim extension As String
    Dim fso As Object
    
    'dirStart = ActiveWorkbook.path
    dirStart = "C:\Users\Gebruiker\Documents\GitHub\ZwerfBoekExcel" 'starting directory
    directory = "\" 'new directories for the vba-scripts
    fName = "Boekhoud-v1-0.xlsm" 'this filename
    'path = dirStart & directory
    
    Set fso = CreateObject("scripting.filesystemobject")
    count = 0
    skiped = 0

    If Not fso.FolderExists(dirStart & directory) Then
        'when directory does not exists, make path
        newDir = dirStart
        Folders = Split(directory, "\")
        For i = 0 To UBound(Folders)
            newDir = fso.BuildPath(newDir, Folders(i))
            If fso.FolderExists(newDir) Then
                Set objFolder = fso.GetFolder(newDir)
            Else
                Set objFolder = fso.CreateFolder(newDir)
                Debug.Print "Create folder: " & newDir
            End If
        Next
    End If
    Set fso = Nothing
    'Check if the right workbook is active
    If ActiveWorkbook.Name <> fName Then Workbooks(fName).Activate
    
    For Each VBComponent In ActiveWorkbook.VBProject.VBComponents
        Select Case VBComponent.Type
            Case ClassModule, Document
                extension = ".cls"
            Case Form
                extension = ".frm"
            Case Module
                extension = ".bas"
            Case Else
                extension = ".txt"
        End Select
            
                
        On Error Resume Next
        Err.Clear
        
        Bladcheck = InStr(VBComponent.Name, "Blad")
        If Bladcheck > 0 Then
            Debug.Print ("Skiped: " & VBComponent.Name)
            skiped = skiped + 1
            GoTo Volgende
        End If
        
        If count = 0 Then directory = dirStart & directory
        path = directory & "\" & VBComponent.Name & extension
        VBComponent.Export (path)
        If Err.Number <> 0 Then
            If InArray("VBAExport", extension) Then
                MsgBox _
                        Prompt:="Failed to export: " & VBComponent.Name & vbNewLine _
                            & " to " & path & vbNewLine _
                            & vbNewLine & "Errornr: " & Err.Number & vbNewLine _
                            & "Description:  " & Err.Description, _
                        Title:="Failed to export"
                        
            End If
        Else
            count = count + 1
            Debug.Print "Exported " & Left$(VBComponent.Name & ":" & Space(Padding), Padding) & path
        End If
        
Volgende:
        On Error GoTo ErrorText
    Next
    
    Application.StatusBar = "Successfully exported: " & CStr(count) & " files | Skiped: " & CStr(skiped) & " files"
    Debug.Print "Successfully exported: " & CStr(count) & " files | Skiped: " & CStr(skiped) & " files"

Exit Sub

ErrorText:
If Err.Number <> 0 Then
    SeeText (SubName)
    End If
    Resume Next
   
End Sub

