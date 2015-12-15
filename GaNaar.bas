Attribute VB_Name = "GaNaar"
Sub Naarmaand()

    Admin.ShowOneSheet ("Maandoverzicht")
End Sub
Sub Naarjaar()

    Admin.ShowOneSheet ("Jaaroverzicht")
    Range("A1").Select
End Sub
Sub Afdrukboeken()

    Admin.ShowOneSheet ("Afdruk boekingen")
    Range("D3").Select
End Sub

Sub GaNaarArtikelen()
    
    Sht = "Artikelen"
    
    Admin.ShowOneSheet (Sht)
    Range("C4").Select
    
    Admin.Bewerkbaar (Sht)
    'Sheets(Sht).ShowDataForm
    'Admin.NietBewerkbaar (Sht)
    
    'Admin.ShowOneSheet ("Factuur invoer")
End Sub

Sub GaNaarDebiteuren()

    Sht = "Debiteuren"
    
    Admin.ShowOneSheet (Sht)
    Range("C4").Select
    
    'Admin.Bewerkbaar (Sht)
    'Sheets(Sht).ShowDataForm
    'Admin.NietBewerkbaar (Sht)
    
    'Admin.ShowOneSheet ("Factuur invoer")
End Sub

Sub bekijkfactuur()
    
    Application.ScreenUpdating = True
    
    Admin.ShowOneSheet ("Factuur")
    Sheets("Factuur").Range("B1").Select
    
    Application.ScreenUpdating = View("Updte")
    
End Sub

Sub Naarboeken()

    Admin.ShowOneSheet ("Boekingslijst")
    Range("C2").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Select
End Sub

Sub NaarFactuurInvoer()
    
    ThisSheet = ActiveSheet.Name
    If ThisSheet = "Debiteuren" Or "Artikelen" Then
        'Check of er nog vreemde input staat en die opruimen
        If NieuweInput.CheckNwInput(ActiveSheet.Name) = False Then
            MsgBox "Problem with CheckNwInput, system has to end"
            Admin.ShowOneSheet ("Factuur invoer")
            Exit Sub
        End If
    End If
    Admin.NietBewerkbaar (ThisSheet)
    
    Admin.ShowOneSheet ("Factuur invoer")
    
End Sub

Sub BTWaangifteOverzicht()
    
    If Sheets("Basisgeg.").Range("C22").Value = "" Then
        MsgBox "Er is geen aangifte termijn geselecteerd"
        Admin.ShowOneSheet ("Basisgeg.")
        Sheets("Basisgeg.").Range("C22").Select
        End
    End If
        
    Sht = Sheets("Basisgeg.").Range("C22").Value & "overzicht"

    Admin.ShowOneSheet (Sht)

End Sub

Sub NaarBasisgegevens()

    Admin.ShowOneSheet ("Basisgeg.")

End Sub
