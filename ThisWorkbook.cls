VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
 
Private Sub Workbook_NewSheet(ByVal Sh As Object)
     
    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False

        ActiveSheet.Delete
         
        MsgBox "Kan geen nieuw blad toevoegen!", 16

    End With
Admin.ShowOneSheet ("Factuur invoer")

End Sub

Private Sub Workbook_Activate()
    
    Admin.ActivateWorkModus
    
    If Sheets("Basisgeg.").Range("A2").Value = "" Then Admin.ShowOneSheet ("Basisgeg.")
    
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)

    If BackgroundFunction.InArray("Modus", Sheets("Basisgeg.").Range("O1").Value) Then _
        Admin.ActivateWorkModus
    
    Admin.Afsluiten
End Sub
