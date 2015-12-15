Attribute VB_Name = "Error"
Function LogFileName() As String

TempFolder = Environ("Temp")
ErrorFile = "ErrorLog-Boekhoud-v1-2_" & Format(Now(), "ddmmyy") & ".txt"
LogFileNameOut = TempFolder & "\" & ErrorFile

'Debug.Print "----" & vbNewLine & "LogFileName: " & LogFileNameOut & vbNewLine & "----"
LogFileName = LogFileNameOut

End Function
Function SendCDOmail(ByVal eTo As String, ByVal eFrom As String, _
                        ByVal eSubject As String, ByVal eBody As String, _
                        Optional ByVal eCopy As String, Optional ByVal eBCC As String, _
                            Optional ByVal eAttach As String)

SubName = "'SendCDOmail'"
If View("Errr") = True Then On Error GoTo ErrorText:

Application.ScreenUpdating = View("Updte")
Application.DisplayAlerts = View("Alrt")
Error.DebugTekst Tekst:="Start", FunctionName:=SubName
'--------Start Function

Dim iMsg As Object
Dim iConf As Object

40
Set iMsg = CreateObject("CDO.Message")
Set iConf = CreateObject("CDO.Configuration")

If IsEmpty(eCopy) Then eCopy = ""
If IsEmpty(eBCC) Then eBCC = ""
If IsEmpty(eAttach) Then eAttach = ""

Error.DebugTekst Tekst:="Input values: " & vbNewLine _
                    & "-->To: " & eTo & vbNewLine _
                    & "-->CC: " & eCopy & vbNewLine _
                    & "-->BCC: " & eBCC & vbNewLine _
                    & "-->From: " & eFrom & vbNewLine _
                    & "-->Subject: " & eSubject & vbNewLine _
                    & "-->Body: " & eBody & vbNewLine _
                    & "-->Attachment: " & eAttach, _
                    FunctionName:=SubName

iConf.Load -1    ' CDO Source Defaults
Set Flds = iConf.Fields
With Flds
    '.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
    '.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
    '.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 2525
    '.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = False
    '.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "mx1.hostfree.nl"
    '.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
    '.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "webbeheerder@lieskebethke.nl"
    '.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "9aveMxFY"

    '------Gmail Gegevens
    .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
    .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
    .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 465
    .Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
    .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.gmail.com"
    .Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
    .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "webbeheerder@zwervers.org"
    .Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "Zw.orgwb85"
    eFrom = "webbeheerder@zwervers.org"
    '------Gmail Gegevens
    
    .Update
End With

70
With iMsg
    Set .Configuration = iConf
    .To = eTo
    .CC = eCopy
    .BCC = eBCC
    .From = eFrom
    .Subject = eSubject
    .HTMLBody = eBody
    .AddAttachment eAttach
    .Send
End With


'--------End Function

Error.DebugTekst Tekst:="Finished", FunctionName:=SubName

OCDinfo = FnctionEnd
Exit Function

ErrorText:
If Err.Number <> 0 Then
    Debug.Print Err.Description
    SeeText (SubName)
End If
Resume Next

End Function


Function View(ByVal What As String) As Boolean

TestModus = False

Select Case TestModus
Case True

    Select Case What
    'View error message
    Case "Errr"
    View = False
    
    'View screen update
    Case "Updte"
    View = True
    
    'View Alerts
    Case "Alrt"
    View = True
    
    Case ""
    View = False
    
    MsgBox "View function is empty, now set to: " & View
    
    End Select

Case False

    Select Case What
    'View error message
    Case "Errr"
    View = True
    
    'View screen update
    Case "Updte"
    View = False
    
    'View Alerts
    Case "Alrt"
    View = False
    
    Case ""
    View = False
    
    MsgBox "View function is empty, now set to: " & View
    
    End Select
End Select
End Function

Function SeeText(SubName As String)

Dim Msg As String
Dim Counter As Integer
Dim versie As String
Dim TitleKlm, OutputKlm As String
Dim StartRow As Integer
Dim ErrorShtName As String
Dim ErrNumber, ErrLine As Integer
Dim ErrSource, ErrDescription As String

'first add the error information in a variable otherwise they are gone
ErrNumber = Err.Number
ErrSource = Err.Source
ErrLine = Erl
ErrDescription = Err.Description

1
    Msg = "Error # " & Str(ErrNumber) & Chr(13) _
            & SubName & " genarated a error." & Chr(13) _
            & "Source: " & ErrSource & Chr(13) _
            & "Error Line: " & ErrLine & Chr(13) _
            & "Description: " & ErrDescription
    
TitleKlm = "N"
OutputKlm = "O"
StartRow = 4
ErrorShtName = "Basisgeg."

    'Error.DebugTekst (Msg)
10    'notice the error in the "Bugfix indicator"
    ActiveSht = ActiveSheet.Name
    
    If ActiveSht <> ErrorShtName Then Admin.ShowOneSheet (ErrorShtName)
    If ActiveSheet.Name <> ErrorShtName Then
        Error.DebugTekst Tekst:=ErrorShtName & " can't get active. End Error saving", FunctionName:="SeeText"
        Exit Function
    End If
    If Admin.Bewerkbaar(ErrorShtName) = False Then
        Error.DebugTekst Tekst:=ErrorShtName & " stays protected. End Error saving", FunctionName:="SeeText"
        Exit Function
    End If
    
    With Sheets(ErrorShtName)
        versie = .Range(OutputKlm & StartRow).Value
        Counter = .Range(OutputKlm & StartRow + 1).Value
        If .Range(OutputKlm & StartRow + 1).Value = "" Then
            Counter = 1
            .Range(OutputKlm & StartRow + 1).Value = Counter
            .Range(OutputKlm & StartRow + 2).Value = SubName
            .Range(OutputKlm & StartRow + 3).Value = ErrLine
            .Range(OutputKlm & StartRow + 4).Value = ErrNumber
            .Range(OutputKlm & StartRow + 5).Value = ErrSource
            .Range(OutputKlm & StartRow + 6).Value = ErrDescription
            Error.DebugTekst "New error in Bugfix indicator"
            
11      ElseIf .Range(OutputKlm & StartRow + 1).Value > 0 Then
            If SubName = "'" & .Range(OutputKlm & StartRow + 2).Value And _
                ErrLine = .Range(OutputKlm & StartRow + 3).Value And _
                ErrNumber = .Range(OutputKlm & StartRow + 4).Value And _
                ErrSource = .Range(OutputKlm & StartRow + 5).Value Then
                    .Range(OutputKlm & StartRow + 1).Value = Counter + 1
                    If .Range(OutputKlm & StartRow + 1).Value > 100 Then End
12          Else
                Counter = 1
                .Range(OutputKlm & StartRow + 1).Value = Counter
                .Range(OutputKlm & StartRow + 2).Value = SubName
                .Range(OutputKlm & StartRow + 3).Value = ErrLine
                .Range(OutputKlm & StartRow + 4).Value = ErrNumber
                .Range(OutputKlm & StartRow + 5).Value = ErrSource
                .Range(OutputKlm & StartRow + 6).Value = ErrDescription
                Error.DebugTekst "Delete previous and add new error in Bugfix indicator"
            End If
        End If
    Counter = .Range(OutputKlm & StartRow + 1).Value
    End With
    
    Error.DebugTekst Tekst:="Error values: " & vbNewLine _
                        & "->Counter: " & Counter & vbNewLine _
                        & "->SubName: " & SubName & vbNewLine _
                        & "->ErrLine: " & ErrLine & vbNewLine _
                        & "->ErrNumber: " & ErrNumber & vbNewLine _
                        & "->ErrSource: " & ErrSource & vbNewLine _
                        & "->ErrDescription: " & ErrDescription, _
                        FunctionName:="SeeText"
    
15 'back to the sheet were the error is indicated
    If Admin.NietBewerkbaar(ErrorShtName) = False Then
        Error.DebugTekst Tekst:="Basisgeg. can't be protected. End Error saving", FunctionName:="SeeText"
        Exit Function
    End If
    If ActiveSht <> ErrorShtName Then Admin.ShowOneSheet (ActiveSht)
        
20  'Send an email to the opporator/bugfix-er
    Error.SendError Counter, SubName, Msg, versie
    
30  Answer = MsgBox(Msg, vbQuestion + vbOKCancel, "Error", Err.HelpFile, Err.HelpContext)
    
    If Answer = vbCancel Then End

'--------End Function
Error.DebugTekst Tekst:="Finish", FunctionName:="SeeText"
Exit Function

ErrorText:
If Err.Number <> 0 Then SeeText (SubName)
Resume Next

End Function


Private Function SendError(Counter As Integer, FunctionName As String, _
                            Problem As String, versie As String)

SubName = "'SendError'"
If View("Errr") = True Then On Error GoTo ErrorText:

Application.ScreenUpdating = View("Updte")
Application.DisplayAlerts = View("Alrt")
Error.DebugTekst Tekst:="Start", FunctionName:=SubName
'--------Start Function

50
   EmailTo = "anko@zwervers.org"
   EmailFrom = "noreply@zwervers.org <noreply@zwervers.org>"
51
   EmailSubject = "Probleem in: " & ThisWorkbook.Name & " > " & FunctionName
52
   BodyText = "<font size=10px color=#FF0000>Er is een probleem gevonden in: <b>" & ThisWorkbook.Name & "</b></font></p>" _
            & "<p>Het probleem doet zich voor in de functie: <b>" & FunctionName & "</b></p>" _
            & "<p>Probleem beschrijving: </b>" & "</p><p><b>" _
            & Problem & "</b></p>"

53
   LogFile = LogFileName
54
   If Dir(LogFile) = "" Then LogFile = ""

70
Error.SendCDOmail eTo:=EmailTo, eFrom:=EmailFrom, eSubject:=EmailSubject, eBody:=BodyText, eAttach:=LogFile

'------------
71 'Check if the problem is mentioned for the first time -> place the error in the bug tracking list
If Counter = 1 Then
    EmailTo = "x+22990705897262@mail.asana.com; anko@zwervers.org"
    EmailFrom = "anko@zwervers.org <anko@zwervers.org>"
    EmailSubject = "v" & versie & " > " & FunctionName
    
    BodyText = "<p>Probleem beschrijving: </b>" & "</p><p><b>" _
                    & Problem & "</b></p>"

    Error.SendCDOmail eTo:=EmailTo, eFrom:=EmailFrom, eSubject:=EmailSubject, eBody:=BodyText, eAttach:=LogFile
    
    DebugTekst "Asana email send", SubName
End If
    
'--------End Function
Error.DebugTekst Tekst:="Finish", FunctionName:=SubName
Exit Function

ErrorText:
If Err.Number <> 0 Then SeeText (SubName)

Resume Next
End Function

Function DebugTekst(Tekst As String, Optional ByVal FunctionName As String, Optional AutoText As Boolean)

Dim s As String
Dim n As Integer
On Error Resume Next

ErrorLog = Error.LogFileName

n = FreeFile()
If Dir(ErrorLog) <> "" Then
    Open ErrorLog For Append As #n
Else
    Open ErrorLog For Output As #n
End If

If IsEmpty(AutoText) Or AutoText = False Then _
    If FunctionName <> "" Then Tekst = FunctionName & ">" & Tekst

If FunctionName = "" Then
    Debug.Print Tekst ' write to immediate
    Print #n, vbNewLine & Tekst ' write to file
Else
    Debug.Print "--" & Format(Now(), "dd-mm-yyyy hh:mm.ss") & vbNewLine & Tekst ' write to immediate
    Print #n, vbNewLine & "----" & Format(Now(), "dd-mm-yyyy hh:mm.ss") & vbNewLine & Tekst ' write to file
End If

Close #n

End Function

Sub SendErrorLog()

SubName = "'SendErrorLog'"
If View("Errr") = True Then On Error GoTo ErrorText:

Application.ScreenUpdating = View("Updte")
Application.DisplayAlerts = View("Alrt")
Error.DebugTekst Tekst:="Start", FunctionName:=SubName
'--------Start Function

Dim EmailTo, EmailFrom, EmailSubject, BodyText, LogFile As String

50
   EmailTo = "anko@zwervers.org"
   EmailFrom = "noreply@zwervers.org <noreply@zwervers.org>"
51
   EmailSubject = "Error log van: " & ThisWorkbook.Name & " > " & FunctionName
52
   BodyText = "<font size=10px color=#FF0000>Logbestand van: <b><br>" & ThisWorkbook.Name & "</b></font></p>"
   
53
   LogFile = LogFileName()
54
   If Dir(LogFile) = "" Then LogFile = ""

70
Error.SendCDOmail eTo:=EmailTo, eFrom:=EmailFrom, eSubject:=EmailSubject, eBody:=BodyText, _
                    eAttach:=LogFile

'--------End Function
Error.DebugTekst Tekst:="Finish", FunctionName:=SubName
Exit Sub

ErrorText:
If Err.Number <> 0 Then SeeText (SubName)

Resume Next

End Sub
