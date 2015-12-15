Attribute VB_Name = "IbanCheck"
'Option Compare Database
Option Explicit

' http://en.wikipedia.org/wiki/International_Bank_Account_Number
Private Const IbanCountryLengths As String = "AL28AD24AT20AZ28BH22BE16BA20BR29BG22CR21HR21CY28CZ24DK18DO28EE20FO18" & _
                                             "FI18FR27GE22DE22GI23GR27GL18GT28HU28IS26IE22IL23IT27KZ20KW30LV21LB28" & _
                                             "LI21LT20LU20MK19MT31MR27MU30MC27MD24ME22NL18NO15PK24PS29PL28PT25RO24" & _
                                             "SM27SA24RS22SK24SI19ES24SE24CH21TN24TR26AE23GB22VG24QA29"

Private Function ValidateIbanCountryLength(CountryCode As String, IbanLength As Integer) As Boolean
    Dim i As Integer
    For i = 0 To Len(IbanCountryLengths) / 4 - 1
        If Mid(IbanCountryLengths, i * 4 + 1, 2) = CountryCode And _
                    CInt(Mid(IbanCountryLengths, i * 4 + 3, 2)) = IbanLength Then
            ValidateIbanCountryLength = True
            Exit Function
        End If
    Next i
    ValidateIbanCountryLength = False
End Function

Private Function Mod97(Num As String) As Integer
    Dim lngTemp As Long
    Dim strTemp As String

    Do While Val(Num) >= 97
        If Len(Num) > 5 Then
            strTemp = Left(Num, 5)
            Num = Right(Num, Len(Num) - 5)
        Else
            strTemp = Num
            Num = ""
        End If
        lngTemp = CLng(strTemp)
        lngTemp = lngTemp Mod 97
        strTemp = CStr(lngTemp)
        Num = strTemp & Num
    Loop
    Mod97 = CInt(Num)
End Function

Public Function ValidateIban(IBAN As String) As Boolean
    Dim strIban As String
    Dim i As Integer

    strIban = UCase(IBAN)
    ' Remove spaces
    strIban = Replace(strIban, " ", "")

    ' Check if IBAN contains only uppercase characters and numbers
    For i = 1 To Len(strIban)
        If Not ((Asc(Mid(strIban, i, 1)) <= Asc("9") And Asc(Mid(strIban, i, 1)) >= Asc("0")) Or _
                (Asc(Mid(strIban, i, 1)) <= Asc("Z") And Asc(Mid(strIban, i, 1)) >= Asc("A"))) Then
            ValidateIban = False
            Exit Function
        End If
    Next i

    ' Check if length of IBAN equals expected length for country
    If Not ValidateIbanCountryLength(Left(strIban, 2), Len(strIban)) Then
        ValidateIban = False
        Exit Function
    End If

    ' Rearrange
    strIban = Right(strIban, Len(strIban) - 4) & Left(strIban, 4)

    ' Replace characters
    For i = 0 To 25
        strIban = Replace(strIban, Chr(i + Asc("A")), i + 10)
    Next i

    ' Check remainder
    ValidateIban = Mod97(strIban) = 1
End Function

