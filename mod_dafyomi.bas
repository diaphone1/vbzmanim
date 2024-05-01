Attribute VB_Name = "mod_dafyomi"
'dafyomi list & functions ported from https://github.com/NykUser/MyZman/

Public masechtosBavli(39) As String
Public masechtosBavliTransliterated(39) As String
Public blattPerMasechta(39) As Integer
Public Type Daf
    masechtaNumber As Integer
    Page As Integer
    HasSecondaryMesechta As Boolean
    SecondaryMesechtaNumber As Integer
End Type
Private dafYomiStartDate As Date
Private shekalimChangeDate As Date
Dim dafyom_ready As Boolean

Sub init_dafyomi()
    masechtosBavli(0) = "ברכות"
    masechtosBavli(1) = "שבת"
    masechtosBavli(2) = "עירובין"
    masechtosBavli(3) = "פסחים"
    masechtosBavli(4) = "שקלים"
    masechtosBavli(5) = "יומא"
    masechtosBavli(6) = "סוכה"
    masechtosBavli(7) = "ביצה"
    masechtosBavli(8) = "ראש השנה"
    masechtosBavli(9) = "תענית"
    masechtosBavli(10) = "מגילה"
    masechtosBavli(11) = "מועד קטן"
    masechtosBavli(12) = "חגיגה"
    masechtosBavli(13) = "יבמות"
    masechtosBavli(14) = "כתובות"
    masechtosBavli(15) = "נדרים"
    masechtosBavli(16) = "נזיר"
    masechtosBavli(17) = "סוטה"
    masechtosBavli(18) = "גיטין"
    masechtosBavli(19) = "קידושין"
    masechtosBavli(20) = "בבא קמא"
    masechtosBavli(21) = "בבא מציעא"
    masechtosBavli(22) = "בבא בתרא"
    masechtosBavli(23) = "סנהדרין"
    masechtosBavli(24) = "מכות"
    masechtosBavli(25) = "שבועות"
    masechtosBavli(26) = "עבודה זרה"
    masechtosBavli(27) = "הוריות"
    masechtosBavli(28) = "זבחים"
    masechtosBavli(29) = "מנחות"
    masechtosBavli(30) = "חולין"
    masechtosBavli(31) = "בכורות"
    masechtosBavli(32) = "ערכין"
    masechtosBavli(33) = "תמורה"
    masechtosBavli(34) = "כריתות"
    masechtosBavli(35) = "מעילה"
    masechtosBavli(36) = "תמיד"
    masechtosBavli(37) = "קינים"
    masechtosBavli(38) = "מידות"
    masechtosBavli(39) = "נדה"
    masechtosBavliTransliterated(0) = "Berachos"
    masechtosBavliTransliterated(1) = "Shabbos"
    masechtosBavliTransliterated(2) = "Eruvin"
    masechtosBavliTransliterated(3) = "Pesachim"
    masechtosBavliTransliterated(4) = "Shekalim"
    masechtosBavliTransliterated(5) = "Yoma"
    masechtosBavliTransliterated(6) = "Sukkah"
    masechtosBavliTransliterated(7) = "Beitzah"
    masechtosBavliTransliterated(8) = "Rosh Hashana"
    masechtosBavliTransliterated(9) = "Taanis"
    masechtosBavliTransliterated(10) = "Megillah"
    masechtosBavliTransliterated(11) = "Moed Katan"
    masechtosBavliTransliterated(12) = "Chagigah"
    masechtosBavliTransliterated(13) = "Yevamos"
    masechtosBavliTransliterated(14) = "Kesubos"
    masechtosBavliTransliterated(15) = "Nedarim"
    masechtosBavliTransliterated(16) = "Nazir"
    masechtosBavliTransliterated(17) = "Sotah"
    masechtosBavliTransliterated(18) = "Gitin"
    masechtosBavliTransliterated(19) = "Kiddushin"
    masechtosBavliTransliterated(20) = "Bava Kamma"
    masechtosBavliTransliterated(21) = "Bava Metzia"
    masechtosBavliTransliterated(22) = "Bava Basra"
    masechtosBavliTransliterated(23) = "Sanhedrin"
    masechtosBavliTransliterated(24) = "Makkos"
    masechtosBavliTransliterated(25) = "Shevuos"
    masechtosBavliTransliterated(26) = "Avodah Zarah"
    masechtosBavliTransliterated(27) = "Horiyos"
    masechtosBavliTransliterated(28) = "Zevachim"
    masechtosBavliTransliterated(29) = "Menachos"
    masechtosBavliTransliterated(30) = "Chullin"
    masechtosBavliTransliterated(31) = "Bechoros"
    masechtosBavliTransliterated(32) = "Arachin"
    masechtosBavliTransliterated(33) = "Temurah"
    masechtosBavliTransliterated(34) = "Kerisos"
    masechtosBavliTransliterated(35) = "Meilah"
    masechtosBavliTransliterated(36) = "Kinnim"
    masechtosBavliTransliterated(37) = "Tamid"
    masechtosBavliTransliterated(38) = "Midos"
    masechtosBavliTransliterated(39) = "Niddah"
    blattPerMasechta(0) = 64
    blattPerMasechta(1) = 157
    blattPerMasechta(2) = 105
    blattPerMasechta(3) = 121
    blattPerMasechta(4) = 22
    blattPerMasechta(5) = 88
    blattPerMasechta(6) = 56
    blattPerMasechta(7) = 40
    blattPerMasechta(8) = 35
    blattPerMasechta(9) = 31
    blattPerMasechta(10) = 32
    blattPerMasechta(11) = 29
    blattPerMasechta(12) = 27
    blattPerMasechta(13) = 122
    blattPerMasechta(14) = 112
    blattPerMasechta(15) = 91
    blattPerMasechta(16) = 66
    blattPerMasechta(17) = 49
    blattPerMasechta(18) = 90
    blattPerMasechta(19) = 82
    blattPerMasechta(20) = 119
    blattPerMasechta(21) = 119
    blattPerMasechta(22) = 176
    blattPerMasechta(23) = 113
    blattPerMasechta(24) = 24
    blattPerMasechta(25) = 49
    blattPerMasechta(26) = 76
    blattPerMasechta(27) = 14
    blattPerMasechta(28) = 120
    blattPerMasechta(29) = 110
    blattPerMasechta(30) = 142
    blattPerMasechta(31) = 61
    blattPerMasechta(32) = 34
    blattPerMasechta(33) = 34
    blattPerMasechta(34) = 28
    blattPerMasechta(35) = 22
    blattPerMasechta(36) = 4
    blattPerMasechta(37) = 9
    blattPerMasechta(38) = 5
    blattPerMasechta(39) = 73

    dafYomiStartDate = DateSerial(1923, 9, 11)
    shekalimChangeDate = DateSerial(1975, 6, 24)
    dafyom_ready = True
    
End Sub
        Private Function IsCurrentBlattWithNextMasechta(ByVal masechta As Integer, ByVal blatt As Integer) As Boolean
            IsCurrentBlattWithNextMasechta = (masechta = 35 And blatt = 22) Or (masechta = 36 And blatt = 25)
        End Function

        Public Function GetDafYomiBavli(ByVal date_in As Date) As Daf
            ''
            ' * The number of daf per masechta. Since the number of blatt in Shekalim changed on the 8th Daf Yomi cycle
            ' * beginning on June 24, 1975 from 13 to 22, the actual calculation for blattPerMasechta[4] will later be
            ' * adjusted based on the cycle.

            
            Dim dafYomi As Daf ' = Nothing
            Dim julianDay '= GetJulianDay([date])SecondaryMesechtaNumber
            julianDay = GetJulianDay(date_in)

            Dim cycleNo ' = 0
            Dim dafNo As Integer ' = 0
            If dafyom_ready = False Then init_dafyomi
            
            If date_in < dafYomiStartDate Then
                ' TODO: should we return a null or throw an IllegalArgumentException?
                'Throw New ArgumentException([date] & " is prior to organized Daf Yomi Bavli cycles that started on " & dafYomiStartDate)
                Exit Function
            End If

            'Changed in Vb to use DateDiff and not JulianDay as it was off, probably issue with time in GetJulianDay
            Dim DaysFromDafYomiStart As Long: DaysFromDafYomiStart = DateDiff("d", dafYomiStartDate, date_in)
            Dim DaysFromShekalimChange As Long: DaysFromShekalimChange = DateDiff("d", shekalimChangeDate, date_in)

            If date_in >= shekalimChangeDate Then
                'cycleNo = 8 + ((julianDay - shekalimJulianChangeDay) / 2711)
                cycleNo = 8 + (DaysFromShekalimChange / 2711)
                'dafNo = ((julianDay - shekalimJulianChangeDay) Mod 2711)
                dafNo = DaysFromShekalimChange Mod 2711
            Else
                'cycleNo = 1 + ((julianDay - dafYomiJulianStartDay) / 2702)
                cycleNo = 1 + (DaysFromDafYomiStart / 2702)
                'dafNo = ((julianDay - dafYomiJulianStartDay) Mod 2702)
                dafNo = DaysFromDafYomiStart Mod 2702
            End If

            'Debug.Print(dafNo)

            Dim total ' = 0
            Dim masechta: masechta = -1
            Dim blatt ' = 0

            ' Fix Shekalim for old cycles.
            If cycleNo <= 7 Then
                blattPerMasechta(4) = 13
            Else
                blattPerMasechta(4) = 22 ' correct any change that may have been changed from a prior calculation
            End If

            For j = 0 To 39
                masechta = masechta + 1
                total = total + blattPerMasechta(j) - 1

                If dafNo < total Then
                    blatt = 1 + blattPerMasechta(j) - (total - dafNo)
                    ' Fiddle with the weird ones near the end.
                    If masechta = 36 Then
                        blatt = blatt + 21
                    ElseIf masechta = 37 Then
                        blatt = blatt + 24
                    ElseIf masechta = 38 Then
                        blatt = blatt + 32
                    End If

                    Dim isWithNextMasechta: isWithNextMasechta = IsCurrentBlattWithNextMasechta(masechta, blatt)
                    'dafYomi= New Daf(masechta, blatt, isWithNextMasechta)
                    dafYomi.masechtaNumber = masechta
                    dafYomi.Page = blatt
                    dafYomi.HasSecondaryMesechta = isWithNextMasechta
                    dafYomi.SecondaryMesechtaNumber = masechta + 1
                    Exit For
                End If
            Next

            GetDafYomiBavli = dafYomi
        End Function

        Public Function GetJulianDay(ByVal date_in As Date) As Long
            Dim calendar As Date
            calendar = date_in
            Dim y As Double: y = year(date_in)
            Dim m As Double: m = month(date_in)
            Dim d As Double: d = day(date_in)
            
            If m <= 2 Then
                y = y - 1
                m = m + 12
            End If

            Dim a As Double: a = y / 100
            Dim b As Double: b = 2 - a + a / 4
            GetJulianDay = Fix((Fix(365.25 * (y + 4716#)) + Fix(30.6001 * (m + 1#)) + d + b - 1524.5))
            'Return CInt((Math.Floor(365.25 * (year + 4716)) + Math.Floor(30.6001 * (month + 1)) + day + b - 1524.5))

            'Return Math.Floor(365.25 * (year + 4716)) + Math.Floor(30.6001 * (month + 1)) + day + b - 1524.5
        End Function

Public Function GetDafYomiFormat(ByVal date_in As Date)
    Dim result As Daf
    If hdateformat_ready = False Then init_hdateformat

    result = GetDafYomiBavli(date_in)
    
    GetDafYomiFormat = masechtosBavli(result.masechtaNumber) & " דף " & NumToHChar(result.Page)
    
    
End Function
