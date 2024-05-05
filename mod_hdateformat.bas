Attribute VB_Name = "mod_hdateformat"
Public hchar(24) As String
Public hmonth(13) As String
Public hwday(7) As String
Public parshahchar(61) As String
Public hdateformat_ready As Boolean

Public Sub init_hdateformat()
    hchar(0) = "?"
    hchar(1) = "�"
    hchar(2) = "�"
    hchar(3) = "�"
    hchar(4) = "�"
    hchar(5) = "�"
    hchar(6) = "�"
    hchar(7) = "�"
    hchar(8) = "�"
    hchar(9) = "�"
    hchar(10) = "�"
    hchar(11) = "�"
    hchar(12) = "�"
    hchar(13) = "�"
    hchar(14) = "�"
    hchar(15) = "�"
    hchar(16) = "�"
    hchar(17) = "�"
    hchar(18) = "�"
    hchar(19) = "�"
    hchar(20) = "�"
    hchar(21) = "�"
    hchar(22) = "�"
    hchar(23) = "�"
    hchar(24) = "�"
    hmonth(0) = "��� ��"
    hmonth(1) = "����"
    hmonth(2) = "����"
    hmonth(3) = "����"
    hmonth(4) = "����"
    hmonth(5) = "��"
    hmonth(6) = "����"
    hmonth(7) = "����"
    hmonth(8) = "����"
    hmonth(9) = "����"
    hmonth(10) = "���"
    hmonth(11) = "���"
    hmonth(12) = "���"
    hmonth(13) = "��� ��"
    hwday(0) = "�����"
    hwday(1) = "�����"
    hwday(2) = "���"
    hwday(3) = "�����"
    hwday(4) = "�����"
    hwday(5) = "�����"
    hwday(6) = "����"
    hwday(7) = "���"
    parshahchar(0) = ""
    parshahchar(1) = "������"
    parshahchar(2) = "��"
    parshahchar(3) = "�� ��"
    parshahchar(4) = "����"
    parshahchar(5) = "��� ���"
    parshahchar(6) = "������"
    parshahchar(7) = "����"
    parshahchar(8) = "�����"
    parshahchar(9) = "����"
    parshahchar(10) = "���"
    parshahchar(11) = "����"
    parshahchar(12) = "����"
    parshahchar(13) = "����"
    parshahchar(14) = "����"
    parshahchar(15) = "��"
    parshahchar(16) = "����"
    parshahchar(17) = "����"
    parshahchar(18) = "������"
    parshahchar(19) = "�����"
    parshahchar(20) = "����"
    parshahchar(21) = "�� ���"
    parshahchar(22) = "�����"
    parshahchar(23) = "�����"
    parshahchar(24) = "�����"
    parshahchar(25) = "��"
    parshahchar(26) = "�����"
    parshahchar(27) = "�����"
    parshahchar(28) = "�����"
    parshahchar(29) = "���� ���"
    parshahchar(30) = "������"
    parshahchar(31) = "����"
    parshahchar(32) = "���"
    parshahchar(33) = "�������"
    parshahchar(34) = "�����"
    parshahchar(35) = "���"
    parshahchar(36) = "�������"
    parshahchar(37) = "���"
    parshahchar(38) = "���"
    parshahchar(39) = "���"
    parshahchar(40) = "���"
    parshahchar(41) = "�����"
    parshahchar(42) = "����"
    parshahchar(43) = "����"
    parshahchar(44) = "�����"
    parshahchar(45) = "������"
    parshahchar(46) = "���"
    parshahchar(47) = "���"
    parshahchar(48) = "������"
    parshahchar(49) = "�� ���"
    parshahchar(50) = "�� ����"
    parshahchar(51) = "�����"
    parshahchar(52) = "����"
    parshahchar(53) = "������"
    parshahchar(54) = "���� �����"
    parshahchar(55) = "����� - �����"
    parshahchar(56) = "����� - �����"
    parshahchar(57) = "���� ��� - ������"
    parshahchar(58) = "��� - �������"
    parshahchar(59) = "��� - ���"
    parshahchar(60) = "���� - ����"
    parshahchar(61) = "����� - ����"
    
    
    
    hdateformat_ready = True
End Sub

Public Function ParshahFormat(current As parshah) As String
    If hdateformat_ready = False Then init_hdateformat
    
    ParshahFormat = parshahchar(current)
End Function

Function GetHChar(ByVal num As Integer) As String
    Dim hchar1 As String
    If hdateformat_ready = False Then init_hdateformat
    Select Case num
        Case 1
            hchar1 = hchar(1)
        Case 2
            hchar1 = hchar(2)
        Case 3
            hchar1 = hchar(3)
        Case 4
            hchar1 = hchar(4)
        Case 5
            hchar1 = hchar(5)
        Case 6
            hchar1 = hchar(6)
        Case 7
            hchar1 = hchar(7)
        Case 8
            hchar1 = hchar(8)
        Case 9
            hchar1 = hchar(9)
        Case 10
            hchar1 = hchar(10)
        Case 20
            hchar1 = hchar(11)
        Case 30
            hchar1 = hchar(12)
        Case 40
            hchar1 = hchar(13)
        Case 50
            hchar1 = hchar(14)
        Case 60
            hchar1 = hchar(15)
        Case 70
            hchar1 = hchar(16)
        Case 80
            hchar1 = hchar(17)
        Case 90
            hchar1 = hchar(18)
        Case 100
            hchar1 = hchar(19)
        Case 200
            hchar1 = hchar(20)
        Case 300
            hchar1 = hchar(21)
        Case 400
            hchar1 = hchar(22)
        Case 99
            hchar1 = hchar(23)
        Case 999
            hchar1 = hchar(24)
        Case Else
            hchar1 = hchar(0)
    End Select
    GetHChar = hchar1
End Function

Function AddChar(ByVal year As String, ByVal charnum As Integer, ByRef num As Integer, ByRef counter As Long, ByVal limit As Long) As String
    Dim charvalue As Integer
    If hdateformat_ready = False Then init_hdateformat
    charvalue = charnum
    If charvalue = 99 Or charvalue = 999 Then charvalue = 0
    Dim Len1 As Long
    Len1 = limit - counter
    Dim endPos As Long
    endPos = counter + IIf(Len1 > 2, 2, Len1)
    Dim endStr As String
    endStr = Left$(year, counter) & GetHChar(charnum) & Mid$(year, endPos + 1)
    counter = counter + IIf(Len1 > 1, 1, Len1)
    num = num - charvalue
    AddChar = endStr
End Function

'convert an int to a Hebrew char based representation. 5779 becomes ���"�
Public Function NumToHChar(ByVal innum As Integer) As String
    Dim num As Integer
    num = innum
    Dim year As String
    Dim counter As Long
    If hdateformat_ready = False Then init_hdateformat
    year = String(13, vbNullChar)
    counter = 0
    If num >= 1000 And num <= 10000 And num Mod 1000 = 0 Then
        year = AddChar(year, num \ 1000, num, counter, 13)
        year = AddChar(year, 999, num, counter, 13)
        NumToHChar = Left$(year, counter)
        Exit Function
    End If
    If num >= 1000 And num <= 10000 Then num = num Mod 1000
    Do While num > 0 And counter < 13
        If num = 15 Or num = 16 Then
            year = AddChar(year, 9, num, counter, 13)
            year = AddChar(year, 99, num, counter, 13)
            year = AddChar(year, num, num, counter, 13)
            NumToHChar = Left$(year, counter)
            Exit Function
        ElseIf num < 10 Or (num < 100 And num Mod 10 = 0) Or (num < 500 And num Mod 100 = 0) Then
            If counter <> 0 Then year = AddChar(year, 99, num, counter, 13)
            year = AddChar(year, num, num, counter, 13)
            NumToHChar = Left$(year, counter)
            If innum < 11 Then NumToHChar = NumToHChar + "'"
            Exit Function
        ElseIf num > 400 Then
            year = AddChar(year, 400, num, counter, 13)
            GoTo ContinueLoop
        ElseIf num > 300 Then
            year = AddChar(year, 300, num, counter, 13)
            GoTo ContinueLoop
        ElseIf num > 200 Then
            year = AddChar(year, 200, num, counter, 13)
            GoTo ContinueLoop
        ElseIf num > 100 Then
            year = AddChar(year, 100, num, counter, 13)
            GoTo ContinueLoop
        ElseIf num \ 10 > 0 Then
            year = AddChar(year, num - (num Mod 10), num, counter, 13)
            GoTo ContinueLoop
        End If
ContinueLoop:
    Loop
    NumToHChar = Left$(year, counter)
    If innum < 11 Then NumToHChar = NumToHChar + "'"
End Function

'convert int based Hebrew weekday (hdate.wday) to char based representation.
'second argument is a booean if to use ��� (true) or ����� (false)
Function NumToWDay(date_in As hdate, ByVal shabbos As Boolean) As String
    If hdateformat_ready = False Then init_hdateformat
    If shabbos And date_in.wday = 0 Then
        NumToWDay = hwday(7)
    Else
        NumToWDay = hwday(date_in.wday)
    End If
End Function

'convert int based Hebrew month (hdate.month) to char based representation.
Function NumToHMonth(ByVal month As Integer, ByVal leap As Integer) As String
    If hdateformat_ready = False Then init_hdateformat
    If leap <> 0 Then
        If month = 12 Then
            NumToHMonth = hmonth(0)
            Exit Function
        ElseIf month = 13 Then
            NumToHMonth = hmonth(month)
            Exit Function
        End If
    End If
    If month > 0 And month < 13 Then
        NumToHMonth = hmonth(month)
    Else
        NumToHMonth = vbNullString
    End If
End Function

'convert hdate to string based representation
Function HDateFormat(date_in As hdate) As String
    Dim day As String
    Dim year As String
    If hdateformat_ready = False Then init_hdateformat
    day = NumToHChar(date_in.day)
    Dim month As String
    month = NumToHMonth(date_in.month, date_in.leap)
    year = NumToHChar(date_in.year)
    HDateFormat = day & " " & month & " " & year
End Function

'convert hdate to string based representation, with evening consideration;
'if the time is between zais to sunrise of next day then add "��� �-" for next day
Function HDateOrFormat(date_in As hdate, here As location) As String
    Dim day As String
    Dim year As String
    Dim date_next As hdate
    Dim date_before As hdate
    Dim is_or As Boolean
    Dim sunset_today As Date
    Dim sunset_yesterday As Date
    Dim current_date As Date
    Dim sunrise_tomorrow As Date
    Dim date_result As hdate
    date_next = date_in
    date_before = date_in
    current_date = (HDateGregorian(date_in))
    Call HDateAddDay(date_next, 1)
    Call HDateAddDay(date_before, -1)
    sunset_today = (HDateGregorian(gettzais8p5(date_in, here)))
    sunrise_today = (HDateGregorian(getsunrise(date_in, here)))
    sunrise_tomorrow = (HDateGregorian(getsunrise(date_next, here)))
    sunset_yesterday = (HDateGregorian(gettzais8p5(date_before, here)))
    If current_date >= sunset_today Then
        is_or = True
        date_result = date_next
    Else
        date_result = date_in
    End If
    If current_date < sunrise_today Then is_or = True
    If hdateformat_ready = False Then init_hdateformat
    day = NumToHChar(date_result.day)
    Dim month As String
    month = NumToHMonth(date_result.month, date_result.leap)
    year = NumToHChar(date_result.year)
    HDateOrFormat = day & " " & month & " " & year
    If is_or Then HDateOrFormat = "��� �-" & HDateOrFormat
End Function

'convert hdate holding molad info to string representation, suitable for molad announcement
Function MoladFormat(molad As hdate, Optional full_date As Boolean = True) As String
    If full_date Then
        MoladFormat = "��� " & NumToWDay(molad, True) & ", " & NumToHChar(molad.day) & " " & NumToHMonth(molad.month, molad.leap) & ", ��� " & Format(TimeSerial(molad.hour, molad.min, 0), "hh:mm") & " �-" & molad.sec & " �����"
    Else
        MoladFormat = "��� " & NumToWDay(molad, True) & " ��� " & Format(TimeSerial(molad.hour, molad.min, 0), "hh:mm") & " �-" & molad.sec & " �����"
    End If
End Function

'convert yomtov enum to char based representation
Function YomTovFormat(ByVal current As yomtov) As String
    If hdateformat_ready = False Then init_hdateformat
    Select Case current
        Case CHOL
            YomTovFormat = ""
        Case PESACH_DAY1, PESACH_DAY2
            YomTovFormat = "���"
        Case SHVEI_SHEL_PESACH
            YomTovFormat = "����� �� ���"
        Case ACHRON_SHEL_PESACH
            YomTovFormat = "����� �� ���"
        Case SHAVOUS_DAY1, SHAVOUS_DAY2
            YomTovFormat = "������"
        Case ROSH_HASHANAH_DAY1, ROSH_HASHANAH_DAY2
            YomTovFormat = "��� ����"
        Case YOM_KIPPUR
            YomTovFormat = "��� �����"
        Case SUKKOS_DAY1, SUKKOS_DAY2
            YomTovFormat = "�����"
        Case SHMEINI_ATZERES
            YomTovFormat = "����� ����"
        Case SIMCHAS_TORAH
            YomTovFormat = "���� ����"
        Case CHOL_HAMOED_PESACH_DAY1 To CHOL_HAMOED_PESACH_DAY5
            YomTovFormat = "��� ����� ���"
        Case CHOL_HAMOED_SUKKOS_DAY1 To CHOL_HAMOED_SUKKOS_DAY5
            YomTovFormat = "��� ����� �����"
        Case HOSHANA_RABBAH
            YomTovFormat = "������ ���"
        Case PESACH_SHEINI
            YomTovFormat = "��� ���"
        Case LAG_BAOMER
            YomTovFormat = "��� �����"
        Case TU_BAV
            YomTovFormat = "��� ���"
        Case CHANUKAH_DAY1 To CHANUKAH_DAY8
            YomTovFormat = "�����"
        Case TU_BISHVAT
            YomTovFormat = "��� ����"
        Case PURIM_KATAN
            YomTovFormat = "����� ���"
        Case SHUSHAN_PURIM_KATAN
            YomTovFormat = "���� ����� ���"
        Case PURIM
            YomTovFormat = "�����"
        Case SHUSHAN_PURIM
            YomTovFormat = "���� �����"
        Case SHIVA_ASAR_BTAAMUZ
            YomTovFormat = "���� ��� �����"
        Case TISHA_BAV
            YomTovFormat = "���"
        Case TZOM_GEDALIA
            YomTovFormat = "��� �����"
        Case ASARAH_BTEVES
            YomTovFormat = "���� ����"
        Case TAANIS_ESTER
            YomTovFormat = "����� ����"
        Case EREV_PESACH
            YomTovFormat = "��� ���"
        Case EREV_SHAVOUS
            YomTovFormat = "��� ������"
        Case EREV_ROSH_HASHANAH
            YomTovFormat = "��� ��� ����"
        Case EREV_YOM_KIPPUR
            YomTovFormat = "��� ��� �����"
        Case EREV_SUKKOS
            YomTovFormat = "��� �����"
        Case SHKALIM
            YomTovFormat = "���� �����"
        Case ZACHOR
            YomTovFormat = "���� ����"
        Case PARAH:
            YomTovFormat = "���� ���"
        Case HACHODESH:
            YomTovFormat = "���� �����"
        Case ROSH_CHODESH:
            YomTovFormat = "��� ����"
        Case MACHAR_CHODESH:
            YomTovFormat = "��� ����"
        Case SHABBOS_MEVORCHIM:
            YomTovFormat = "��� ������"
        Case HAGADOL:
            YomTovFormat = "��� �����"
        Case CHAZON:
            YomTovFormat = "��� ����"
        Case NACHAMU:
            YomTovFormat = "��� ����"
        Case SHUVA:
            YomTovFormat = "��� ����"
        Case SHIRA:
            YomTovFormat = "��� ����"
        Case SHABBOS_CHOL_HAMOED:
            YomTovFormat = "��� ��� �����"
        Case Else
            YomTovFormat = ""
    End Select
End Function

'convert avos int to char based representation
Function AvosFormat(ByVal avos As Integer) As String
    If hdateformat_ready = False Then init_hdateformat
    Select Case avos
        Case 1
            AvosFormat = "�"
        Case 2
            AvosFormat = "�"
        Case 3
            AvosFormat = "�"
        Case 4
            AvosFormat = "�"
        Case 5
            AvosFormat = "�"
        Case 6
            AvosFormat = "�"
        Case 12
            AvosFormat = "�-�"
        Case 34
            AvosFormat = "�-�"
        Case 56
            AvosFormat = "�-�"
        Case Else
            AvosFormat = ""
    End Select
End Function

'rounds seconds of a time to an added or substracted minute
Function tround(t As Date) As Date
    tround = IIf(second(t) > 29, DateAdd("n", 1, t), t)
End Function

