Attribute VB_Name = "Module1"
Dim hebrewDate As hdate
Dim shabbos As hdate
Dim erevshabbos As hdate
Dim here As location

Sub main()
    Dim parsh As parshah
    Dim shabbosstr As String
    Dim limudstr As String
    Dim jerusalem_offset As Integer
        
    jerusalem_offset = (2 + IsDST(Date))

    'convert date to hebrew date
    hebrewDate = ConvertDate(now)
    hebrewDate.offset = 3600 * jerusalem_offset
    Call SetEY(hebrewDate, 1)

    'load jerusalem location info
    With here ' =Yerushalayim (CL - 40 mins)
    .latitude = 31.788
    .longitude = 35.218
    .elevation = 800
    End With

    
    'show hebrew date + daf yomi:
    show_hebrew_today_and_daf_and_rambam_yomi

    'show today zmanim (jerusalem)
    show_zmanim_jer
    
    'show next shabbos / yomtov + candle lighting
    show_next_shabbos
    
    'show molad info for next month
    Call HDateAddMonth(hebrewDate, 1)
    molad_str = "מולד הלבנה חודש הבא - " & NumToHMonth(hebrewDate.month, hebrewDate.leap) & " " & NumToHChar(hebrewDate.year) & vbCrLf & MoladFormat(GetMolad(hebrewDate.year, hebrewDate.month))
    MsgBox molad_str

End Sub

Sub show_next_shabbos()
    Dim shabbos As hdate
    Dim erevshabbos As hdate
    Dim parsh As parshah, ytov As yomtov
    Dim shabbos_title As String
    Dim candle_lighting As String
    Dim mozash As String
    shabbos = hebrewDate
    
    'find next shabbos day, increment days until reaching a day which is AssurBeMelachah
    Do While IsAssurBeMelachah(shabbos) = 0
        Call HDateAddDay(shabbos, 1)
    Loop

    'get the parsha (as enum value)
    parsh = GetParshah(shabbos)

    If parsh <> NOPARSHAH Then
        'if the found date has a parshah then add it's parsha to the string
        'note - ParshahFormat converts parshah enum values to their titles in string
        shabbos_title = "שבת פרשת " & ParshahFormat(parsh)
    Else
        'if the found date has no parsha then it is probably a yomtov or shabbos chol ha'moed
        ytov = GetYomTov(shabbos)
        'convert the found date from a yomyov enum value type to string using YomTovFormat
        If ytov <> CHOL Then shabbos_title = YomTovFormat(ytov)
    End If

    'check if the found date is a spacial shabbos (hagadol, 4 parshios etc)
    If GetSpecialShabbos(shabbos) <> CHOL Then
            shabbos_title = shabbos_title & vbCrLf & YomTovFormat(GetSpecialShabbos(shabbos))
    End If
    
    'get erev shabbos (day before shabbos)
    erevshabbos = shabbos
    Call HDateAddDay(erevshabbos, -1)

    'get candle lighting for yerushalayim - 40 minutes before sunset
    candle_lighting = "הדלקת נרות: " & Format(DateAdd("n", -40, (HDateGregorian(getelevationsunset(erevshabbos, here)))), "hh:mm:ss")

    'get havdalah times, if relevant
    If IsCandleLighting(shabbos) = 0 Then
        mozash = "מוצש: " & Format((HDateGregorian(gettzais8p5(shabbos, here))), "hh:mm:ss") & vbCrLf & _
        "מוצש רת: " & Format(DateAdd("n", 72, (HDateGregorian(getelevationsunset(shabbos, here)))), "hh:mm:ss")
    End If
    'show the result
    MsgBox "השבת הקרובה:" & vbCrLf & shabbos_title & vbCrLf & candle_lighting & vbCrLf & mozash
End Sub

Sub show_hebrew_today_and_daf_and_rambam_yomi()
    MsgBox "התאריך העברי:" & vbCrLf & HDateFormat(hebrewDate) & vbCrLf & _
    "הדף היומי: " & GetDafYomiFormat(HDateGregorian(hebrewDate)) & vbCrLf & _
    "רמבם יומי: " & vbCrLf & GetRambam(hebrewDate, 0) & vbCrLf & _
    "פרק רמבם יומי: " & vbCrLf & GetRambam(hebrewDate, 1)
End Sub

Sub show_zmanim_jer()
    Dim zmanstr As String

    zmanstr = "זמני היום - ירושלים" & vbCrLf
    zmanstr = zmanstr & vbCrLf & "עלות השחר: " & Format((HDateGregorian(getalos(hebrewDate, here))), "hh:mm:ss")
    zmanstr = zmanstr & vbCrLf & "זריחה: " & Format((HDateGregorian(getsunrise(hebrewDate, here))), "hh:mm:ss")
    zmanstr = zmanstr & vbCrLf & "חצות היום: " & Format((HDateGregorian(getchatzosbaalhatanya(hebrewDate, here))), "hh:mm:ss")
    zmanstr = zmanstr & vbCrLf & "שקיעה במישור: " & Format((HDateGregorian(getsunset(hebrewDate, here))), "hh:mm:ss")
    zmanstr = zmanstr & vbCrLf & "שקיעה נראית: " & Format((HDateGregorian(getelevationsunset(hebrewDate, here))), "hh:mm:ss")
    zmanstr = zmanstr & vbCrLf & "צאת הכוכבים: " & Format((HDateGregorian(gettzaisbaalhatanya(hebrewDate, here))), "hh:mm:ss")
    zmanstr = zmanstr & vbCrLf & "צאת הכוכבים 8.5: " & Format((HDateGregorian(gettzais8p5(hebrewDate, here))), "hh:mm:ss")

    MsgBox zmanstr
End Sub



