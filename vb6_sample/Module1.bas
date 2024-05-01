Attribute VB_Name = "Module1"
Dim hebrewDate As hdate
Dim shabbos As hdate
Dim erevshabbos As hdate
Dim here As location

Sub Test_Zmanim2()
Dim heb_date As hdate
Dim parsh As parshah, ytov As yomtov
Dim shabbos_title As String
heb_date = ConvertDate(Date)
'find next shabbos day, increment days until reaching a day which is AssurBeMelachah
Do While IsAssurBeMelachah(heb_date) = 0
    Call HDateAddDay(heb_date, 1)
Loop

'get the parsha (as enum value)
parsh = GetParshah(heb_date)

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
'show the result
MsgBox "השבת הקרובה:" & vbCrLf & shabbos_title
End Sub

Sub main()
    Dim zmanstr As String
    Dim parsh As parshah
    Dim shabbosstr As String
    Dim limudstr As String
    
    With here ' =Yerushalayim
    .latitude = 31.788
    .longitude = 35.218
    .elevation = 800
    End With
    
    hebrewDate = ConvertDate(now)
    hebrewDate.offset = 3600 * 3
    
    Call SetEY(hebrewDate, 1)
    
    shabbos = hebrewDate
    erevshabbos = hebrewDate
    
    If IsAssurBeMelachah(erevshabbos) Then
        Call HDateAddDay(erevshabbos, -1)
    Else
        Do While IsCandleLighting(erevshabbos) = 0
            Call HDateAddDay(erevshabbos, 1)
        Loop 'TODO check chanuka CL on wday=1
        shabbos = erevshabbos
        Call HDateAddDay(shabbos, 1)
    End If
    


parsh = GetParshah(shabbos)
If parsh <> NOPARSHAH Then
    shabbosstr = "שבת פרשת " & ParshahFormat(parsh)
Else
    ytov = GetYomTov(shabbos)
    If ytov <> CHOL Then shabbosstr = YomTovFormat(ytov)
End If
ytov = GetSpecialShabbos(shabbos)
shabbosstr = shabbosstr & IIf(ytov <> CHOL, vbCrLf & "(" & YomTovFormat(ytov) & ")", "")


zmanstr = HDateOrFormat(hebrewDate, here) & vbCrLf
zmanstr = zmanstr & vbCrLf & "עלות השחר: " & Format((HDateGregorian(getalos(hebrewDate, here))), "hh:mm:ss")
zmanstr = zmanstr & vbCrLf & "זריחה: " & Format((HDateGregorian(getsunrise(hebrewDate, here))), "hh:mm:ss")
zmanstr = zmanstr & vbCrLf & "חצות היום: " & Format((HDateGregorian(getchatzosbaalhatanya(hebrewDate, here))), "hh:mm:ss")
zmanstr = zmanstr & vbCrLf & "שקיעה במישור: " & Format((HDateGregorian(getsunset(hebrewDate, here))), "hh:mm:ss")
zmanstr = zmanstr & vbCrLf & "שקיעה נראית: " & Format((HDateGregorian(getelevationsunset(hebrewDate, here))), "hh:mm:ss")
zmanstr = zmanstr & vbCrLf & "צאת הכוכבים: " & Format((HDateGregorian(gettzaisbaalhatanya(hebrewDate, here))), "hh:mm:ss")
zmanstr = zmanstr & vbCrLf & "צאת הכוכבים 8.5: " & Format((HDateGregorian(gettzais8p5(hebrewDate, here))), "hh:mm:ss")

limudstr = HDateOrFormat(hebrewDate, here) & vbCrLf
limudstr = limudstr & vbCrLf & "דף יומי: " & vbCrLf & GetDafYomiFormat((HDateGregorian(hebrewDate))) & vbCrLf
limudstr = limudstr & vbCrLf & "רמבם יומי: " & vbCrLf & GetRambam(hebrewDate, 0) & vbCrLf
limudstr = limudstr & vbCrLf & "פרק רמבם יומי: " & vbCrLf & GetRambam(hebrewDate, 1) & vbCrLf

shabbosstr = shabbosstr & vbCrLf
shabbosstr = shabbosstr & vbCrLf & "הדלקת נרות: " & Format(DateAdd("n", -40, (HDateGregorian(getelevationsunset(erevshabbos, here)))), "hh:mm:ss")
shabbosstr = shabbosstr & vbCrLf & "מוצש: " & Format((HDateGregorian(gettzais8p5(shabbos, here))), "hh:mm:ss")
shabbosstr = shabbosstr & vbCrLf & "מוצש רת: " & Format(DateAdd("n", 72, (HDateGregorian(getelevationsunset(shabbos, here)))), "hh:mm:ss")

'molad_str = "מולד הלבנה חודש נוכחי - " & NumToHMonth(hebrewDate.month, hebrewDate.leap) & vbCrLf & MoladFormat(GetMolad(hebrewDate.year, hebrewDate.month))
Call HDateAddMonth(hebrewDate, 1)
molad_str = "מולד הלבנה חודש הבא - " & NumToHMonth(hebrewDate.month, hebrewDate.leap) & " " & NumToHChar(hebrewDate.year) & vbCrLf & MoladFormat(GetMolad(hebrewDate.year, hebrewDate.month))

MsgBox "זמני היום - ירושלים" & vbCrLf & zmanstr
MsgBox "לימוד יומי - " & limudstr
Test_Zmanim2
MsgBox "זמני השבת הקרובה - ירושלים" & vbCrLf & shabbosstr
MsgBox molad_str

End Sub




