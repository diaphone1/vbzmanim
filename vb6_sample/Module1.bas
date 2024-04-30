Attribute VB_Name = "Module1"
Dim hebrewDate As hdate
Dim shabbos As hdate
Dim erevshabbos As hdate
Dim here As location
Dim data_tag_col As New Collection


Sub Main()
    Dim zmanstr As String
    Dim parsh As parshah
    Dim shabbosstr As String
    Dim limudstr As String
    
    With here ' =Yerushalayim
    .latitude = 31.788
    .longitude = 35.218
    .elevation = 800
    End With
    
    hebrewDate = ConvertDate(mktm(now))
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
    shabbosstr = "��� ���� " & parshahformat(parsh)
Else
    ytov = GetYomTov(shabbos)
    If ytov <> CHOL Then shabbosstr = YomTovFormat(ytov)
End If
ytov = GetSpecialShabbos(shabbos)
shabbosstr = shabbosstr & IIf(ytov <> CHOL, vbCrLf & "(" & YomTovFormat(ytov) & ")", "")


zmanstr = HDateOrFormat(hebrewDate, here) & vbCrLf
zmanstr = zmanstr & vbCrLf & "���� ����: " & Format(mkdate(HDateGregorian(getalos(hebrewDate, here))), "hh:mm:ss")
zmanstr = zmanstr & vbCrLf & "�����: " & Format(mkdate(HDateGregorian(getsunrise(hebrewDate, here))), "hh:mm:ss")
zmanstr = zmanstr & vbCrLf & "���� ����: " & Format(mkdate(HDateGregorian(getchatzosbaalhatanya(hebrewDate, here))), "hh:mm:ss")
zmanstr = zmanstr & vbCrLf & "����� ������: " & Format(mkdate(HDateGregorian(getsunset(hebrewDate, here))), "hh:mm:ss")
zmanstr = zmanstr & vbCrLf & "����� �����: " & Format(mkdate(HDateGregorian(getelevationsunset(hebrewDate, here))), "hh:mm:ss")
zmanstr = zmanstr & vbCrLf & "��� �������: " & Format(mkdate(HDateGregorian(gettzaisbaalhatanya(hebrewDate, here))), "hh:mm:ss")
zmanstr = zmanstr & vbCrLf & "��� ������� 8.5: " & Format(mkdate(HDateGregorian(gettzais8p5(hebrewDate, here))), "hh:mm:ss")

limudstr = HDateOrFormat(hebrewDate, here) & vbCrLf
limudstr = limudstr & vbCrLf & "�� ����: " & vbCrLf & GetDafYomiFormat(mkdate(HDateGregorian(hebrewDate))) & vbCrLf
limudstr = limudstr & vbCrLf & "���� ����: " & vbCrLf & GetRambam(hebrewDate, 0) & vbCrLf
limudstr = limudstr & vbCrLf & "��� ���� ����: " & vbCrLf & GetRambam(hebrewDate, 1) & vbCrLf

shabbosstr = shabbosstr & vbCrLf
shabbosstr = shabbosstr & vbCrLf & "����� ����: " & Format(DateAdd("n", -40, mkdate(HDateGregorian(getelevationsunset(erevshabbos, here)))), "hh:mm:ss")
shabbosstr = shabbosstr & vbCrLf & "����: " & Format(mkdate(HDateGregorian(gettzais8p5(shabbos, here))), "hh:mm:ss")
shabbosstr = shabbosstr & vbCrLf & "���� ��: " & Format(DateAdd("n", 72, mkdate(HDateGregorian(getelevationsunset(shabbos, here)))), "hh:mm:ss")

MsgBox "���� ���� - �������" & vbCrLf & zmanstr
MsgBox "����� ���� - " & limudstr
MsgBox "���� ���� ������ - �������" & vbCrLf & shabbosstr
        
End Sub


