Attribute VB_Name = "Start"
Public MyFilePribor As String 'Файл в котором храним приборы
Public AllCBp As New Collection  'Переменная для хранения  всех пунктов
Sub А2()
'
' А1 Макрос
' Макрос создан 15.04.2011
'
'Set AllCBp = New Scripting.Dictionary
MyFilePribor = ActiveDocument.AttachedTemplate.Path
MyFilePribor = MyFilePribor & "\tablprib.txt"
UF1.Show
    
'    Dim ctl As Variant
'    For Each ctl In UF1.Controls
'    If TypeName(ctl) = "TextBox" Then
'    Call FoAndRe(ctl.Name & "$", ctl.Value)
'    End If
'    Next

Application.ScreenUpdating = True
ActiveDocument.Fields.Update
With Dialogs(wdDialogFileSaveAs)
    .Name = ActiveDocument.BuiltInDocumentProperties("Title").Value
    .Show
End With
Call UnlinkBookmarks("Punkt") 'удаляем ссылки
Call FoAndRe(Strings.Chr(187) & Strings.Chr(187), Strings.Chr(187))
'Call FoAndRe("^-;", "")
Call FoAndRe("^-", "") ' удаление всех пустых переменных в тексте
Call FoAndRe(", Феде", " Феде")
Call FoAndRe("..", ".")
ActiveDocument.AttachedTemplate = ""

End Sub

'Функция подставляет расширенный формат даты
 Function FormDat(dat1 As Date)

FormDat = Format(dat1, Strings.Chr(171) & "dd" & Strings.Chr(187) & " MMMM yyyy" & " г.")
'FormDat = Strings.Chr(171) & Left(FormDat, 2) & Strings.Chr(187) & Right(FormDat, Len(FormDat) - 2) & " г."

End Function

'Фунцция поиска и замены
Sub FoAndRe(A1 As String, A2)

    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = A1
        .Replacement.Text = A2
'        .Forward = True
        .Wrap = wdFindContinue
'        .Format = False
'        .MatchCase = False
'        .MatchWholeWord = False
'        .MatchWildcards = False
'        .MatchSoundsLike = False
'        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

End Sub
'собираем основные пункты документа
Sub OsnovnPunkt()
'If UF1.CBFNPORPD.Value = True Then
'    ActiveDocument.Variables("punkt1-1").Value = "п.п."
'    For Each mark In AllCBp
'        If UF1.Controls.Item("CBp" & mark).Value = True Then ActiveDocument.Variables("punkt1-1").Value = ActiveDocument.Variables("punkt1-1").Value & ActiveDocument.Variables("CBp" & mark).Value
'    Next
'    ActiveDocument.Variables("punkt1-1").Value = Left(ActiveDocument.Variables("punkt1-1").Value, Len(ActiveDocument.Variables("punkt1-1").Value) - 1) & UF1.FNPORPDR.Value
'Else
'    ActiveDocument.Variables("punkt1-1").Value = "п.п. 161, 164, 169, 177, 178, 179" & UF1.FNPOPVBR.Value
'End If
'If UF1.CBFNPORPD.Value = True And UF1.CBFNPOPVB.Value = True Then
'    ActiveDocument.Variables("punkt1-1").Value = ActiveDocument.Variables("punkt1-1").Value & ";" & Strings.Chr(13) & "п.п. 161, 164, 169, 177, 178, 179" & UF1.FNPOPVBR.Value
'End If
'ActiveDocument.Variables("punkt1-1").Value = ActiveDocument.Variables("punkt1-1").Value & "."
If UF1.CBp466.Value = True Then  'если активен пункт 466 то разрешаем до 30.09
    If (DateDiff("d", UF1.AktGID.Value, ActiveDocument.Variables("CBp466data").Value)) < 0 Then
        ActiveDocument.Variables("DoNgoda") = Format(DateAdd("yyyy", Val(UF1.NaNLet.Value), ActiveDocument.Variables("CBp466data").Value), "dd.mm.yyyy")
    Else
        ActiveDocument.Variables("DoNgoda") = Format(DateAdd("yyyy", Val(UF1.NaNLet.Value) - 1, ActiveDocument.Variables("CBp466data").Value), "dd.mm.yyyy")
    End If
End If
'Индивидуальные изменения для пунктов
'Пункт 7.5.4
If UF1.CBGOST34347.Value = True Then
    ActiveDocument.Variables("punkt7-5-4Mat").Value = ActiveDocument.Variables("GOST34347PMat").Value
    If UF1.CBFNPORPD.Value = True Then ActiveDocument.Variables("punkt7-5-4Mat").Value = ActiveDocument.Variables("punkt7-5-4Mat").Value & ";" & ActiveDocument.Variables("p100FNPORPD").Value
    If UF1.CBFNPOPVB.Value = True Then ActiveDocument.Variables("punkt7-5-4Mat").Value = ActiveDocument.Variables("punkt7-5-4Mat").Value & ActiveDocument.Variables("p7-1OPVB").Value
Else
    If UF1.CBFNPORPD.Value = True Then
        ActiveDocument.Variables("punkt7-5-4Mat").Value = ActiveDocument.Variables("p100FNPORPD").Value
        If UF1.CBFNPOPVB.Value = True Then ActiveDocument.Variables("punkt7-5-4Mat").Value = ActiveDocument.Variables("punkt7-5-4Mat").Value & ActiveDocument.Variables("p7-1OPVB").Value
    Else
        If UF1.CBFNPOPVB.Value = True Then
            ActiveDocument.Variables("punkt7-5-4Mat").Value = ActiveDocument.Variables("p7-1OPVB").Value
        Else
            ActiveDocument.Variables("punkt7-5-4Mat").Value = Strings.ChrW(31)
        End If
    End If
End If
'If UF1.CBFNPORPD.Value = True And UF1.CBGOST34347.Value = True Then
'    ActiveDocument.Variables("punkt7-5-4Mat").Value = ActiveDocument.Variables("punkt7-5-4Mat").Value & ";" & ActiveDocument.Variables("p100FNPORPD").Value
'Else
'    ActiveDocument.Variables("punkt7-5-4Mat").Value = ActiveDocument.Variables("punkt7-5-4Mat").Value & ActiveDocument.Variables("p100FNPORPD").Value
'End If
If Strings.Len(ActiveDocument.Variables("punkt7-5-4Mat").Value) > 8 Then
    ActiveDocument.Variables("punkt7-5-4Mat").Value = " (материалы " & ActiveDocument.Variables("TechUsrtva").Value & " соответствуют требованиям" & ActiveDocument.Variables("punkt7-5-4Mat").Value & ")"
Else
    ActiveDocument.Variables("punkt7-5-4Mat").Value = Strings.ChrW(31)
End If

If UF1.PnIs.Value = True Then 'добавляем давления испытаний для акустики
    ActiveDocument.Variables("RazmContZon").Value = Format((3.14 * Val(UF1.odiam.Value) * (Val(UF1.odlina.Value) + Val(UF1.odiam.Value) / 2) / 1000000), "###0.0")
    ActiveDocument.Variables("RazreshaemoeP0_5").Value = Format(UF1.RazreshaemoeP.Value * 0.5, "###0.0###")
    ActiveDocument.Variables("RazreshaemoeP0_5MP").Value = Format(UF1.RazreshaemoeP.Value / 10 * 0.5, "###0.0###")
    ActiveDocument.Variables("RazreshaemoeP0_75").Value = Format(UF1.RazreshaemoeP.Value * 0.75, "###0.0##")
    ActiveDocument.Variables("RazreshaemoeP0_75MP").Value = Format(UF1.RazreshaemoeP.Value / 10 * 0.75, "###0.0###")
    ActiveDocument.Variables("RazreshaemoeP0_25").Value = Format(UF1.RazreshaemoeP.Value * 0.25, "###0.0##")
End If
End Sub

Sub ParKotl()
    ActiveDocument.Variables("ElemVodKot").Value = " (барабанов, коллекторов, труб поверхностей нагрева, необогреваемых труб в пределах котла и т.д.)"
    ActiveDocument.Variables("VIKrdSO469").Value = "; п.п. 5.4, 5.5, 5.15, 5.16, 5.17, 5.18" & UF1.SO469.Value
End Sub

Sub VodgKotl()
ActiveDocument.Variables("ElemVodKot").Value = " (коллекторов, труб поверхностей нагрева, необогреваемых труб в пределах котла и т.д.)"
End Sub

Sub ElektroKotel()
    ActiveDocument.Variables.Item("TimeGI").Value = "10 минут"
    ActiveDocument.Variables("punkt7-3NTD").Value = " п.п. 10, 22" & ActiveDocument.Variables("FNPORPDR").Value
    ActiveDocument.Variables("VIKrdSO469").Value = "; п.п. 5.4, 5.5" & UF1.SO469.Value
    ActiveDocument.Variables.Item("PassatT").Value = ", для проведения поверочного расчета на прочность использована программа " & Strings.Chr(171) & "Пассат" & Strings.Chr(187) & ", разработанная ООО " & Strings.Chr(171) & "НТП Трубопровод" & Strings.Chr(187)
End Sub

Sub Ekonomayzer()
If (ActiveDocument.Variables("ZavodIzg").Value Like "*Кус*") Or (ActiveDocument.Variables("ZavodIzg").Value Like "*Белг*") Then
    ActiveDocument.Variables("VIKrdSO469").Value = " п. 3.1.1 Приложения 9" & UF1.SO469.Value
End If
ActiveDocument.Variables("p12-1pril2").Value = Strings.ChrW(31)
ActiveDocument.Variables("tverdSO469").Value = " Приложения 1 ГОСТ 1412-85 " & Strings.Chr(171) & "Чугун с пластинчатым графитом для отливок. Марки" & Strings.Chr(187)
ActiveDocument.Variables("ISvarSoed").Value = Strings.ChrW(31)
End Sub

Sub Avtozisterna()
'ActiveDocument.Variables.Item("OzOsR").Value = "ГОСТ 34233.1-2017, ГОСТ 34233.2-2017, ГОСТ 34233.6-2017"
End Sub

Sub Gasifikator()
    ActiveDocument.Bookmarks("P7p1mat").Range.Delete ' пункт 7.1. материалы
    ActiveDocument.Variables.Item("VnuOsmotr").Value = "Наружный осмотр," & Strings.Chr(13) & "пневматическое испытание," & Strings.Chr(13) & "техническое диагностирование."
    ActiveDocument.Variables.Item("VIKrd").Value = Strings.ChrW(31)
    ActiveDocument.Variables("obmerz").Value = "Обмерзания обечайки внешнего сосуда и патрубков не обнаружено."
    ActiveDocument.Variables("KorrozPovr").Value = ", повреждения во время эксплуатации"
    ActiveDocument.Variables("KorrozPovr1").Value = Strings.ChrW(31)
    ActiveDocument.Variables("KorrozPovr2").Value = ", локальные повреждения во время эксплуатации"
End Sub

Sub VakuumSosud()
    ActiveDocument.Bookmarks("R7p4").Range.Delete
    ActiveDocument.Variables("IndxP").Value = "абс"
    ActiveDocument.Variables("GOST34347PGiI").Value = Strings.ChrW(31)
    ActiveDocument.Variables("RazreshaemoePMP").Value = "вакуум до " & Format(1 - ActiveDocument.Variables("RazreshaemoeP").Value, "###0.0#####")
    ActiveDocument.Variables("RazreshaemoeP").Value = "абсолютное " & ActiveDocument.Variables("RazreshaemoeP").Value
    ActiveDocument.Variables("DavlNeVishe").Value = " без избыточного давления (" & ActiveDocument.Variables("RazreshaemoePMP").Value & ")"
End Sub

Sub SosudPodNaliv()
    ActiveDocument.Bookmarks("R7p4").Range.Delete
    ActiveDocument.Variables("GOST34347PGiI").Value = Strings.ChrW(31)
    ActiveDocument.Variables("IspitatP").Value = "полный налив"
    ActiveDocument.Variables("TimeGI").Value = "4 часов"
    ActiveDocument.Variables("PadDavl").Value = Strings.ChrW(31)
    ActiveDocument.Variables("PodRabDav").Value = Strings.ChrW(31)
    ActiveDocument.Variables("DavlNeVishe").Value = " без избыточного давления (под налив)"
End Sub

Sub TruboprovPara()
    ActiveDocument.Variables("Izgotovitel6").Value = "Монтажная организация"
    ActiveDocument.Variables("SostElTrub").Value = "Дефекты трассы и опорно-подвесной системы отсутствуют, повреждений изоляции, и её внешнего кожуха, следов намокания, пропуска среды, не обнаружено. "
    ActiveDocument.Bookmarks("R7p4").Range.Delete
    ActiveDocument.Variables("TempDeform").Value = ", температурными деформациями"
    ActiveDocument.Variables("punkt7-5-4Mat").Value = Strings.ChrW(31)
End Sub

Sub NGUCGUUDH()
'    ActiveDocument.Variables("ovalnrd").Value = ActiveDocument.Variables("ovalnrd").Value & "; п. 5.4.3.2" & UF1.RD26_260.Value
End Sub

Sub variable()
'ActiveDocument.Variables("FNPOPVB").Value = " Федеральных норм и правил в области промышленной безопасности " & Strings.Chr(171) & "Общие правила взрывобезопасности для взрывопожароопасных химических, нефтехимических и нефтеперерабатывающих производств" & Strings.Chr(187) & " утвержденных приказом Федеральной службы по экологическому, технологическому и атомному надзору №533 от 15.12.2020 г., зарегистрированных в Минюсте России 25.12.2020 г., рег.№61808"
ActiveDocument.Variables("RazreshaemoePMP").Value = "1,0 МПа"  'Strings.ChrW(31)
'MsgBox (ActiveDocument.Variables("FNPOPVB").Value)
End Sub

Sub BallGroUst()
    ActiveDocument.Variables("GOST34347PMat").Value = UF1.GOST9731.Value
    ActiveDocument.Variables.Item("PassatT").Value = Strings.ChrW(31)
    ActiveDocument.Variables.Item("OzOsR").Value = UF1.GOST25215.Value
End Sub

Sub MnogCHislo()
ActiveDocument.Variables("MnNum1").Value = "технические устройства"
ActiveDocument.Variables("MnNum2").Value = "применяемые"
ActiveDocument.Variables("MnNum3").Value = "установленные"
ActiveDocument.Variables("MnNum4").Value = "технических устройств"
ActiveDocument.Variables("MnNum5").Value = "отработавших"
ActiveDocument.Variables("MnNum6").Value = "Заводские номера"
ActiveDocument.Variables("MnNum7").Value = "Регистрационные номера"
ActiveDocument.Variables("MnNum8").Value = "техническим устройствам"
ActiveDocument.Variables("No").Value = UF1.poleRegNum.Value & "№№"
ActiveDocument.Variables("ZavNo").Value = "№№"

End Sub

Sub QuickSort(coll As Collection, first As Long, last As Long)
Dim vCentreVal As Variant, vTemp As Variant
Dim lTempLow As Long
Dim lTempHi As Long
lTempLow = first
lTempHi = last
vCentreVal = coll((first + last) \ 2)
Do While lTempLow <= lTempHi
Do While coll(lTempLow) < vCentreVal And lTempLow < last
    lTempLow = lTempLow + 1
Loop
Do While vCentreVal < coll(lTempHi) And lTempHi > first
    lTempHi = lTempHi - 1
Loop
If lTempLow <= lTempHi Then ' Поменять значения
vTemp = coll(lTempLow)
coll.Add coll(lTempHi), After:=lTempLow
coll.Remove lTempLow
coll.Add vTemp, Before:=lTempHi
coll.Remove lTempHi + 1
lTempLow = lTempLow + 1
lTempHi = lTempHi - 1
End If
Loop
If first < lTempHi Then QuickSort coll, first, lTempHi
If lTempLow < last Then QuickSort coll, lTempLow, last
End Sub

'удаление ссылок
Sub UnlinkBookmarks(A1 As String)
S1 = 0
NameBookmarks = A1 & S1
Do While ActiveDocument.Bookmarks.Exists(NameBookmarks) = True
ActiveDocument.Bookmarks(NameBookmarks).Range.Fields.Unlink
S1 = S1 + 1
NameBookmarks = A1 & S1
Loop
End Sub

'удаление содержимого закладок
Sub DeleteBookmarks(A1 As String)
S1 = 0
NameBookmarks = A1 & S1
Do While ActiveDocument.Bookmarks.Exists(NameBookmarks) = True
ActiveDocument.Bookmarks(NameBookmarks).Range.Delete
S1 = S1 + 1
NameBookmarks = A1 & S1
Loop
End Sub

'изменение текста закладок
Sub SetBookmark(NameBookmarks As String, ValueBookmarks As String)
    Set TTMP = ActiveDocument.Bookmarks(NameBookmarks).Range
    TTMP.Text = ValueBookmarks
    ActiveDocument.Bookmarks.Add Name:=NameBookmarks, Range:=TTMP
End Sub

Sub ClearAllF()
UF1.CBFNPORPD.Value = False
UF1.CBFNPOPVB.Value = False
UF1.CBSO439.Value = False
UF1.CBSO464.Value = False
UF1.CBSO469.Value = False
UF1.CBGOST34347.Value = False
UF1.CBRD2626012.Value = False
UF1.CBVM030104.Value = False
UF1.CBRUA93.Value = False
End Sub

