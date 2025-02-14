Attribute VB_Name = "IndIz"
'Индивидуальные изменения для каждого техустройства

Sub ParKotl()
'    ActiveDocument.Variables("punkt7-3NTD").Value = " п." & ActiveDocument.Variables("CBp10").Value & UF1.FNPORPDR.Value
    ActiveDocument.Variables("ElemVodKot").Value = " (барабанов, коллекторов, труб поверхностей нагрева, необогреваемых труб в пределах котла и т.д.)"
    ActiveDocument.Variables("VIKrdSO469").Value = "; п.п. 5.4, 5.5, 5.15, 5.16, 5.17, 5.18" & UF1.SO469.Value
End Sub

Sub VodgKotl()
'    ActiveDocument.Variables("punkt7-3NTD").Value = " п." & ActiveDocument.Variables("CBp10").Value & UF1.FNPORPDR.Value
    ActiveDocument.Variables("ElemVodKot").Value = " (коллекторов, труб поверхностей нагрева, необогреваемых труб в пределах котла и т.д.)"
End Sub

Sub ElektroKotel()
    ActiveDocument.Variables.Item("TimeGI").Value = "10 минут"
'    ActiveDocument.Variables("punkt7-3NTD").Value = " п.п. 10, 22" & UF1.FNPORPDR.Value
    ActiveDocument.Variables("VIKrdSO469").Value = "; п.п. 5.4, 5.5" & UF1.SO469.Value
    ActiveDocument.Variables.Item("PassatT").Value = ", для проведения поверочного расчета на прочность использована программа " & Strings.Chr(171) & "Пассат" & Strings.Chr(187) & ", разработанная ООО " & Strings.Chr(171) & "НТП Трубопровод" & Strings.Chr(187)
End Sub

Sub Ekonomayzer()
If (ActiveDocument.Variables("ZavodIzg").Value Like "*Кус*") Or (ActiveDocument.Variables("ZavodIzg").Value Like "*Белг*") Then
    If (ActiveDocument.Variables("RabSreda").Value Like "*[Гг]аз*") Or (ActiveDocument.Variables("RabSreda").Value = Strings.ChrW(31)) = 0 Then
        ActiveDocument.Variables("VIKrdSO469").Value = "; п. 3.2.1 Приложения 9" & UF1.SO469.Value
    Else
        ActiveDocument.Variables("VIKrdSO469").Value = "; п. 3.1.1 Приложения 9" & UF1.SO469.Value
    End If
End If
'    ActiveDocument.Variables("punkt7-3NTD").Value = " п." & ActiveDocument.Variables("CBp10").Value & UF1.FNPORPDR.Value
    ActiveDocument.Variables("p12-1pril2").Value = " п.п. 2, 3, 4, 5 Приложения №8" & UF1.FNPORPDR.Value
    ActiveDocument.Variables("tverdSO469").Value = " Приложения 1 ГОСТ 1412-85 " & Strings.Chr(171) & "Чугун с пластинчатым графитом для отливок. Марки" & Strings.Chr(187)
    ActiveDocument.Variables("ISvarSoed").Value = " металла"
    ActiveDocument.Variables("svarnih").Value = Strings.ChrW(31)
    ActiveDocument.Variables("korpusa").Value = " элементов экономайзера"
    ActiveDocument.Variables("ObechBarKotl").Value = "ребристой трубы экономайзера"
    ActiveDocument.Variables("Punkt3211RD1024998").Value = "3.3.1.1."
End Sub

Sub Vozduhosbornik()
End Sub

Sub Avtozisterna()
    ActiveDocument.Variables("KIPiA").Value = "Оснащение " & ActiveDocument.Variables("TechUsrtva").Value & " устройствами безопасности и средствами КИПиА соответствует"
End Sub

Sub Podogrevatel()
'    ActiveDocument.Variables("TechUsrtva").Value = "подогревателя"
'    ActiveDocument.Variables("TechUsrtvo").Value = "подогреватель"
'    ActiveDocument.Variables("TechUsrtvoB").Value = UCase(Left(ActiveDocument.Variables("TechUsrtvo").Value, 1)) & LCase(Mid(ActiveDocument.Variables("TechUsrtvo").Value, 2)) 'Делаем первую букву большой
End Sub

Sub Gasifikator()
    ActiveDocument.Bookmarks("P7p1mat").Range.Delete ' пункт 7.1. материалы
    ActiveDocument.Variables.Item("VnuOsmotr").Value = "Наружный осмотр," & Strings.Chr(13) & "пневматическое испытание," & Strings.Chr(13) & "техническое диагностирование."
    ActiveDocument.Variables.Item("VIKrd").Value = Strings.ChrW(31)
    ActiveDocument.Variables("obmerz").Value = "Обмерзания обечайки внешнего сосуда и патрубков не обнаружено."
    ActiveDocument.Variables("KorrozPovr").Value = ", повреждения во время эксплуатации"
    ActiveDocument.Variables("KorrozPovr1").Value = Strings.ChrW(31)
    ActiveDocument.Variables("KorrozPovr2").Value = ", локальные повреждения во время эксплуатации"
    ActiveDocument.Variables("DopDoc").Value = Strings.Chr(13) & "- акт на проведение замера вакуума и натекания в теплоизоляционной полости криогенного газификатора до проведения пневмоиспытания, выданный АО " & Strings.Chr(171) & "Сибтехгаз им. Кима Ф.И." & Strings.Chr(187) & " от " & UF1.AktGID.Value & " г. - 1 л.;"
    ActiveDocument.Variables("DopDoc").Value = ActiveDocument.Variables("DopDoc").Value & Strings.Chr(13) & "- акт на проведение замера вакуума и натекания после пневмоиспытания криогенного газификатора ГХК-3/1,6-200М, зав.№" & UF1.ZavN.Value & ", выданный АО " & Strings.Chr(171) & "Сибтехгаз им. Кима Ф.И." & Strings.Chr(187) & " от " & UF1.AktGID.Value & " г. - 1 л.;"
    ActiveDocument.Variables("DopDoc").Value = ActiveDocument.Variables("DopDoc").Value & Strings.Chr(13) & "- акт на обезжиривание криогенного газификатора ГХК-3/1,6-200М, выданный АО " & Strings.Chr(171) & "Сибтехгаз им. Кима Ф.И." & Strings.Chr(187) & " от " & UF1.AktGID.Value & " г. - 1 л.;"
    ActiveDocument.Variables("DopDoc").Value = ActiveDocument.Variables("DopDoc").Value & Strings.Chr(13) & "- акт на настройку предохранительных клапанов на ГХК-3/1,6-200М зав. №" & UF1.ZavN.Value & ", выданный АО " & Strings.Chr(171) & "Сибтехгаз им. Кима Ф.И." & Strings.Chr(187) & " от " & UF1.AktGID.Value & " г. - 1 л.;"
End Sub

Sub VakuumSosud()
'    ActiveDocument.Bookmarks("R7p4").Range.Delete
    ActiveDocument.Variables("IndxP").Value = "абс"
    ActiveDocument.Variables("GOST34347PGiI").Value = Strings.ChrW(31)
    ActiveDocument.Variables("RazreshaemoeVKM").Value = "вакуум до " & Format((1 - CDbl(UF1.RazreshaemoeP.Value)) / 10, "###0.0#####") & "(" & Format(1 - CDbl(UF1.RazreshaemoeP.Value), "###0.0#####") & ")"
'    ActiveDocument.Variables("RazreshaemoeP").Value = "абсолютное " & ActiveDocument.Variables("RazreshaemoeP").Value
    ActiveDocument.Variables("DavlNeVishe").Value = " без избыточного давления (вакуум до " & Format(1 - CDbl(UF1.RazreshaemoeP.Value), "###0.0#####") & " кгс/см" & Strings.ChrW(178) & ")"
End Sub

Sub SosudPodNaliv()
'    ActiveDocument.Bookmarks("R7p4").Range.Delete
'    Call DeleteBookmarks("zikl") ' циклы
    ActiveDocument.Variables("GOST34347PGiI").Value = Strings.ChrW(31)
    ActiveDocument.Variables("IspitatP").Value = "полный налив"
    ActiveDocument.Variables("TimeGI").Value = "4 часов"
    ActiveDocument.Variables("PadDavl").Value = Strings.ChrW(31)
    ActiveDocument.Variables("PodRabDav").Value = Strings.ChrW(31)
    ActiveDocument.Variables("DavlNeVishe").Value = " без избыточного давления (под налив)"
    ActiveDocument.Variables("p2pr8RD3417302").Value = Strings.ChrW(31)
End Sub

Sub SosudHOPO()
'    ActiveDocument.Variables("TechUsrtva").Value = "бака"
'    ActiveDocument.Variables("TechUsrtvo").Value = "бак"
'    Call DeleteBookmarks("zikl") ' циклы
'    ActiveDocument.Variables("punkt7-3-1").Value = "Размещение "
    ActiveDocument.Variables("punkt7-3").Value = Strings.ChrW(31)
    ActiveDocument.Variables("KIPiA").Value = "Оснащение " & ActiveDocument.Variables("TechUsrtva").Value & " соответствует"
    ActiveDocument.Variables("VnutrIzbP").Value = "гидростатическим"
    ActiveDocument.Variables("VnutrP").Value = "гидростатического"
    ActiveDocument.Variables("punkt7-5-4Mat").Value = Strings.ChrW(31)
    ActiveDocument.Variables.Item("PassatT").Value = Strings.ChrW(31)
'    ActiveDocument.Variables.Item("OzOsR").Value = "п. 11.10" & UF1.RD1533413752696.Value
    ActiveDocument.Variables("ProbnDavlen").Value = "наливом"
    
End Sub

Sub TruboprovPara()
    ActiveDocument.Variables("Izgotovitel6").Value = "Монтажная организация"
    ActiveDocument.Variables("SostElTrub").Value = "Дефекты трассы и опорно-подвесной системы отсутствуют, повреждений изоляции, и её внешнего кожуха, следов намокания, пропуска среды, не обнаружено. "
    If ActiveDocument.Bookmarks.Exists("Ne_peremesh") = True Then ActiveDocument.Bookmarks("Ne_peremesh").Range.Delete
    ActiveDocument.Variables("TempDeform").Value = ", температурными деформациями"
    ActiveDocument.Variables("punkt7-5-4Mat").Value = Strings.ChrW(31)
End Sub

Sub NGUCGUUDH()
'    ActiveDocument.Variables("ovalnrd").Value = ActiveDocument.Variables("ovalnrd").Value & "; п. 5.4.3.2" & UF1.RD26_260.Value
End Sub

Sub BakKislota()
'    Call DeleteBookmarks("zikl") ' циклы
'    Call DeleteBookmarks("TehnichUstr") ' удаляем пункты про техустройства
    ActiveDocument.Bookmarks("VikRezTruboprov").Range.Delete ' ВИК результаты осмотра трубопровода
    ActiveDocument.Variables("punkt7-3").Value = Strings.ChrW(31)
    ActiveDocument.Variables("KIPiA").Value = "Оснащение " & ActiveDocument.Variables("TechUsrtva").Value & " соответствует"
    ActiveDocument.Variables("VnutrIzbP").Value = "гидростатическим"
    ActiveDocument.Variables("punkt7-5-4Mat").Value = Strings.ChrW(31)
    ActiveDocument.Variables.Item("PassatT").Value = Strings.ChrW(31)
    ActiveDocument.Variables("ProbnDavlen").Value = "наливом"
    ActiveDocument.Variables("SvedOMerPFNP").Value = "26, 27"
End Sub

Sub TehnTruboprovod()
    Call DeleteBookmarks("Rezervuar") ' удаляем часть про резервуары
'    ActiveDocument.Bookmarks("KotlObor").Range.Delete ' пункт 7.3. про оборудование котла
'    ActiveDocument.Bookmarks("R7p4").Range.Delete 'Пункт про КИПиА
    ActiveDocument.Bookmarks("VikRezTruboprov").Range.Delete ' ВИК результаты осмотра трубопровода в холодном и горячем состоянии
    ActiveDocument.Tables(1).Cell(2, 2).Range.Text = "Труба"
    ActiveDocument.Tables(1).Cell(3, 2).Range.Text = "Труба"
    If ActiveDocument.Variables("UZKRekTT").Value = Strings.ChrW(31) Then ActiveDocument.Variables("UZKRekTT").Value = ". Дефектов, по результатам контроля сварных соединений технологического трубопровода в соответствии с ГОСТ Р 55724-2013 " & _
    Strings.Chr(171) & "Контроль неразрушающий. Соединения сварные. Методы ультразвуковые" & Strings.Chr(187) & ", не зафиксировано."
'    ActiveDocument.Variables("MnNum3").Value = "применяемый"
'    ActiveDocument.Variables("MnNum4").Value = "сооружения"
'    ActiveDocument.Variables("MnNum8").Value = "сооружению"
    Call VarSoorugenie
'    ActiveDocument.Variables("TechDiagn").Value = "обследование"
'    ActiveDocument.Variables("TechDiagnB").Value = "Обследование"
'    ActiveDocument.Variables("tehdiagnnk").Value = "обследования"
'    ActiveDocument.Variables("tehdiagn").Value = "обследования"
'    ActiveDocument.Variables("tehdiagn1").Value = "обследования"
'    ActiveDocument.Variables("Izgotovitel6").Value = "Монтажная организация"
    ActiveDocument.Variables("TechTr0").Value = ", поверочного расчета на прочность, оценки остаточной несущей способности и пригодности к дальнейшей эксплуатации"
    ActiveDocument.Variables("punkt7-3-1").Value = "Размещение, прокладка, устройство и оснащение "
    ActiveDocument.Variables("punkt7-3").Value = Strings.ChrW(31)
    ActiveDocument.Variables("SvedOMer0").Value = "26, 27" 'Strings.ChrW(31)
    ActiveDocument.Variables("SvedOMerPFNP").Value = "24, 25, 26, 27"
'    ActiveDocument.Variables("NaznTehUstr").Value = ActiveDocument.Variables("NaznTehUstr").Value & ActiveDocument.Variables("RabSreda").Value
    ActiveDocument.Variables("SostElTrub").Value = "Пространственное положение трубопровода соответствует паспортной схеме. Фактические сечения элементов трубопровода соответствуют паспортной схеме. Определена степень коррозии элементов трубопровода. "
    ActiveDocument.Variables("TTrDop").Value = " Степень влияния гидрологических, аэрологических и атмосферных воздействий на трубопровод определена как незначительная. Изучена химическая агрессивность производственной среды в отношении материалов трубопровода."
    ActiveDocument.Variables("TTrDop3").Value = Strings.Chr(13) & "Определена фактическая прочность материалов трубопровода. На основании анализа эксплуатационно-технической документации, проведенных визуально-измерительного и других методов неразрушающего контроля сделан вывод: фактическая прочность материалов соответствует проектным параметрам."
    ActiveDocument.Variables("TTrDop2").Value = " оценки остаточной несущей способности и пригодности трубопровода к дальнейшей эксплуатации,"
    ActiveDocument.Variables("TTrDop4").Value = ", определение соответствия конструкций проектной документации и требованиям нормативных документов, выявление дефектов и повреждений элементов и узлов, определение пространственного положения конструкций, их фактических сечений и состояния соединений, определение степени коррозии металлических элементов"
    ActiveDocument.Variables("TTrDop5").Value = " Определение степени влияния гидрологических, аэрологических и атмосферных воздействий. Изучение химической агрессивности производственной среды в отношении материалов " & ActiveDocument.Variables("TechUsrtva").Value & "."
    ActiveDocument.Variables("TTrDop5Som").Value = ";" & Strings.Chr(13) & vbTab & "- определение степени влияния гидрологических, аэрологических и атмосферных воздействий;" & Strings.Chr(13) & vbTab & "- изучение химической агрессивности производственной среды в отношении материалов " & ActiveDocument.Variables("TechUsrtva").Value
    ActiveDocument.Variables("TTrDop6").Value = " Определение фактической прочности материалов в сравнении с проектными параметрами."
    ActiveDocument.Variables("TTrDop6SoM").Value = ";" & Strings.Chr(13) & vbTab & "- определение фактической прочности материалов в сравнении с проектными параметрами"
    ActiveDocument.Variables("TTrDop7").Value = " Поверочный расчет конструкций с учетом выявленных при обследовании отклонений, дефектов и повреждений, фактических нагрузок."
    ActiveDocument.Variables("TTrDop7SoM").Value = ";" & Strings.Chr(13) & vbTab & "- поверочный расчет конструкций с учетом выявленных при обследовании отклонений, дефектов и повреждений, фактических нагрузок"
    ActiveDocument.Variables("TTrDop8").Value = " Оценка остаточной несущей способности и пригодности трубопровода к дальнейшей эксплуатации."
    ActiveDocument.Variables("TTrDop8SoM").Value = ";" & Strings.Chr(13) & vbTab & "- оценка остаточной несущей способности и пригодности трубопровода к дальнейшей эксплуатации"
    ActiveDocument.Variables("P71ProekDok").Value = " По предоставленным данным, трубопровод смонтирован в соответствии с проектной документацией."
    ActiveDocument.Variables.Item("PassatT").Value = Strings.ChrW(31)
    ActiveDocument.Variables("p2pr8RD3417302").Value = Strings.ChrW(31)
    ActiveDocument.Variables("ProbnDavlen").Value = "на прочность и плотность,"
    If UF1.CBtt100.Value = True Then ActiveDocument.Variables("ProbnDavlen").Value = ActiveDocument.Variables("ProbnDavlen").Value & " и пневматические испытания на герметичность,"
    ActiveDocument.Variables("TTPoverRash").Value = " поверочного расчета для"
    ActiveDocument.Variables("korpusa").Value = Strings.ChrW(31)
    ActiveDocument.Variables("CBp466-1").Value = " и п. 177" & UF1.FNPTehnTrub.Value
    If ActiveDocument.Bookmarks.Exists("Ne_peremesh") = True Then ActiveDocument.Bookmarks("Ne_peremesh").Range.Delete
   
End Sub
Sub TruboprovodKislota()
    Call DeleteBookmarks("Rezervuar") ' удаляем часть про резервуары
'    ActiveDocument.Bookmarks("KotlObor").Range.Delete ' пункт 7.3. про оборудование котла
'    ActiveDocument.Variables("MnNum1").Value = "сооружение"
'    ActiveDocument.Variables("MnNum2").Value = "применяемый"
'    ActiveDocument.Variables("MnNum3").Value = "применяемый"
    Call VarSoorugenie
    ActiveDocument.Variables("TechTr0").Value = ", поверочного расчета и оценки остаточной несущей способности и пригодности к дальнейшей эксплуатации"
    ActiveDocument.Variables("punkt7-3-1").Value = "Размещение, прокладка, устройство и оснащение "
    ActiveDocument.Variables("punkt7-3").Value = Strings.ChrW(31)
    ActiveDocument.Variables("SvedOMer0").Value = "25, 27" 'Strings.ChrW(31)
    ActiveDocument.Variables("SvedOMerPFNP").Value = "24, 25, 26, 27"
'    ActiveDocument.Variables("NaznTehUstr").Value = ActiveDocument.Variables("NaznTehUstr").Value & ActiveDocument.Variables("RabSreda").Value
    ActiveDocument.Variables("SostElTrub").Value = "Пространственное положение и фактические сечения элементов трубопровода соответствуют паспортной схеме. Определена степень коррозии элементов трубопровода. "
    ActiveDocument.Variables("TTrDop").Value = " Степень влияния гидрологических, аэрологических и атмосферных воздействий на трубопровод определена как незначительная. Изучена химическая агрессивность производственной среды в отношении материалов трубопровода."
    ActiveDocument.Variables("TTrDop1").Value = " Техническое состояние трубопровода соответствует требованиям п.п. 30, 169" & UF1.FNPOPVBR.Value & "."
    ActiveDocument.Variables("TTrDop3").Value = Strings.Chr(13) & "Определена фактическая прочность материалов трубопровода. На основании анализа эксплуатационно-технической документации, проведенных визуально-измерительного и других методов неразрушающего контроля сделан вывод: фактическая прочность материалов соответствует проектным параметрам."
    ActiveDocument.Variables("TTrDop2").Value = " оценки остаточной несущей способности и пригодности трубопровода к дальнейшей эксплуатации,"
    ActiveDocument.Variables.Item("PassatT").Value = Strings.ChrW(31)
    ActiveDocument.Variables("p2pr8RD3417302").Value = Strings.ChrW(31)
    ActiveDocument.Variables("ProbnDavlen").Value = "на прочность и плотность,"
    ActiveDocument.Variables("korpusa").Value = Strings.ChrW(31)
    ActiveDocument.Variables("CBp466-1").Value = " и п. 177" & UF1.FNPTehnTrub.Value
End Sub
Sub RezervuarMazut()
'    ActiveDocument.Variables("MnNum2").Value = "применяемый"
'    ActiveDocument.Variables("MnNum3").Value = "применяемый"
'    ActiveDocument.Variables("MnNum1").Value = "сооружение"
'    ActiveDocument.Variables("MnNum4").Value = "сооружения"
'    ActiveDocument.Variables("MnNum8").Value = "сооружению"
    Call VarSoorugenie
'    ActiveDocument.Variables("TechDiagn").Value = "обследование"
'    ActiveDocument.Variables("TechDiagnB").Value = "Обследование"
'    ActiveDocument.Variables("tehdiagnnk").Value = "обследования"
'    ActiveDocument.Variables("tehdiagn").Value = "обследования"
'    ActiveDocument.Variables("tehdiagn1").Value = "обследования"
'    ActiveDocument.Variables("Izgotovitel6").Value = "Монтажная организация"
    ActiveDocument.Variables("PrSrSl-PBSNN").Value = Strings.Chr(13) & "Продление срока службы резервуара осуществляется в соответствии с требованиями п. 261" & UF1.FNPPESNN.Value
End Sub

Sub BallGroUst()
    Call MnogCHislo
'    ActiveDocument.Variables("TechUsrtva").Value = "баллонов"
'    ActiveDocument.Variables("TechUsrtvo").Value = "баллоны групповой установки"
'    ActiveDocument.Variables("TechUsrtvoB").Value = UCase(Left(ActiveDocument.Variables("TechUsrtvo").Value, 1)) & LCase(Mid(ActiveDocument.Variables("TechUsrtvo").Value, 2)) 'Делаем первую букву большой
    ActiveDocument.Variables("GOST34347PMat").Value = UF1.GOST9731.Value
    ActiveDocument.Variables.Item("PassatT").Value = Strings.ChrW(31)
End Sub
Sub Ballon()
'    ActiveDocument.Variables("TechUsrtva").Value = "баллонов"
'    ActiveDocument.Variables("TechUsrtvo").Value = "баллон"
'    ActiveDocument.Variables("TechUsrtvoB").Value = UCase(Left(ActiveDocument.Variables("TechUsrtvo").Value, 1)) & LCase(Mid(ActiveDocument.Variables("TechUsrtvo").Value, 2)) 'Делаем первую букву большой
    ActiveDocument.Variables("GOST34347PMat").Value = UF1.GOST9731.Value
    ActiveDocument.Variables.Item("PassatT").Value = Strings.ChrW(31)
End Sub

Sub Nasos()
    ActiveDocument.Tables(2).Cell(Row:=2, Column:=1).Range = "Напор, м.вод.ст."
    ActiveDocument.Tables(2).Cell(Row:=2, Column:=2).Range = UF1.RaschetnP.Value
    ActiveDocument.Tables(2).Cell(Row:=3, Column:=1).Range = "Производительность, м" & Strings.ChrW(179) & "/ч."
    ActiveDocument.Variables("RaschSreda").Value = Replace(ActiveDocument.Variables("RaschSreda").Value, ",", ".")
    ActiveDocument.Variables("PasportPar").Value = "Основные технические характеристики"
    ActiveDocument.Variables("RabSredaToplRasch").Value = Strings.Chr(13) & ActiveDocument.Variables("RabSredaToplRasch").Value
    ActiveDocument.Variables("RaschSreda").Value = ActiveDocument.Variables("RaschSreda").Value & Strings.Chr(13) & "Напор:"
    ActiveDocument.Variables("RaschetnP").Value = UF1.RaschetnP.Value & " м.вод.ст." & Strings.Chr(13) & "Производительность: "
    ActiveDocument.Variables("Raschetnt").Value = Trim(UF1.Raschetnt.Value) & " м" & Strings.ChrW(179) & "/ч."
    ActiveDocument.Variables("VKorp").Value = Strings.Chr(13) & "Мощность электродвигателя: " & UF1.VKorp.Value & " кВт"
    If ActiveDocument.Bookmarks.Exists("ORPD10") = True Then ActiveDocument.Bookmarks("ORPD10").Range.Delete
'    ActiveDocument.Variables("NaznTehUstr").Value = ActiveDocument.Variables("NaznTehUstr").Value & ActiveDocument.Variables("RabSreda").Value
'    ActiveDocument.Variables("punkt7-3").Value = Strings.ChrW(31)
    ActiveDocument.Variables("korpusa").Value = " элементов насосной установки"
    ActiveDocument.Variables("ISvarSoed").Value = Strings.ChrW(31)
    ActiveDocument.Variables("p191PBETT").Value = " таблицы 1 " & UF1.M2_96.Value
    If UF1.CBFNPHOPO.Value = True Then ActiveDocument.Variables("TTrDop3").Value = " Состояние насоса соответствует требованиям п.п. 15, 132" & UF1.FNPHOPO.Value & "."
    ActiveDocument.Variables("RezNKPP").Value = "Техническое состояние " & ActiveDocument.Variables("TechUsrtva").Value & " оценивается как работоспособное при малой вероятности отказа, допустима длительная работа агрегата."
    ActiveDocument.Variables("M2_96").Value = Strings.Chr(13) & UF1.M2_96.Value
    ActiveDocument.Variables("GOST32106").Value = UF1.GOST32106.Value
    ActiveDocument.Variables("OORNasos").Value = Strings.Chr(13) & "Проведен анализ режимов работы и исследование напряженно-деформированного состояния (учтено: отсутствие видимых деформаций корпуса; геометрические параметры элементов насоса)."
    ActiveDocument.Bookmarks("OORAnPr").Range.Delete
    ActiveDocument.Variables("OORNasos").Value = ActiveDocument.Variables("OORNasos").Value & Strings.Chr(13) & "При вибродиагностировании насоса получены оценки " & Strings.Chr(171) & "Хорошо" & Strings.Chr(187) & " и " & Strings.Chr(171) & "Допустимо" & Strings.Chr(187) & " (допустима дальнейшая длительная эксплуатация), в соответствии с требованиями п. 6.1 и таблицы А.1 приложения А " & UF1.GOST32106.Value
    ActiveDocument.Variables("OORNasos").Value = ActiveDocument.Variables("OORNasos").Value & "; требованиями п. 4.4.2 и таблицы Б.1 приложения Б " & UF1.SA03_001_05.Value
    ActiveDocument.Variables("OORNasos").Value = ActiveDocument.Variables("OORNasos").Value & "." & Strings.Chr(13) & "В соответствии с требованиями " & UF1.M2_96.Value & ", регламентирующей виды, периодичность и содержание технического обслуживания, период эксплуатации насоса между капитальными ремонтами, составляет не более 30000 часов. С учетом общего технического состояния насоса, возможно продление срока службы насоса, при действующем режиме эксплуатации, на "
    ActiveDocument.Variables("OORNasos").Value = ActiveDocument.Variables("OORNasos").Value & ActiveDocument.Variables("NaNLet").Value & "."
    ActiveDocument.Variables("VIKRezKontr").Value = "при визуальном и измерительном контроле насоса, механических и коррозионных повреждений, усталостных трещин и других видимых дефектов, препятствующих дальнейшей эксплуатации, не обнаружено"
End Sub


