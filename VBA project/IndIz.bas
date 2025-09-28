Attribute VB_Name = "IndIz"
'Индивидуальные изменения для каждого техустройства

Sub ParKotl()
'    ActiveDocument.Variables("punkt7-3NTD").Value = " п." & ActiveDocument.Variables("CBp10").Value & UF1.FNPORPDR.Value
    ActiveDocument.Variables("korpusa").Value = " элементов котла"
    ActiveDocument.Variables("CBp466-1").Value = " и п. 6.6" & UF1.SO469.Value
    ActiveDocument.Variables("ElemVodKot").Value = " (барабанов, коллекторов, труб поверхностей нагрева, необогреваемых труб в пределах котла и т.д.)"
    ActiveDocument.Variables("VIKrdSO469").Value = "; п.п. 5.4, 5.5, 5.15, 5.16, 5.17, 5.18" & UF1.SO469.Value
'    ActiveDocument.Tables(2).Rows(3).Delete
End Sub

Sub VodgKotl()
'    ActiveDocument.Variables("punkt7-3NTD").Value = " п." & ActiveDocument.Variables("CBp10").Value & UF1.FNPORPDR.Value
    ActiveDocument.Variables("korpusa").Value = " элементов котла"
    ActiveDocument.Variables("CBp466-1").Value = " и п. 6.6" & UF1.SO469.Value
    ActiveDocument.Variables("ElemVodKot").Value = " (коллекторов, труб поверхностей нагрева, необогреваемых труб в пределах котла и т.д.)"
End Sub

Sub ElektroKotel()
    ActiveDocument.Variables.item("TimeGI").Value = "10 минут"
'    ActiveDocument.Variables("punkt7-3NTD").Value = " п.п. 10, 22" & UF1.FNPORPDR.Value
    ActiveDocument.Variables("VIKrdSO469").Value = "; п.п. 5.4, 5.5" & UF1.SO469.Value
    ActiveDocument.Variables.item("PassatT").Value = ", для проведения поверочного расчета на прочность использована программа " & Strings.Chr(171) & "Пассат" & Strings.Chr(187) & ", разработанная ООО " & Strings.Chr(171) & "НТП Трубопровод" & Strings.Chr(187)
    ActiveDocument.Variables("CBp466-1").Value = " и п. 6.6" & UF1.SO469.Value
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
    ActiveDocument.Variables("CBp466-1").Value = " и п. 6.6" & UF1.SO469.Value
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
    ActiveDocument.Variables("punkt7-3-1").Value = "Установка, размещение и оснащение "
    ActiveDocument.Variables.item("VnuOsmotr").Value = "Наружный осмотр," & Strings.Chr(13) & "пневматическое испытание," & Strings.Chr(13) & "техническое диагностирование."
    ActiveDocument.Variables.item("VIKrd").Value = Strings.ChrW(31)
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
    ActiveDocument.Variables("RazreshaemoeVKM").Value = "вакуум до " & Format((1 - CDbl(UF1.RazreshaemoeP.Value)) / 10, "0.0#####") & "(" & Format(1 - CDbl(UF1.RazreshaemoeP.Value), "0.0#####") & ")"
'    ActiveDocument.Variables("RazreshaemoeP").Value = "абсолютное " & ActiveDocument.Variables("RazreshaemoeP").Value
    ActiveDocument.Variables("DavlNeVishe").Value = " без избыточного давления (вакуум до " & Format(1 - CDbl(UF1.RazreshaemoeP.Value), "0.0#####") & " кгс/см" & Strings.ChrW(178) & ")"
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
    ActiveDocument.Variables("punkt7-3").Value = Strings.ChrW(31)
    ActiveDocument.Variables("KIPiA").Value = "Оснащение " & ActiveDocument.Variables("TechUsrtva").Value & " соответствует"
    ActiveDocument.Variables("VnutrIzbP").Value = "гидростатическим"
    ActiveDocument.Variables("VnutrP").Value = "гидростатического"
    ActiveDocument.Variables("punkt7-5-4Mat").Value = Strings.ChrW(31)
    ActiveDocument.Variables.item("PassatT").Value = Strings.ChrW(31)
    ActiveDocument.Variables("ProbnDavlen").Value = "наливом"
    ActiveDocument.Tables(1).Cell(Row:=1, Column:=4).Range = "Ширина, мм"
    ActiveDocument.Variables("RabTempP6").Value = ActiveDocument.Variables("RabTempP6").Value & ", высота налива - " & Trim(UF1.RabocheePRub.Value) & " мм"
    ActiveDocument.Tables(2).Cell(Row:=3, Column:=1).Range = "Высота налива, мм"
    ActiveDocument.Tables(2).Cell(Row:=3, Column:=2).Range = "до " & UF1.RabocheePRub.Value
End Sub

Sub TruboprovPara()
    ActiveDocument.Variables("Izgotovitel6").Value = "Монтажная организация"
    ActiveDocument.Variables("SostElTrub").Value = "Дефекты трассы и опорно-подвесной системы отсутствуют, повреждений изоляции, и её внешнего кожуха, следов намокания, пропуска среды, не обнаружено. "
    If ActiveDocument.Bookmarks.Exists("Ne_peremesh") = True Then ActiveDocument.Bookmarks("Ne_peremesh").Range.Delete
    ActiveDocument.Variables("TempDeform").Value = ", температурными деформациями"
    ActiveDocument.Variables("punkt7-5-4Mat").Value = Strings.ChrW(31)
    ActiveDocument.Variables("DataIzg6").Value = "Дата монтажа"
    ActiveDocument.Variables("VIKRezKontr").Value = "для оценки общего состояния трубопроводной системы, трубопровод осмотрен в горячем (рабочем) состоянии: " & _
    "провисания, прогибы, уводы линий отсутствуют, дефектов опорно-подвесной системы, защемлений трубопровода не обнаружено." & Strings.Chr(13) & _
    "В холодном (нерабочем) состоянии осматривались:" & Strings.Chr(13) & _
    "- провисания, прогибы, уводы линий отсутствуют, дефектов опорно-подвесной системы, защемлений трубопровода не обнаружено;" & Strings.Chr(13) & _
    "- состояние изоляции: признаки намокания изоляции и подтопления отсутствуют;" & Strings.Chr(13) & _
    "- опорно-подвесная система: в результате ревизии опорно-подвесной системы недопустимых дефектов не обнаружено;" & Strings.Chr(13) & _
    "- арматура: недопустимых дефектов не обнаружено;" & Strings.Chr(13) & _
    "Визуальному и измерительному контролю были подвергнуты:" & Strings.Chr(13) & _
    "- криволинейные элементы (отводы);" & Strings.Chr(13) & "- арматура;" & Strings.Chr(13) & "- врезки;" & Strings.Chr(13) & _
    "- прямые участки, примыкающие к подвергнутым визуальному контролю элементам;" & Strings.Chr(13) & "-   стыковые и угловые сварные соединения труб с подлежащими визуальному контролю элементами." & Strings.Chr(13)
    
End Sub

Sub NGUCGUUDH()
'    ActiveDocument.Variables("ovalnrd").Value = ActiveDocument.Variables("ovalnrd").Value & "; п. 5.4.3.2" & UF1.RD26_260.Value
End Sub

Sub BakKislota()
'    Call DeleteBookmarks("zikl") ' циклы
'    Call DeleteBookmarks("TehnichUstr") ' удаляем пункты про техустройства
'    ActiveDocument.Bookmarks("VikRezTruboprov").Range.Delete ' ВИК результаты осмотра трубопровода
    ActiveDocument.Variables("punkt7-3").Value = Strings.ChrW(31)
    ActiveDocument.Variables("KIPiA").Value = "Оснащение " & ActiveDocument.Variables("TechUsrtva").Value & " соответствует"
    ActiveDocument.Variables("VnutrIzbP").Value = "гидростатическим"
    ActiveDocument.Variables("punkt7-5-4Mat").Value = Strings.ChrW(31)
    ActiveDocument.Variables.item("PassatT").Value = Strings.ChrW(31)
    ActiveDocument.Variables("ProbnDavlen").Value = "наливом"
    ActiveDocument.Variables("SvedOMerPFNP").Value = "26, 27"
End Sub

Sub TehnTruboprovod()
    Call DeleteBookmarks("Rezervuar") ' удаляем часть про резервуары
'    ActiveDocument.Bookmarks("KotlObor").Range.Delete ' пункт 7.3. про оборудование котла
'    ActiveDocument.Bookmarks("R7p4").Range.Delete 'Пункт про КИПиА
'    ActiveDocument.Bookmarks("VikRezTruboprov").Range.Delete ' ВИК результаты осмотра трубопровода в холодном и горячем состоянии
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
    ActiveDocument.Variables.item("PassatT").Value = Strings.ChrW(31)
    ActiveDocument.Variables("p2pr8RD3417302").Value = Strings.ChrW(31)
    ActiveDocument.Variables("ProbnDavlen").Value = "на прочность и плотность,"
    If UF1.CBtt164.Value = True Then ActiveDocument.Variables("ProbnDavlen").Value = ActiveDocument.Variables("ProbnDavlen").Value & " и пневматические испытания на герметичность,"
    ActiveDocument.Variables("TTPoverRash").Value = " поверочного расчета для"
    ActiveDocument.Variables("korpusa").Value = Strings.ChrW(31)
    ActiveDocument.Variables("CBp466-1").Value = " и п. 177" & UF1.FNPTehnTrub.Value
    If ActiveDocument.Bookmarks.Exists("Ne_peremesh") = True Then ActiveDocument.Bookmarks("Ne_peremesh").Range.Delete
    ActiveDocument.Variables("DataIzg6").Value = "Дата монтажа"
   
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
    ActiveDocument.Variables.item("PassatT").Value = Strings.ChrW(31)
    ActiveDocument.Variables("p2pr8RD3417302").Value = Strings.ChrW(31)
    ActiveDocument.Variables("ProbnDavlen").Value = "на прочность и плотность,"
    ActiveDocument.Variables("korpusa").Value = Strings.ChrW(31)
    ActiveDocument.Variables("CBp466-1").Value = " и п. 177" & UF1.FNPTehnTrub.Value
End Sub
Sub RezervuarMazut()
    Call VarSoorugenie
'    ActiveDocument.Variables("SposSoedEl").Value = "Сведения о технологии и сварочных материалах:"
    ActiveDocument.Variables("MnNum7").Value = "Позиция по технологической схеме"
    ActiveDocument.Tables(2).Cell(Row:=2, Column:=1).Range = "Высота налива, мм"
    ActiveDocument.Tables(2).Cell(Row:=2, Column:=2).Range = UF1.RabocheePRub.Value
    ActiveDocument.Tables(3).Delete
    ActiveDocument.Variables("VKorp").Value = ActiveDocument.Variables("VKorp").Value & ", высота " & ActiveDocument.Variables("TechUsrtva").Value & " - " & UF1.RaschetnPRub.Value & " мм"
    ActiveDocument.Variables("RabTempP6").Value = ActiveDocument.Variables("RabTempP6").Value & ", высота налива - " & Trim(UF1.RabocheePRub.Value) & " мм"
    If UF1.CBFNPPBSNN.Value = True Then ActiveDocument.Variables("PrSrSl-PBSNN").Value = Strings.Chr(13) & "Продление срока службы резервуара осуществляется в соответствии с требованиями п. 261" & UF1.FNPPESNN.Value
    ActiveDocument.Variables("DataIzg6").Value = "Дата монтажа"
    ActiveDocument.Variables("SvedOMerPFNP").Value = "26, 27"
    ActiveDocument.Variables("VIKRezKontr").Value = "Корпус резервуара." & Strings.Chr(13) _
    & "Визуальный осмотр конструкций корпуса производился с внутренней и наружной (в доступных местах) сторон. Корпус резервуара по высоте сварен из " & ActiveDocument.Variables("RaschetntRub").Value & " поясов, высотой 1500 мм каждый. " & _
    "На корпусе резервуара расположены 2 люка-лаза. При визуальном осмотре корпуса недопустимых дефектов не обнаружено. Расположение штуцеров, патрубков и люков-лазов соответствует требованиям НТД. " & _
    "Сварные швы окраек днища смещены относительно вертикальных сварных швов корпуса на длину более 200 мм. Состояние корпуса резервуара соответствует требованиям п.п. 8.7, 8.8, 8.13 РД 08-95-95." & _
    Strings.Chr(13) & "Днище резервуара." & Strings.Chr(13) & _
    "При визуальном осмотре днища недопустимых дефектов не обнаружено. При геометрическом контроле днища недопустимых дефектов не обнаружено. " & _
    "Состояние днища резервуара соответствует требованиям п.п. 8.7, 8.8, 8.15 РД 08-95-95." & _
    Strings.Chr(13) & "Кровля и несущие поверхности." & Strings.Chr(13) & _
    "Осмотр кровли и несущих конструкций проводился с внутренней и наружной (в доступных местах) сторон резервуара. При визуальном осмотре кровли недопустимых дефектов не обнаружено. " & _
    "При осмотре сварных швов недопустимых дефектов не обнаружено. Особое внимание было уделено контролю сварных швов приварки люков, штуцеров для установки оборудования. " & _
    "Состояние кровли резервуара соответствует требованиям п.п. 8.7, 8.8, 8.11 РД 08-95-95." & _
    Strings.Chr(13) & "Отмостка резервуара." & Strings.Chr(13) & _
    "- погружение нижней части резервуара в грунт и скопление воды по контуру резервуара отсутствует;" & Strings.Chr(13) & _
    "- поверхность отмостки чистая от растительности;" & Strings.Chr(13) & _
    "- отмостка резервуара имеет необходимый уклон (1:10)." & Strings.Chr(13) & _
    "Состояние отмостки соответствует требованиям п. 5.6 РД 08-95-95." & Strings.Chr(13)
    If UF1.CBRD089595.Value = True And Val(UF1.NaNLet.Value) > 4 Then
        ActiveDocument.Variables("p8Osvid").Value = "Частичное техническое обследование " & ActiveDocument.Variables("TechUsrtva").Value _
        & ", в соответствии с требованиями п. 3.7.1" & UF1.RD089595.Value & ", необходимо провести в срок до " & Format(DateAdd("yyyy", 4, UF1.AktGID.Value), "dd.mm.yyyy") & "."
    Else
        ActiveDocument.Variables("p8Osvid").Value = Strings.ChrW(31)
    End If
    If UF1.ComboBoxRaschet.ListIndex = 0 Then ActiveDocument.Variables.item("OzOsR").Value = "ГОСТ 34233.1-2017, СТО-СА-03-002-2009"
End Sub

Sub BallGroUst()
    Call MnogCHislo
'    ActiveDocument.Variables("TechUsrtva").Value = "баллонов"
'    ActiveDocument.Variables("TechUsrtvo").Value = "баллоны групповой установки"
'    ActiveDocument.Variables("TechUsrtvoB").Value = UCase(Left(ActiveDocument.Variables("TechUsrtvo").Value, 1)) & LCase(Mid(ActiveDocument.Variables("TechUsrtvo").Value, 2)) 'Делаем первую букву большой
    ActiveDocument.Variables("GOST34347PMat").Value = UF1.GOST9731.Value
    ActiveDocument.Variables.item("PassatT").Value = Strings.ChrW(31)
End Sub
Sub Ballon()
'    ActiveDocument.Variables("TechUsrtva").Value = "баллонов"
'    ActiveDocument.Variables("TechUsrtvo").Value = "баллон"
'    ActiveDocument.Variables("TechUsrtvoB").Value = UCase(Left(ActiveDocument.Variables("TechUsrtvo").Value, 1)) & LCase(Mid(ActiveDocument.Variables("TechUsrtvo").Value, 2)) 'Делаем первую букву большой
    ActiveDocument.Variables("GOST34347PMat").Value = UF1.GOST9731.Value
    ActiveDocument.Variables.item("PassatT").Value = Strings.ChrW(31)
End Sub

Sub Nasos()
    Dim EdIzP, StansAssoc As String
    EdIzP = " м.вод.ст."
    If UF1.ComboBoxTipUstroistva.Value = "Компрессор" Then EdIzP = " кгс/см" & Strings.ChrW(178) & "."
    ActiveDocument.Tables(2).Cell(Row:=2, Column:=1).Range = "Напор," & EdIzP
    ActiveDocument.Tables(2).Cell(Row:=2, Column:=2).Range = UF1.RaschetnP.Value
    ActiveDocument.Tables(2).Cell(Row:=3, Column:=1).Range = "Производительность, м" & Strings.ChrW(179) & "/ч."
    ActiveDocument.Variables("RaschSreda").Value = Replace(ActiveDocument.Variables("RaschSreda").Value, ",", ".")
    ActiveDocument.Variables("PasportPar").Value = "Основные технические характеристики"
    ActiveDocument.Variables("RabSredaToplRasch").Value = Strings.Chr(13) & "Рабочая среда - " & UF1.RaschSreda.Value & ". " & Strings.Chr(13) & "Напор: "
'    ActiveDocument.Variables("RaschSreda").Value = ActiveDocument.Variables("RaschSreda").Value & Strings.Chr(13) & "Напор:"
    ActiveDocument.Variables("RaschetnP").Value = UF1.RaschetnP.Value & EdIzP & Strings.Chr(13) & "Производительность: "
    ActiveDocument.Variables("Raschetnt").Value = UF1.VKorp.Value & " м" & Strings.ChrW(179) & "/ч."
    If UF1.Raschetnt.Value <> "" Then ActiveDocument.Variables("Raschetnt").Value = ActiveDocument.Variables("Raschetnt").Value & Strings.Chr(13) & "Температура рабочей среды: " & Trim(UF1.Raschetnt.Value) & Strings.ChrW(176) & "С."
    ActiveDocument.Variables("VKorp").Value = Strings.Chr(13) & "Мощность электродвигателя: " & UF1.RaschetnPRub.Value & " кВт." & Strings.Chr(13) & "Частота вращения: " & UF1.RaschetntRub.Value & " об/мин"
    If ActiveDocument.Bookmarks.Exists("ORPD10") = True Then ActiveDocument.Bookmarks("ORPD10").Range.Delete
'    ActiveDocument.Variables("NaznTehUstr").Value = ActiveDocument.Variables("NaznTehUstr").Value & ActiveDocument.Variables("RabSreda").Value
    ActiveDocument.Variables("punkt7-3-1").Value = "Установка, размещение и оснащение "
    ActiveDocument.Variables("korpusa").Value = " элементов насосной установки"
    ActiveDocument.Variables("ISvarSoed").Value = Strings.ChrW(31)
    ActiveDocument.Variables("p191PBETT").Value = " таблицы 1 " & UF1.M2_96.Value
    ActiveDocument.Variables("RezNKPP").Value = "Техническое состояние " & ActiveDocument.Variables("TechUsrtva").Value & " оценивается как работоспособное при малой вероятности отказа, допустима длительная работа агрегата."
    If UF1.CBFNPHOPO.Value = True Then ActiveDocument.Variables("RezNKPP").Value = ActiveDocument.Variables("RezNKPP").Value & " Состояние насоса соответствует требованиям п.п. 15, 132" & UF1.FNPHOPO.Value & "."
    ActiveDocument.Variables("NTDAktVIK").Value = ActiveDocument.Variables("NTDAktVIK").Value & Strings.Chr(13) & UF1.M2_96.Value
    ActiveDocument.Bookmarks("OORAnPr").Range.Delete
    If UF1.CBGOST20816_1.Value = True Then
        ActiveDocument.Variables.item("OzOsR").Value = "п. 6.3.2.3. и таблицы С.1 Приложения С " & UF1.GOST20816_1.Value
        ActiveDocument.Variables("OORNasosOzenka").Value = "При вибродиагностировании насоса получены оценки " & Strings.Chr(171) & "зона А" & Strings.Chr(187) & " и " & Strings.Chr(171) & "зона Б" & Strings.Chr(187) & " (пригоден для дальнейшей эксплуатации), в соответствии с требованиями " & ActiveDocument.Variables.item("OzOsR").Value
        ActiveDocument.Variables("GOST32106").Value = UF1.GOST20816_1.Value
        StansAssoc = ""
    Else
        ActiveDocument.Variables("OORNasosOzenka").Value = "При вибродиагностировании насоса получены оценки " & Strings.Chr(171) & "Хорошо" & Strings.Chr(187) & " и " & Strings.Chr(171) & "Допустимо" & Strings.Chr(187) & " (допустима дальнейшая длительная эксплуатация), в соответствии с требованиями п. 6.1 и таблицы А.1 приложения А " & UF1.GOST32106.Value
        ActiveDocument.Variables("GOST32106").Value = UF1.GOST32106.Value
        StansAssoc = "; требованиями п. 4.4.2 и таблицы Б.1 приложения Б " & UF1.SA03_001_05.Value
    End If
    ActiveDocument.Variables("OORNasos").Value = Strings.Chr(13) & "Проведен анализ режимов работы и исследование напряженно-деформированного состояния (учтено: отсутствие видимых деформаций корпуса; геометрические параметры элементов насоса)."
    ActiveDocument.Variables("OORNasos").Value = ActiveDocument.Variables("OORNasos").Value & Strings.Chr(13) & ActiveDocument.Variables("OORNasosOzenka").Value
    ActiveDocument.Variables("OORNasos").Value = ActiveDocument.Variables("OORNasos").Value & StansAssoc
    ActiveDocument.Variables("OORNasos").Value = ActiveDocument.Variables("OORNasos").Value & "." & Strings.Chr(13) & "В соответствии с требованиями " & UF1.M2_96.Value & ", регламентирующей виды, периодичность и содержание технического обслуживания, период эксплуатации насоса между капитальными ремонтами, составляет не более 30000 часов. С учетом общего технического состояния насоса, возможно продление срока службы насоса, при действующем режиме эксплуатации, на "
    ActiveDocument.Variables("OORNasos").Value = ActiveDocument.Variables("OORNasos").Value & ActiveDocument.Variables("NaNLet").Value & "."
    ActiveDocument.Variables("VIKRezKontr").Value = "при визуальном и измерительном контроле насоса, механических и коррозионных повреждений, усталостных трещин и других видимых дефектов, препятствующих дальнейшей эксплуатации, не обнаружено."
End Sub

Sub NTDAktVIK()                             'Заполняем НТД в актах НК
    tmp = Strings.ChrW(31)
    tmp1 = "Федеральные нормы и правила в области промышленной безопасности "
    If UF1.CBFNPORPD.Value = True Then      'ФНП ОРПД
        ActiveDocument.Variables("NTDAktVIK").Value = tmp1 & Mid(UF1.FNPORPDR.Value, 64, 104) & "."
        tmp = Strings.Chr(13)
    End If
    If UF1.CBFNPOPVB.Value = True Then      'ФНП ОПВБ
        ActiveDocument.Variables("NTDAktVIK").Value = ActiveDocument.Variables("NTDAktVIK").Value & tmp & tmp1 & Mid(UF1.FNPOPVBR.Value, 64, 122) & "."
        tmp = Strings.Chr(13)
    End If
    If UF1.CBFNPHOPO.Value = True Then      'ФНП ХОПО
        ActiveDocument.Variables("NTDAktVIK").Value = ActiveDocument.Variables("NTDAktVIK").Value & tmp & tmp1 & Mid(UF1.FNPHOPO.Value, 64, 66) & "."
        tmp = Strings.Chr(13)
    End If
    If UF1.CBFNPPBETT.Value = True Then     'ФНП ПБЭТТ
        ActiveDocument.Variables("NTDAktVIK").Value = ActiveDocument.Variables("NTDAktVIK").Value & tmp & tmp1 & Mid(UF1.FNPTehnTrub.Value, 64, 63) & "."
        tmp = Strings.Chr(13)
    End If
    If UF1.CBFNPPBSNN.Value = True Then     'ФНП СНН
        ActiveDocument.Variables("NTDAktVIK").Value = ActiveDocument.Variables("NTDAktVIK").Value & tmp & tmp1 & Mid(UF1.FNPPESNN.Value, 64, 66) & "."
        tmp = Strings.Chr(13)
    End If
    If UF1.CBSO439.Value = True Then        'СО 439 сосуды
        ActiveDocument.Variables("NTDAktVIK").Value = ActiveDocument.Variables("NTDAktVIK").Value & tmp & Mid(UF1.SO439.Value, 2, 94) & "."
        ActiveDocument.Variables("NTDAktNKPD").Value = Strings.Chr(13) & Mid(UF1.SO439.Value, 2, 94) & "."
        ActiveDocument.Variables("NTDAktNK").Value = Mid(UF1.SO439.Value, 2, 94) & "."
        tmp = Strings.Chr(13)
    End If
    If UF1.CBSO464.Value = True Then        'СО 464 трубопроводы
        ActiveDocument.Variables("NTDAktVIK").Value = ActiveDocument.Variables("NTDAktVIK").Value & tmp & Mid(UF1.SO464.Value, 2, 97) & "."
        ActiveDocument.Variables("NTDAktNKPD").Value = Strings.Chr(13) & Mid(UF1.SO464.Value, 2, 97) & "."
        ActiveDocument.Variables("NTDAktNK").Value = Mid(UF1.SO464.Value, 2, 97) & "."
        tmp = Strings.Chr(13)
    End If
    If UF1.CBSO469.Value = True Then        'СО 469 котлы
        ActiveDocument.Variables("NTDAktVIK").Value = ActiveDocument.Variables("NTDAktVIK").Value & tmp & Mid(UF1.SO469.Value, 2, 188) & "."
        ActiveDocument.Variables("NTDAktNKPD").Value = Strings.Chr(13) & Mid(UF1.SO469.Value, 2, 188) & "."
        ActiveDocument.Variables("NTDAktNK").Value = Mid(UF1.SO469.Value, 2, 188) & "."
        tmp = Strings.Chr(13)
    End If
    If UF1.CBGOST34347.Value = True Then        'ГОСТ сосуды и аппараты стальные
        ActiveDocument.Variables("NTDAktVIK").Value = ActiveDocument.Variables("NTDAktVIK").Value & tmp & Mid(UF1.GOST34347.Value, 2, 79) & "."
        tmp = Strings.Chr(13)
    End If
    If UF1.CBRD2626012.Value = True Then        'РД емкости СО2
        ActiveDocument.Variables("NTDAktVIK").Value = ActiveDocument.Variables("NTDAktVIK").Value & tmp & Mid(UF1.GOST34347.Value, 2, 79) & "."
        tmp = Strings.Chr(13)
    End If
    If UF1.CBVM030104.Value = True Then        'РД Временная методика газификаторы
        ActiveDocument.Variables("NTDAktVIK").Value = ActiveDocument.Variables("NTDAktVIK").Value & tmp & Mid(UF1.RDVM03.Value, 2, 229) & "."
        tmp = Strings.Chr(13)
    End If
    If UF1.CBRUA93.Value = True Then            'РУА-93
        ActiveDocument.Variables("NTDAktVIK").Value = ActiveDocument.Variables("NTDAktVIK").Value & tmp & Mid(UF1.RUA93.Value, 2, 140) & "."
        tmp = Strings.Chr(13)
    End If
    If UF1.CBRD1533413752696.Value = True Then  'Баки кислоты и едкого натра
        ActiveDocument.Variables("NTDAktVIK").Value = ActiveDocument.Variables("NTDAktVIK").Value & tmp & Mid(UF1.RD1533413752696.Value, 2, 104) & "." & Strings.Chr(13) & Mid(UF1.Snip31875.Value, 2, 42) & "."
        If ActiveDocument.Variables("TechUsrtvo").Value = "резервуар" Then ActiveDocument.Variables("NTDAktVIK").Value = ActiveDocument.Variables("NTDAktVIK").Value & Strings.Chr(13) & Mid(UF1.Snip31875.Value, 2, 42) & "."
        tmp = Strings.Chr(13)
    End If
    If UF1.CBRD089595.Value = True Then         'РД 08-95-95 Резервуары
        ActiveDocument.Variables("NTDAktVIK").Value = ActiveDocument.Variables("NTDAktVIK").Value & tmp & Mid(UF1.RD089595.Value, 2, 138) & "."
        tmp = Strings.Chr(13)
    End If
End Sub

Sub OformlenBase()
    UF1.Label8.Caption = "Дата изготовления"
    UF1.Label5.Caption = "Завод изг."
    UF1.Label513.Visible = False
    UF1.MontagOrg.Visible = False
    UF1.Label514.Visible = False
    UF1.DataMontaga.Visible = False
    UF1.CBPodNaliv.Visible = True
    UF1.CBVakuum.Visible = True
    UF1.Label475.Caption = "P="
    UF1.VKorp.Visible = True
    UF1.CBRubashka.Visible = True
    UF1.CBRubashka.Caption = "Рубашка"
    UF1.Label480.Visible = False
    UF1.Label480.Caption = "В рубашке Р="
    UF1.RaschetnPRub.Visible = False
    UF1.Label479.Visible = False
    UF1.Label479.Caption = "t="
    UF1.RaschetntRub.Visible = False
    UF1.RaschetntRub.ControlTipText = "Расчетная температура в рубашке"
    UF1.Label481.Visible = False
    UF1.VRub.Visible = False
    UF1.Label486.Visible = False
    UF1.RaschSredaRub.Visible = False
    UF1.Label487.Visible = False
    UF1.Label487.Caption = "В рубашке Р="
    UF1.RabocheePRub.Visible = False
    UF1.Label482.Visible = False
    UF1.Label482.Caption = "t="
    UF1.RabTempRub.Visible = False
    UF1.RabTempRub.ControlTipText = "Рабочая температура в рубашке"
    UF1.Label484.Visible = False
    UF1.RabSredaRub.Visible = False
    UF1.Label491.Visible = False
    UF1.IspitatPRub.Visible = False
    
    UF1.ProtokolVD.Visible = False
    UF1.ProtokolVDD.Visible = False
    UF1.Label465.Visible = False
    UF1.Label466.Visible = False
    UF1.Label511.Caption = "Раб. ср. класс опас."
    
    UF1.Label501.Caption = "Обечайка"
    UF1.Label502.Caption = "Днище"
    UF1.Label432.Caption = "Диаметр"
    UF1.CBZikl.Visible = True
    UF1.CBZikl.Value = True
    UF1.Label18.Visible = True
    UF1.Label462.Caption = "Коэф.осл."
        
    If UF1.OptionSosud.Value = False Then
        UF1.KolZicl.Visible = False
        UF1.CBZikl.Visible = False
        UF1.CBZikl.Value = False
        UF1.Label18.Visible = False
    End If
    
    If UF1.OptionKotel.Value = True Then
        UF1.CBVakuum.Visible = False
        UF1.CBPodNaliv.Visible = False
        UF1.VikMK.Value = True
        UF1.KontrGibCh.Value = True
        UF1.Label418.Caption = "топл."
        UF1.Label485.Caption = "топл."
    End If

    If UF1.OptionTruboprovod.Value = True Then
        UF1.VKorp.Visible = False
        UF1.CBRubashka.Caption = "РОУ"
        UF1.Label5.Caption = "Монт. орг."
        UF1.Label501.Caption = "Труба"
        UF1.Label502.Caption = "Труба"
    End If
    
    If UF1.OptionSoorugenie.Value = True Then
        UF1.Label513.Visible = True
        UF1.MontagOrg.Visible = True
        UF1.Label514.Visible = True
        UF1.DataMontaga.Visible = True
    End If
    
    If UF1.OptionOstalnoe.Value = True Then
        UF1.CBPodNaliv.Visible = False
        UF1.CBRubashka.Visible = False
        UF1.Label475.Caption = "напор"
        UF1.Label479.Caption = "об."
        UF1.Label480.Caption = "Двиг. мощ."
        UF1.RaschetnPRub.Visible = True
        UF1.RaschetntRub.Visible = True
    End If
End Sub
