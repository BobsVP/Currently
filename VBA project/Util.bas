Attribute VB_Name = "Util"
'В этом модуле мелкие корректировки для каждого устройства
Option Explicit

Sub MnogCHislo()
ActiveDocument.Variables("MnNum1").Value = "технические устройства"
ActiveDocument.Variables("MnNum2").Value = "применяемые"
ActiveDocument.Variables("MnNum3").Value = "применяемые"
ActiveDocument.Variables("MnNum4").Value = "технических устройств"
ActiveDocument.Variables("MnNum5").Value = "отработавших"
ActiveDocument.Variables("MnNum6").Value = "Заводские номера"
ActiveDocument.Variables("MnNum7").Value = "Регистрационные номера"
ActiveDocument.Variables("MnNum8").Value = "техническим устройствам"
ActiveDocument.Variables("No").Value = UF1.poleRegNum.Value & "№№"
ActiveDocument.Variables("ZavNo").Value = "№№"

End Sub

Sub VarSoorugenie() 'Устаноавливаем переменные для сооружений
    ActiveDocument.Variables("MnNum1").Value = "сооружение"
    ActiveDocument.Variables("MnNum2").Value = "применяемое"
    ActiveDocument.Variables("MnNum3").Value = "применяемое"
    ActiveDocument.Variables("MnNum4").Value = "сооружения"
    ActiveDocument.Variables("MnNum8").Value = "сооружению"
    ActiveDocument.Variables("MnNum9").Value = "Сооружение"
    ActiveDocument.Variables("TechDiagn").Value = "обследование"
    ActiveDocument.Variables("TechDiagnB").Value = "Обследование"
    ActiveDocument.Variables("tehdiagnnk").Value = "обследования"
    ActiveDocument.Variables("tehdiagn").Value = "обследования"
    ActiveDocument.Variables("tehdiagn1").Value = "обследования"
    ActiveDocument.Variables("Izgotovitel6").Value = "Монтажная организация"
End Sub

Sub LegVosG() ' Устанавливаем пункты для ЛВЖ

End Sub

Sub GorGid() ' Устанавливаем пункты для горючих жидкостей
'    Call SetComboBox(SelectComboBox, "CBtt")
'    SelectComboBox = Array(137, 141, 142, 144, 146, 147, 148, 149, 150)
End Sub

Sub RaschOstRes() ' Расчет остаточного ресурса

Dim H As Double
Dim R As Double
    'Скорость коррозии
    If IsNumeric(UF1.otolsh.Value) And IsNumeric(UF1.otolshfakt.Value) And Val(ActiveDocument.Variables("SrokSlugb").Value) <> 0 Then
        ActiveDocument.Variables("oSkorKorroz").Value = Format((UF1.otolsh.Value - UF1.otolshfakt.Value) / Val(ActiveDocument.Variables("SrokSlugb").Value), "0.0##")
    Else
        ActiveDocument.Variables("oSkorKorroz").Value = Strings.ChrW(31)
    End If
    If IsNumeric(UF1.dtolsh.Value) And IsNumeric(UF1.dtolshfakt.Value) And Val(ActiveDocument.Variables("SrokSlugb").Value) <> 0 Then
        ActiveDocument.Variables("dSkorKorroz").Value = Format((UF1.dtolsh.Value - UF1.dtolshfakt.Value) / Val(ActiveDocument.Variables("SrokSlugb").Value), "0.0##")
    Else
        ActiveDocument.Variables("dSkorKorroz").Value = Strings.ChrW(31)
    End If
    If ActiveDocument.Variables("oSkorKorroz").Value <> Strings.ChrW(31) And ActiveDocument.Variables("dSkorKorroz").Value <> Strings.ChrW(31) Then
        ActiveDocument.Variables("MaxSK").Value = Max(CDbl(ActiveDocument.Variables("oSkorKorroz").Value), CDbl(ActiveDocument.Variables("dSkorKorroz").Value))
    End If
    
    'Первая строка расчет
    If IsNumeric(UF1.odiam.Value) And IsNumeric(UF1.DopuskNapro.Value) And IsNumeric(UF1.Koof_fio.Value) And IsNumeric(UF1.PribNaKorro.Value) Then
        Select Case UF1.ComboBoxRaschet.ListIndex
        
            Case Is = 0 'Пассат
                If IsNumeric(UF1.RazreshaemoeP.Value) Then
                    ActiveDocument.Variables("otolshrasch").Value = Format(CDbl(UF1.RazreshaemoeP.Value) * CDbl(UF1.odiam.Value) / (2 * CDbl(UF1.DopuskNapro.Value) * 10 * CDbl(UF1.Koof_fio.Value) - _
                    (CDbl(UF1.RazreshaemoeP.Value))) + CDbl(UF1.PribNaKorro.Value), "0.0#")
                End If
            Case Is = 1 'РД 10-249-98
                If IsNumeric(UF1.RazreshaemoeP.Value) Then
                    If UF1.OptionTruboprovod.Value = True Then
                        ActiveDocument.Variables("Sotbrako").Value = OtbrTolTabl(CDbl(UF1.odiam.Value))
                        ActiveDocument.Variables("otolshrasch").Value = RashTolshTr(1)
                    Else
                        ActiveDocument.Variables("otolshrasch").Value = Format(CDbl(UF1.RazreshaemoeP.Value) / 10 * CDbl(UF1.odiam.Value) / (2 * CDbl(UF1.DopuskNapro.Value) * CDbl(UF1.Koof_fio.Value) - _
                        CDbl(UF1.RazreshaemoeP.Value) / 10) + CDbl(UF1.PribNaKorro.Value), "0.0#")
                        ActiveDocument.Variables("otolshraschK").Value = RashTolshTr(1)
                        ActiveDocument.Variables("ProvUslo").Value = Format((CDbl(ActiveDocument.Variables("otolshrasch").Value) - CDbl(UF1.PribNaKorro.Value)) / CDbl(UF1.odiam.Value), "0.0###")
                        ActiveDocument.Variables("ProvUslK").Value = Format((CDbl(ActiveDocument.Variables("otolshraschK").Value) - CDbl(UF1.PribNaKorro.Value)) / CDbl(UF1.odiam.Value), "0.0###")
                    End If
                End If
            Case Is = 2 'РД 153-34.1-37.525-96 кислота и щелочь
            
            Case Is = 3 'ГОСТ 32388-2013 тех. трубопровод
                If IsNumeric(UF1.RazreshaemoeP.Value) Then
                    ActiveDocument.Variables("Sotbrako").Value = OtbrTolTablTT(CDbl(UF1.odiam.Value))
                    ActiveDocument.Variables("otolshrasch").Value = RashTolshTr(1)
                End If
            Case Is = 4 'ГОСТ 25215-82 баллоны
            
            Case Is = 5 'Справочник Лащинский
                If IsNumeric(UF1.odlina.Value) And IsNumeric(UF1.RabocheePRub.Value) And IsNumeric(UF1.RabTempRub.Value) Then
                    Dim l As Double, b As Double, ratio As Double, K As Double, P As Double
                    If Val(UF1.odlina.Value) > Val(UF1.odiam.Value) Then
                        l = CDbl(UF1.odlina.Value)
                        b = CDbl(UF1.odiam.Value)
                    Else
                        l = CDbl(UF1.odiam.Value)
                        b = CDbl(UF1.odlina.Value)
                    End If
                    ActiveDocument.Variables("BakRaschl").Value = l / 1000
                    ActiveDocument.Variables("BakRaschb").Value = b / 1000
                    ratio = l / b
                    ActiveDocument.Variables("ratio").Value = Format(ratio, "0.0#")
                    K = Format(SpravKoeffK(ratio), "0.0##")
                    ActiveDocument.Variables("KoefTT").Value = K
                    H = CDbl(UF1.RabocheePRub.Value)
                    If (UF1.RabocheePRub.Value / 100) > 1 Then H = (UF1.RabocheePRub.Value / 1000)
                    P = 9.81 * H * CDbl(UF1.RabTempRub.Value) / 1000000
                    ActiveDocument.Variables("BakRaschPgidrst").Value = Format(P, "0.0##")
                    ActiveDocument.Variables("otolshrasch").Value = Format(K * b * Sqr((P / UF1.DopuskNapro.Value)), "0.0#")
                End If
        
            Case Is = 7 'Резервуары
                If IsNumeric(UF1.odlina.Value) And IsNumeric(UF1.odiam.Value) And IsNumeric(UF1.RabTempRub.Value) And IsNumeric(UF1.RabocheePRub.Value) And IsNumeric(UF1.DopuskNapro.Value) And IsNumeric(UF1.PribNaKorro.Value) Then
                    H = CDbl(UF1.RabocheePRub.Value)
                    If (H / 1000) > 1 Then H = (H / 1000)
                    R = UF1.odiam.Value / 2
                    If (R / 100) > 1 Then R = (R / 1000)
                    ActiveDocument.Variables("otolshrasch").Value = Format(((CDbl(UF1.RabTempRub.Value) * (H - 0.3) + 1.2 * 0.0002) * R) / (100 * 0.8 * UF1.DopuskNapro.Value) + UF1.PribNaKorro.Value + UF1.Koof_fio.Value, "0.0#")
                    If IsNumeric(UF1.otolshfakt.Value) And Val(ActiveDocument.Variables("SrokSlugb").Value) <> 0 Then ActiveDocument.Variables("Tosto").Value = _
                    Format((CDbl(UF1.otolshfakt.Value) - CDbl(ActiveDocument.Variables("otolshrasch").Value)) / CDbl(ActiveDocument.Variables("oSkorKorroz").Value), "0.0")
                    ActiveDocument.Variables("LZNII").Value = Format(-Log(0.95) / Val(ActiveDocument.Variables("SrokSlugb").Value), "0.000")
                End If

        End Select
        
    Else
        ActiveDocument.Variables("otolshrasch").Value = Strings.ChrW(31)
    End If
    
    'Вторая строка расчет
    If IsNumeric(UF1.ddiam.Value) And IsNumeric(UF1.DopuskNaprd.Value) And IsNumeric(UF1.Koof_fid.Value) And IsNumeric(UF1.PribNaKorrd.Value) Then
        Select Case UF1.ComboBoxRaschet.ListIndex
        
            Case Is = 0 'Пассат
                If IsNumeric(UF1.ddlina.Value) And IsNumeric(UF1.RazreshaemoeP.Value) Then
                    R = (CDbl(UF1.ddiam.Value) * CDbl(UF1.ddiam.Value)) / (4 * CDbl(UF1.ddlina.Value))
                    ActiveDocument.Variables("dtolshrasch").Value = Format(CDbl(UF1.RazreshaemoeP.Value) * R / (2 * CDbl(UF1.DopuskNaprd.Value) * 10 * CDbl(UF1.Koof_fid.Value) - _
                    0.5 * (CDbl(UF1.RazreshaemoeP.Value))) + CDbl(UF1.PribNaKorrd.Value), "0.0#")
                End If
                
            Case Is = 1 'РД 10-249-98
                If IsNumeric(UF1.RazreshaemoeP.Value) Then
                    If UF1.OptionTruboprovod.Value = True Then
                        ActiveDocument.Variables("Sotbrakd").Value = OtbrTolTabl(CDbl(UF1.ddiam.Value))
                        ActiveDocument.Variables("dtolshrasch").Value = RashTolshTr(2)
                    Else
                        If IsNumeric(UF1.ddlina.Value) Then
                            R = CDbl(UF1.ddiam.Value) / (2 * CDbl(UF1.ddlina.Value))
                            ActiveDocument.Variables("dtolshrasch").Value = Format((CDbl(UF1.RazreshaemoeP.Value) / 10 * CDbl(UF1.ddiam.Value) / (4 * CDbl(UF1.DopuskNaprd.Value) * CDbl(UF1.Koof_fid.Value) - _
                            CDbl(UF1.RazreshaemoeP.Value) / 10)) * R + CDbl(UF1.PribNaKorrd.Value), "0.0#")
                        End If
                        If IsNumeric(UF1.ddlina.Value) Then ActiveDocument.Variables("ProvUsld1").Value = Format(CDbl(UF1.ddlina.Value) / CDbl(UF1.ddiam.Value), "0.0")
                        ActiveDocument.Variables("ProvUsld2").Value = Format((CDbl(ActiveDocument.Variables("dtolshrasch").Value) - CDbl(UF1.PribNaKorrd.Value)) / CDbl(UF1.ddiam.Value), "0.0###")
                    End If
                End If
            Case Is = 2 'РД 153-34.1-37.525-96 кислота и щелочь
            
            Case Is = 3 'ГОСТ 32388-2013 тех. трубопровод
                If IsNumeric(UF1.RazreshaemoeP.Value) Then
                    ActiveDocument.Variables("Sotbrakd").Value = OtbrTolTablTT(CDbl(UF1.ddiam.Value))
                    ActiveDocument.Variables("dtolshrasch").Value = RashTolshTr(2)
                End If
            Case Is = 4 'ГОСТ 25215-82 баллоны
            
            Case Is = 5 'Справочник Лащинский
                If IsNumeric(UF1.ddlina.Value) And IsNumeric(UF1.RabocheePRub.Value) And IsNumeric(UF1.RabTempRub.Value) Then
                    If Val(UF1.ddlina.Value) > Val(UF1.ddiam.Value) Then
                        l = CDbl(UF1.ddlina.Value)
                        b = CDbl(UF1.ddiam.Value)
                    Else
                        l = CDbl(UF1.ddiam.Value)
                        b = CDbl(UF1.ddlina.Value)
                    End If
                    ActiveDocument.Variables("BakRaschls").Value = l / 1000
                    ActiveDocument.Variables("BakRaschbs").Value = b / 1000
                    ratio = l / b
                    ActiveDocument.Variables("ratios").Value = Format(ratio, "0.0#")
                    K = Format(SpravKoeffK(ratio), "0.0##")
                    ActiveDocument.Variables("KoefZakr").Value = K
                    H = CDbl(UF1.RabocheePRub.Value)
                    If (UF1.RabocheePRub.Value / 100) > 1 Then H = (UF1.RabocheePRub.Value / 1000)
                    P = 9.81 * H * CDbl(UF1.RabTempRub.Value) / 1000000
                    ActiveDocument.Variables("BakRaschPgidrst").Value = Format(P, "0.0##")
                    ActiveDocument.Variables("dtolshrasch").Value = Format(K * b * Sqr((P / UF1.DopuskNaprd.Value)), "0.0#")
                End If
        
            Case Is = 7 'Резервуары
                If IsNumeric(UF1.odlina.Value) And IsNumeric(UF1.ddiam.Value) And IsNumeric(UF1.RabTempRub.Value) And IsNumeric(UF1.RabocheePRub.Value) And IsNumeric(UF1.DopuskNaprd.Value) And IsNumeric(UF1.PribNaKorrd.Value) Then
                    H = CDbl(UF1.RabocheePRub.Value - UF1.odlina.Value)
                    If (H / 1000) > 1 Then H = (H / 1000)
                    R = UF1.ddiam.Value / 2
                    If (R / 100) > 1 Then R = (R / 1000)
                    ActiveDocument.Variables("dtolshrasch").Value = Format(((CDbl(UF1.RabTempRub.Value) * H + 1.2 * 0.0002) * R) / (100 * 0.8 * UF1.DopuskNaprd.Value) + UF1.PribNaKorrd.Value + UF1.Koof_fid.Value, "0.0#")
'                    If IsNumeric(UF1.otolshfakt.Value) And Val(ActiveDocument.Variables("SrokSlugb").Value) <> 0 Then ActiveDocument.Variables("Tosto").Value = _
'                    Format((CDbl(UF1.otolshfakt.Value) - CDbl(ActiveDocument.Variables("otolshrasch").Value)) / CDbl(ActiveDocument.Variables("oSkorKorroz").Value), "0.0")
                End If
            
        End Select
        
    Else
        ActiveDocument.Variables("dtolshrasch").Value = Strings.ChrW(31)
    End If
    
    'Заполняем табличку-подсказку
    UF1.Inform.Caption = "Срок службы: " & ActiveDocument.Variables("SrokSlugb").Value & Strings.Chr(13) & "Скорость коррозии: " & ActiveDocument.Variables("oSkorKorroz").Value & _
    " мм/год, расчетная толщина - " & ActiveDocument.Variables("otolshrasch").Value & Strings.Chr(13) & "Скорость коррозии: " & ActiveDocument.Variables("dSkorKorroz").Value & " мм/год, расчетная толщина - " _
    & ActiveDocument.Variables("dtolshrasch").Value & Strings.Chr(13)
End Sub

'Public Function A1_Get()
'    A1_Get = Array(2, 3, 10, 22, 38, 39, 43, 45, 46, 47, 49, 50, 61, 64, 65, 68, 69, 71, 80, 81, 85, 86, 90, 91, 100, 154, 156, 175, 177, 178, 179, 184, 185, 186, 187, 188, 190, 246, 257, _
'    258, 260, 267, 268, 269, 270, 271, 338, 339, 340, 341, 343, 348, 353, 372, 373, 374, 378, 379, 394, 465, 466, 468, 469, 471, 500, 502, 503, 505, 506, 519, 521, 523, 538, 539, _
'    540, 577, 589)
'End Function

Public Function RashTolshTr(VarRasch As Integer) ' Расчетная толщина для труб и коллекторов
If VarRasch = 1 Then
    RashTolshTr = Format(CDbl(UF1.RazreshaemoeP.Value) / 10 * CDbl(UF1.odiam.Value) / (2 * CDbl(UF1.DopuskNapro.Value) * CDbl(UF1.Koof_fio.Value) + _
    CDbl(UF1.RazreshaemoeP.Value) / 10) + CDbl(UF1.PribNaKorro.Value), "0.0#")
Else
    RashTolshTr = Format(CDbl(UF1.RazreshaemoeP.Value) / 10 * CDbl(UF1.ddiam.Value) / (2 * CDbl(UF1.DopuskNaprd.Value) * CDbl(UF1.Koof_fid.Value) + _
    CDbl(UF1.RazreshaemoeP.Value) / 10) + CDbl(UF1.PribNaKorrd.Value), "0.0#")
End If
End Function

Public Function OtbrTolTabl(VarRasch As Double) ' Возвращаем отбраковочную толщину трубопровода по РД 10-249-98
    If VarRasch > 108 Then OtbrTolTabl = "3,2"
    If VarRasch <= 108 Then OtbrTolTabl = "2,8"
    If VarRasch <= 90 Then OtbrTolTabl = "2,4"
    If VarRasch <= 70 Then OtbrTolTabl = "2,0"
    If VarRasch <= 51 Then OtbrTolTabl = "1,6"
    If VarRasch < 38 Then OtbrTolTabl = "1,45"
End Function

Public Function OtbrTolTablTT(VarRasch As Double) ' Возвращаем отбраковочную толщину трубопровода по ГОСТ 32388-2013 тех. трубопровод
    If VarRasch >= 426 Then OtbrTolTablTT = "4,0"
    If VarRasch <= 377 Then OtbrTolTablTT = "3,5"
    If VarRasch <= 325 Then OtbrTolTablTT = "3,0"
    If VarRasch <= 219 Then OtbrTolTablTT = "2,5"
    If VarRasch <= 114 Then OtbrTolTablTT = "2,0"
    If VarRasch <= 57 Then OtbrTolTablTT = "1,5"
    If VarRasch <= 25 Then OtbrTolTablTT = "1,0"
    
    If Val(ActiveDocument.Variables("SrokSlugb").Value) > 30 Then
        ActiveDocument.Variables("KoefTT").Value = "0,95"
        ActiveDocument.Variables("bolmen").Value = "более"
    Else
        ActiveDocument.Variables("KoefTT").Value = "1,0"
        ActiveDocument.Variables("bolmen").Value = "менее"
    End If

End Function

Public Function SpravKoeffK(ratio As Double)
    Dim thresholds As Variant
    Dim values As Variant
    thresholds = Array(1#, 1.05, 1.1, 1.15, 1.2, 1.25, 1.3, 1.35, 1.4, 1.45, 1.5, 1.55, 1.6, 1.65, 1.7, 1.8, 1.9, 2#)
    values = Array(0.31, 0.32, 0.33, 0.34, 0.36, 0.38, 0.395, 0.41, 0.43, 0.44, 0.45, 0.46, 0.465, 0.47, 0.48, 0.49, 0.495, 0.5)
    Dim i As Integer
    Dim found As Boolean
    found = False
    For i = LBound(thresholds) To UBound(thresholds)
        If ratio = thresholds(i) Then
            SpravKoeffK = values(i)
            found = True
            Exit For
        ElseIf ratio < thresholds(i) Then
            ' Если ratio находится между двумя порогами, выполняем интерполяцию
            If i > LBound(thresholds) Then
                SpravKoeffK = values(i - 1) + (values(i) - values(i - 1)) * (ratio - thresholds(i - 1)) / (thresholds(i) - thresholds(i - 1))
            End If
            found = True
            Exit For
        End If
    Next i
    If Not found Then SpravKoeffK = values(UBound(values))

End Function

Function Max(a, b)
    If a > b Then Max = a Else Max = b
End Function

Sub PnGIsp()
Dim SelectComboBox, mark As Variant
If UF1.PnIs.Value = True Then
    UF1.Controls.item("CBp190").Value = False
    If UF1.OptionKotel.Value = True Then
        SelectComboBox = Array(175, 177, 185, 186, 187, 188)
        Call SetComboBox(SelectComboBox, "CBp")
    End If
    If UF1.OptionSosud.Value = True Then
        SelectComboBox = Array(175, 178, 185, 187, 188)
        Call SetComboBox(SelectComboBox, "CBp")
    End If
    If UF1.OptionTruboprovod.Value = True Then
        SelectComboBox = Array(175, 184, 185, 187, 188)
        Call SetComboBox(SelectComboBox, "CBp")
    End If
End If
If IsNull(UF1.PnIs.Value) Then
    For Each mark In AllCBp
        If mark > 174 And mark < 191 Then UF1.Controls.item("CBp" & mark).Value = False
    Next
    UF1.Controls.item("CBp175").Value = True
    UF1.Controls.item("CBp190").Value = True
End If

End Sub

Public Sub SetComboBox(SelectComboBox, tipFNP)
Dim mark As Variant
    For Each mark In SelectComboBox
        UF1.Controls.item(tipFNP & mark).Value = True
    Next mark
End Sub

Public Sub SetExpert(ExpertORPD, ExpertHim, ExpertGas, ExpertZS, ExpertZSM, ExpertSNN)
    UF1.ExpertORPD.Value = ExpertORPD
    UF1.ExpertHim.Value = ExpertHim
    UF1.ExpertGas.Value = ExpertGas
    UF1.ExpertZS.Value = ExpertZS
    UF1.ExpertZSM.Value = ExpertZSM
    UF1.ExpertSNN.Value = ExpertSNN
End Sub

Public Sub SetFNP(CBFNPORPD, CBFNPOPVB, CBFNPHOPO, CBFNPPBETT, CBFNPPBSNN)
    UF1.CBFNPORPD.Value = CBFNPORPD
    UF1.CBFNPOPVB.Value = CBFNPOPVB
    UF1.CBFNPHOPO.Value = CBFNPHOPO
    UF1.CBFNPPBETT.Value = CBFNPPBETT
    UF1.CBFNPPBSNN.Value = CBFNPPBSNN
End Sub

Public Sub Sroki()
    ActiveDocument.Variables("SrokSlugbZNII").Value = AddYears(Val(ActiveDocument.Variables("SrokSlugb").Value))
    ActiveDocument.Variables("NaNLet").Value = AddYears(UF1.NaNLet.Value)
    ActiveDocument.Variables("DoNgoda").Value = Format(DateAdd("yyyy", Val(UF1.NaNLet.Value), UF1.AktGID.Value), "dd.mm.yyyy")
    ActiveDocument.Variables("DoNgodaOsv").Value = Format(DateAdd("yyyy", 3, UF1.AktGID.Value), "dd.mm.yyyy")

End Sub

Function AddYears(NumberOfYears As Integer)
    Dim tmp As Integer
    
    If NumberOfYears > 20 Then
        tmp = NumberOfYears Mod 10
    Else
        tmp = NumberOfYears
    End If
    
    If tmp = 1 Then
        AddYears = NumberOfYears & " год"
    ElseIf tmp > 1 And tmp < 5 Then
        AddYears = NumberOfYears & " года"
    ElseIf tmp > 4 Then
        AddYears = NumberOfYears & " лет"
    End If
End Function

Function MinusDopusk(Tolchina As Double)
    Dim thresholds As Variant
    Dim values As Variant
    thresholds = Array(0#, 5.5, 7.5, 25#, 30#, 34#, 40#, 50#)
    values = Array(0.5, 0.6, 0.8, 0.9, 1#, 1.1, 1.2, 1.3)
    Dim i As Integer
    For i = LBound(thresholds) To UBound(thresholds)
        If Tolchina > thresholds(i) Then MinusDopusk = values(i)
    Next i

End Function
