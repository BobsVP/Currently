Attribute VB_Name = "Util"
'� ���� ������ ������ ������������� ��� ������� ����������
Option Explicit

Sub MnogCHislo()
ActiveDocument.Variables("MnNum1").Value = "����������� ����������"
ActiveDocument.Variables("MnNum2").Value = "�����������"
ActiveDocument.Variables("MnNum3").Value = "�����������"
ActiveDocument.Variables("MnNum4").Value = "����������� ���������"
ActiveDocument.Variables("MnNum5").Value = "������������"
ActiveDocument.Variables("MnNum6").Value = "��������� ������"
ActiveDocument.Variables("MnNum7").Value = "��������������� ������"
ActiveDocument.Variables("MnNum8").Value = "����������� �����������"
ActiveDocument.Variables("No").Value = UF1.poleRegNum.Value & "��"
ActiveDocument.Variables("ZavNo").Value = "��"

End Sub

Sub VarSoorugenie() '�������������� ���������� ��� ����������
    ActiveDocument.Variables("MnNum1").Value = "����������"
    ActiveDocument.Variables("MnNum2").Value = "�����������"
    ActiveDocument.Variables("MnNum3").Value = "�����������"
    ActiveDocument.Variables("MnNum4").Value = "����������"
    ActiveDocument.Variables("MnNum8").Value = "����������"
    ActiveDocument.Variables("MnNum9").Value = "����������"
    ActiveDocument.Variables("TechDiagn").Value = "������������"
    ActiveDocument.Variables("TechDiagnB").Value = "������������"
    ActiveDocument.Variables("tehdiagnnk").Value = "������������"
    ActiveDocument.Variables("tehdiagn").Value = "������������"
    ActiveDocument.Variables("tehdiagn1").Value = "������������"
    ActiveDocument.Variables("Izgotovitel6").Value = "��������� �����������"
End Sub

Sub LegVosG() ' ������������� ������ ��� ���

End Sub

Sub GorGid() ' ������������� ������ ��� ������� ���������
'    Call SetComboBox(SelectComboBox, "CBtt")
'    SelectComboBox = Array(137, 141, 142, 144, 146, 147, 148, 149, 150)
End Sub

Sub RaschOstRes()
If IsNumeric(UF1.ddiam.Value) And IsNumeric(UF1.otolsh.Value) And IsNumeric(UF1.otolshfakt.Value) And IsNumeric(UF1.DopuskNapro.Value) Then
    Dim R As Double
        If IsNumeric(UF1.otolshfakt.Value) Then ActiveDocument.Variables("oSkorKorroz").Value = Format((UF1.otolsh.Value - UF1.otolshfakt.Value) / Val(ActiveDocument.Variables("SrokSlugb").Value), "#0.0##")
        If IsNumeric(UF1.dtolshfakt.Value) Then ActiveDocument.Variables("dSkorKorroz").Value = Format((UF1.dtolsh.Value - UF1.dtolshfakt.Value) / Val(ActiveDocument.Variables("SrokSlugb").Value), "#0.0##")
    If UF1.ComboBoxRaschet.ListIndex = 0 Then '������
        ActiveDocument.Variables("otolshrasch").Value = Format(CDbl(UF1.RazreshaemoeP.Value) * CDbl(UF1.odiam.Value) / (2 * CDbl(UF1.DopuskNapro.Value) * 10 * CDbl(UF1.Koof_fio.Value) - _
        (CDbl(UF1.RazreshaemoeP.Value))) + CDbl(UF1.PribNaKorro.Value), "##0.0#")
        If IsNumeric(UF1.ddlina.Value) Then R = (CDbl(UF1.ddiam.Value) * CDbl(UF1.ddiam.Value)) / (4 * CDbl(UF1.ddlina.Value))
        ActiveDocument.Variables("dtolshrasch").Value = Format(CDbl(UF1.RazreshaemoeP.Value) * R / (2 * CDbl(UF1.DopuskNaprd.Value) * 10 * CDbl(UF1.Koof_fid.Value) - _
        0.5 * (CDbl(UF1.RazreshaemoeP.Value))) + CDbl(UF1.PribNaKorrd.Value), "##0.0#")
    End If
    If UF1.ComboBoxRaschet.ListIndex = 1 Then  '�� 10-249-98
        If UF1.OptionTruboprovod.Value = True Then
            ActiveDocument.Variables("Sotbrako").Value = OtbrTolTabl(CDbl(UF1.odiam.Value))
            ActiveDocument.Variables("Sotbrakd").Value = OtbrTolTabl(CDbl(UF1.ddiam.Value))
            ActiveDocument.Variables("MaxSK").Value = Max(CDbl(ActiveDocument.Variables("oSkorKorroz").Value), CDbl(ActiveDocument.Variables("dSkorKorroz").Value))
            ActiveDocument.Variables("otolshrasch").Value = RashTolshTr(1)
            ActiveDocument.Variables("dtolshrasch").Value = RashTolshTr(2)
        Else
            ActiveDocument.Variables("otolshrasch").Value = Format(CDbl(UF1.RazreshaemoeP.Value) / 10 * CDbl(UF1.odiam.Value) / (2 * CDbl(UF1.DopuskNapro.Value) * CDbl(UF1.Koof_fio.Value) - _
            CDbl(UF1.RazreshaemoeP.Value) / 10) + CDbl(UF1.PribNaKorro.Value), "##0.0#")
            ActiveDocument.Variables("otolshraschK").Value = RashTolshTr(1)
            ActiveDocument.Variables("ProvUslo").Value = Format((CDbl(ActiveDocument.Variables("otolshrasch").Value) - CDbl(UF1.PribNaKorro.Value)) / CDbl(UF1.odiam.Value), "##0.0###")
            ActiveDocument.Variables("ProvUslK").Value = Format((CDbl(ActiveDocument.Variables("otolshraschK").Value) - CDbl(UF1.PribNaKorro.Value)) / CDbl(UF1.odiam.Value), "##0.0###")
            R = CDbl(UF1.ddiam.Value) / (2 * CDbl(UF1.ddlina.Value))
            ActiveDocument.Variables("dtolshrasch").Value = Format((CDbl(UF1.RazreshaemoeP.Value) / 10 * CDbl(UF1.ddiam.Value) / (4 * CDbl(UF1.DopuskNaprd.Value) * CDbl(UF1.Koof_fid.Value) - _
            CDbl(UF1.RazreshaemoeP.Value) / 10)) * R + CDbl(UF1.PribNaKorrd.Value), "##0.0#")
            ActiveDocument.Variables("ProvUsld1").Value = CDbl(UF1.ddlina.Value) / CDbl(UF1.ddiam.Value)
            ActiveDocument.Variables("ProvUsld2").Value = Format((CDbl(ActiveDocument.Variables("dtolshrasch").Value) - CDbl(UF1.PribNaKorrd.Value)) / CDbl(UF1.ddiam.Value), "##0.0###")
        End If
    End If
    If UF1.ComboBoxRaschet.ListIndex = 2 Then   '�� 153-34.1-37.525-96 ������� � ������
    End If
    If UF1.ComboBoxRaschet.ListIndex = 3 Then  '���� 32388-2013 ���. �����������
        ActiveDocument.Variables("Sotbrako").Value = OtbrTolTablTT(CDbl(UF1.odiam.Value))
        ActiveDocument.Variables("Sotbrakd").Value = OtbrTolTablTT(CDbl(UF1.ddiam.Value))
        ActiveDocument.Variables("MaxSK").Value = Max(CDbl(ActiveDocument.Variables("oSkorKorroz").Value), CDbl(ActiveDocument.Variables("dSkorKorroz").Value))
        ActiveDocument.Variables("otolshrasch").Value = RashTolshTr(1)
        ActiveDocument.Variables("dtolshrasch").Value = RashTolshTr(2)
        If Val(ActiveDocument.Variables("SrokSlugb").Value) > 30 Then
            ActiveDocument.Variables("KoefTT").Value = "0,95"
            ActiveDocument.Variables("bolmen").Value = "�����"
        Else
            ActiveDocument.Variables("KoefTT").Value = "1,0"
            ActiveDocument.Variables("bolmen").Value = "�����"
        End If
    End If
    If UF1.ComboBoxRaschet.ListIndex = 4 Then  '���� 25215-82 �������
    End If
    If UF1.ComboBoxRaschet.ListIndex = 5 Then   '���������� ���������
    End If
'    ActiveDocument.Variables("SrokSlugb").Value = Year(Date) - Val(Right(UF1.DataVvoda.Value, 4))
'    If UF1.otolshfakt.Value <> "" Then
'        ActiveDocument.Variables("oSkorKorroz").Value = Format((UF1.otolsh.Value - UF1.otolshfakt.Value) / Val(ActiveDocument.Variables("SrokSlugb").Value), "#0.0##")
'        ActiveDocument.Variables("otolshrasch").Value = Format(CDbl(UF1.RazreshaemoeP.Value) / 10 * CDbl(UF1.odiam.Value) / (2 * 145 + (CDbl(UF1.RazreshaemoeP.Value) / 10)), "##0.0#")
'    End If
'    If UF1.dtolshfakt.Value <> "" Then
'        ActiveDocument.Variables("dSkorKorroz").Value = Format((UF1.dtolsh.Value - UF1.dtolshfakt.Value) / Val(ActiveDocument.Variables("SrokSlugb").Value), "#0.0##")
'        ActiveDocument.Variables("dtolshrasch").Value = Format(CDbl(UF1.RazreshaemoeP.Value) / 10 * CDbl(UF1.ddiam.Value) / (2 * 145 + (CDbl(UF1.RazreshaemoeP.Value) / 10)), "##0.0#")
'    End If
    If Val(ActiveDocument.Variables("SrokSlugb").Value) < 31 Then ActiveDocument.Variables("KoefTT").Value = "1,0"
    UF1.Inform.Caption = "���� ������: " & ActiveDocument.Variables("SrokSlugb").Value & Strings.Chr(13) & "�������� ��������: " & ActiveDocument.Variables("oSkorKorroz").Value & _
    " ��/���, ��������� ������� - " & ActiveDocument.Variables("otolshrasch").Value & Strings.Chr(13) & "�������� ��������: " & ActiveDocument.Variables("dSkorKorroz").Value & " ��/���, ��������� ������� - " _
    & ActiveDocument.Variables("dtolshrasch").Value & Strings.Chr(13)
End If
End Sub

'Public Function A1_Get()
'    A1_Get = Array(2, 3, 10, 22, 38, 39, 43, 45, 46, 47, 49, 50, 61, 64, 65, 68, 69, 71, 80, 81, 85, 86, 90, 91, 100, 154, 156, 175, 177, 178, 179, 184, 185, 186, 187, 188, 190, 246, 257, _
'    258, 260, 267, 268, 269, 270, 271, 338, 339, 340, 341, 343, 348, 353, 372, 373, 374, 378, 379, 394, 465, 466, 468, 469, 471, 500, 502, 503, 505, 506, 519, 521, 523, 538, 539, _
'    540, 577, 589)
'End Function

Public Function RashTolshTr(VarRasch As Integer) ' ��������� ������� ��� ���� � �����������
If VarRasch = 1 Then
    RashTolshTr = Format(CDbl(UF1.RazreshaemoeP.Value) / 10 * CDbl(UF1.odiam.Value) / (2 * CDbl(UF1.DopuskNapro.Value) * CDbl(UF1.Koof_fio.Value) + _
    CDbl(UF1.RazreshaemoeP.Value) / 10) + CDbl(UF1.PribNaKorro.Value), "##0.0#")
Else
    RashTolshTr = Format(CDbl(UF1.RazreshaemoeP.Value) / 10 * CDbl(UF1.ddiam.Value) / (2 * CDbl(UF1.DopuskNaprd.Value) * CDbl(UF1.Koof_fid.Value) + _
    CDbl(UF1.RazreshaemoeP.Value) / 10) + CDbl(UF1.PribNaKorrd.Value), "##0.0#")
End If
End Function

Public Function OtbrTolTabl(VarRasch As Double) ' ���������� ������������� ������� ������������ �� �� 10-249-98
    If VarRasch > 108 Then OtbrTolTabl = "3,2"
    If VarRasch <= 108 Then OtbrTolTabl = "2,8"
    If VarRasch <= 90 Then OtbrTolTabl = "2,4"
    If VarRasch <= 70 Then OtbrTolTabl = "2,0"
    If VarRasch <= 51 Then OtbrTolTabl = "1,6"
    If VarRasch < 38 Then OtbrTolTabl = "1,45"
End Function

Public Function OtbrTolTablTT(VarRasch As Double) ' ���������� ������������� ������� ������������ �� ���� 32388-2013 ���. �����������
    If VarRasch >= 426 Then OtbrTolTablTT = "4,0"
    If VarRasch <= 377 Then OtbrTolTablTT = "3,5"
    If VarRasch <= 325 Then OtbrTolTablTT = "3,0"
    If VarRasch <= 219 Then OtbrTolTablTT = "2,5"
    If VarRasch <= 114 Then OtbrTolTablTT = "2,0"
    If VarRasch <= 57 Then OtbrTolTablTT = "1,5"
    If VarRasch <= 25 Then OtbrTolTablTT = "1,0"
End Function

Function Max(a, b)
If a > b Then Max = a Else Max = b
End Function

Public Function Indx(ByRef Arr, str)
    Indx = UBound(Arr)
    Dim n As Long
    For n = LBound(Arr) To UBound(Arr)
        If Arr(n, 1) = str Then Indx = n - 1
        If Indx <> UBound(Arr) Then Exit For
    Next n
End Function
