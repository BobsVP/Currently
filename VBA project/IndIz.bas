Attribute VB_Name = "IndIz"
'�������������� ��������� ��� ������� �������������

Sub ParKotl()
'    ActiveDocument.Variables("punkt7-3NTD").Value = " �." & ActiveDocument.Variables("CBp10").Value & UF1.FNPORPDR.Value
    ActiveDocument.Variables("ElemVodKot").Value = " (���������, �����������, ���� ������������ �������, �������������� ���� � �������� ����� � �.�.)"
    ActiveDocument.Variables("VIKrdSO469").Value = "; �.�. 5.4, 5.5, 5.15, 5.16, 5.17, 5.18" & UF1.SO469.Value
End Sub

Sub VodgKotl()
'    ActiveDocument.Variables("punkt7-3NTD").Value = " �." & ActiveDocument.Variables("CBp10").Value & UF1.FNPORPDR.Value
    ActiveDocument.Variables("ElemVodKot").Value = " (�����������, ���� ������������ �������, �������������� ���� � �������� ����� � �.�.)"
End Sub

Sub ElektroKotel()
    ActiveDocument.Variables.Item("TimeGI").Value = "10 �����"
'    ActiveDocument.Variables("punkt7-3NTD").Value = " �.�. 10, 22" & UF1.FNPORPDR.Value
    ActiveDocument.Variables("VIKrdSO469").Value = "; �.�. 5.4, 5.5" & UF1.SO469.Value
    ActiveDocument.Variables.Item("PassatT").Value = ", ��� ���������� ����������� ������� �� ��������� ������������ ��������� " & Strings.Chr(171) & "������" & Strings.Chr(187) & ", ������������� ��� " & Strings.Chr(171) & "��� �����������" & Strings.Chr(187)
End Sub

Sub Ekonomayzer()
If (ActiveDocument.Variables("ZavodIzg").Value Like "*���*") Or (ActiveDocument.Variables("ZavodIzg").Value Like "*����*") Then
    If (ActiveDocument.Variables("RabSreda").Value Like "*[��]��*") Or (ActiveDocument.Variables("RabSreda").Value = Strings.ChrW(31)) = 0 Then
        ActiveDocument.Variables("VIKrdSO469").Value = "; �. 3.2.1 ���������� 9" & UF1.SO469.Value
    Else
        ActiveDocument.Variables("VIKrdSO469").Value = "; �. 3.1.1 ���������� 9" & UF1.SO469.Value
    End If
End If
'    ActiveDocument.Variables("punkt7-3NTD").Value = " �." & ActiveDocument.Variables("CBp10").Value & UF1.FNPORPDR.Value
    ActiveDocument.Variables("p12-1pril2").Value = " �.�. 2, 3, 4, 5 ���������� �8" & UF1.FNPORPDR.Value
    ActiveDocument.Variables("tverdSO469").Value = " ���������� 1 ���� 1412-85 " & Strings.Chr(171) & "����� � ������������ �������� ��� �������. �����" & Strings.Chr(187)
    ActiveDocument.Variables("ISvarSoed").Value = " �������"
    ActiveDocument.Variables("svarnih").Value = Strings.ChrW(31)
    ActiveDocument.Variables("korpusa").Value = " ��������� ������������"
    ActiveDocument.Variables("ObechBarKotl").Value = "��������� ����� ������������"
    ActiveDocument.Variables("Punkt3211RD1024998").Value = "3.3.1.1."
End Sub

Sub Vozduhosbornik()
End Sub

Sub Avtozisterna()
    ActiveDocument.Variables("KIPiA").Value = "��������� " & ActiveDocument.Variables("TechUsrtva").Value & " ������������ ������������ � ���������� ����� �������������"
End Sub

Sub Podogrevatel()
'    ActiveDocument.Variables("TechUsrtva").Value = "�������������"
'    ActiveDocument.Variables("TechUsrtvo").Value = "�������������"
'    ActiveDocument.Variables("TechUsrtvoB").Value = UCase(Left(ActiveDocument.Variables("TechUsrtvo").Value, 1)) & LCase(Mid(ActiveDocument.Variables("TechUsrtvo").Value, 2)) '������ ������ ����� �������
End Sub

Sub Gasifikator()
    ActiveDocument.Bookmarks("P7p1mat").Range.Delete ' ����� 7.1. ���������
    ActiveDocument.Variables.Item("VnuOsmotr").Value = "�������� ������," & Strings.Chr(13) & "�������������� ���������," & Strings.Chr(13) & "����������� ����������������."
    ActiveDocument.Variables.Item("VIKrd").Value = Strings.ChrW(31)
    ActiveDocument.Variables("obmerz").Value = "���������� �������� �������� ������ � ��������� �� ����������."
    ActiveDocument.Variables("KorrozPovr").Value = ", ����������� �� ����� ������������"
    ActiveDocument.Variables("KorrozPovr1").Value = Strings.ChrW(31)
    ActiveDocument.Variables("KorrozPovr2").Value = ", ��������� ����������� �� ����� ������������"
    ActiveDocument.Variables("DopDoc").Value = Strings.Chr(13) & "- ��� �� ���������� ������ ������� � ��������� � ����������������� ������� ����������� ������������ �� ���������� ���������������, �������� �� " & Strings.Chr(171) & "��������� ��. ���� �.�." & Strings.Chr(187) & " �� " & UF1.AktGID.Value & " �. - 1 �.;"
    ActiveDocument.Variables("DopDoc").Value = ActiveDocument.Variables("DopDoc").Value & Strings.Chr(13) & "- ��� �� ���������� ������ ������� � ��������� ����� ��������������� ����������� ������������ ���-3/1,6-200�, ���.�" & UF1.ZavN.Value & ", �������� �� " & Strings.Chr(171) & "��������� ��. ���� �.�." & Strings.Chr(187) & " �� " & UF1.AktGID.Value & " �. - 1 �.;"
    ActiveDocument.Variables("DopDoc").Value = ActiveDocument.Variables("DopDoc").Value & Strings.Chr(13) & "- ��� �� ������������� ����������� ������������ ���-3/1,6-200�, �������� �� " & Strings.Chr(171) & "��������� ��. ���� �.�." & Strings.Chr(187) & " �� " & UF1.AktGID.Value & " �. - 1 �.;"
    ActiveDocument.Variables("DopDoc").Value = ActiveDocument.Variables("DopDoc").Value & Strings.Chr(13) & "- ��� �� ��������� ����������������� �������� �� ���-3/1,6-200� ���. �" & UF1.ZavN.Value & ", �������� �� " & Strings.Chr(171) & "��������� ��. ���� �.�." & Strings.Chr(187) & " �� " & UF1.AktGID.Value & " �. - 1 �.;"
End Sub

Sub VakuumSosud()
'    ActiveDocument.Bookmarks("R7p4").Range.Delete
    ActiveDocument.Variables("IndxP").Value = "���"
    ActiveDocument.Variables("GOST34347PGiI").Value = Strings.ChrW(31)
    ActiveDocument.Variables("RazreshaemoeVKM").Value = "������ �� " & Format((1 - CDbl(UF1.RazreshaemoeP.Value)) / 10, "###0.0#####") & "(" & Format(1 - CDbl(UF1.RazreshaemoeP.Value), "###0.0#####") & ")"
'    ActiveDocument.Variables("RazreshaemoeP").Value = "���������� " & ActiveDocument.Variables("RazreshaemoeP").Value
    ActiveDocument.Variables("DavlNeVishe").Value = " ��� ����������� �������� (������ �� " & Format(1 - CDbl(UF1.RazreshaemoeP.Value), "###0.0#####") & " ���/��" & Strings.ChrW(178) & ")"
End Sub

Sub SosudPodNaliv()
'    ActiveDocument.Bookmarks("R7p4").Range.Delete
'    Call DeleteBookmarks("zikl") ' �����
    ActiveDocument.Variables("GOST34347PGiI").Value = Strings.ChrW(31)
    ActiveDocument.Variables("IspitatP").Value = "������ �����"
    ActiveDocument.Variables("TimeGI").Value = "4 �����"
    ActiveDocument.Variables("PadDavl").Value = Strings.ChrW(31)
    ActiveDocument.Variables("PodRabDav").Value = Strings.ChrW(31)
    ActiveDocument.Variables("DavlNeVishe").Value = " ��� ����������� �������� (��� �����)"
    ActiveDocument.Variables("p2pr8RD3417302").Value = Strings.ChrW(31)
End Sub

Sub SosudHOPO()
'    ActiveDocument.Variables("TechUsrtva").Value = "����"
'    ActiveDocument.Variables("TechUsrtvo").Value = "���"
'    Call DeleteBookmarks("zikl") ' �����
'    ActiveDocument.Variables("punkt7-3-1").Value = "���������� "
    ActiveDocument.Variables("punkt7-3").Value = Strings.ChrW(31)
    ActiveDocument.Variables("KIPiA").Value = "��������� " & ActiveDocument.Variables("TechUsrtva").Value & " �������������"
    ActiveDocument.Variables("VnutrIzbP").Value = "����������������"
    ActiveDocument.Variables("VnutrP").Value = "�����������������"
    ActiveDocument.Variables("punkt7-5-4Mat").Value = Strings.ChrW(31)
    ActiveDocument.Variables.Item("PassatT").Value = Strings.ChrW(31)
'    ActiveDocument.Variables.Item("OzOsR").Value = "�. 11.10" & UF1.RD1533413752696.Value
    ActiveDocument.Variables("ProbnDavlen").Value = "�������"
    
End Sub

Sub TruboprovPara()
    ActiveDocument.Variables("Izgotovitel6").Value = "��������� �����������"
    ActiveDocument.Variables("SostElTrub").Value = "������� ������ � ������-��������� ������� �����������, ����������� ��������, � � �������� ������, ������ ���������, �������� �����, �� ����������. "
    If ActiveDocument.Bookmarks.Exists("Ne_peremesh") = True Then ActiveDocument.Bookmarks("Ne_peremesh").Range.Delete
    ActiveDocument.Variables("TempDeform").Value = ", �������������� ������������"
    ActiveDocument.Variables("punkt7-5-4Mat").Value = Strings.ChrW(31)
End Sub

Sub NGUCGUUDH()
'    ActiveDocument.Variables("ovalnrd").Value = ActiveDocument.Variables("ovalnrd").Value & "; �. 5.4.3.2" & UF1.RD26_260.Value
End Sub

Sub BakKislota()
'    Call DeleteBookmarks("zikl") ' �����
'    Call DeleteBookmarks("TehnichUstr") ' ������� ������ ��� �������������
    ActiveDocument.Bookmarks("VikRezTruboprov").Range.Delete ' ��� ���������� ������� ������������
    ActiveDocument.Variables("punkt7-3").Value = Strings.ChrW(31)
    ActiveDocument.Variables("KIPiA").Value = "��������� " & ActiveDocument.Variables("TechUsrtva").Value & " �������������"
    ActiveDocument.Variables("VnutrIzbP").Value = "����������������"
    ActiveDocument.Variables("punkt7-5-4Mat").Value = Strings.ChrW(31)
    ActiveDocument.Variables.Item("PassatT").Value = Strings.ChrW(31)
    ActiveDocument.Variables("ProbnDavlen").Value = "�������"
    ActiveDocument.Variables("SvedOMerPFNP").Value = "26, 27"
End Sub

Sub TehnTruboprovod()
    Call DeleteBookmarks("Rezervuar") ' ������� ����� ��� ����������
'    ActiveDocument.Bookmarks("KotlObor").Range.Delete ' ����� 7.3. ��� ������������ �����
'    ActiveDocument.Bookmarks("R7p4").Range.Delete '����� ��� �����
    ActiveDocument.Bookmarks("VikRezTruboprov").Range.Delete ' ��� ���������� ������� ������������ � �������� � ������� ���������
    ActiveDocument.Tables(1).Cell(2, 2).Range.Text = "�����"
    ActiveDocument.Tables(1).Cell(3, 2).Range.Text = "�����"
    If ActiveDocument.Variables("UZKRekTT").Value = Strings.ChrW(31) Then ActiveDocument.Variables("UZKRekTT").Value = ". ��������, �� ����������� �������� ������� ���������� ���������������� ������������ � ������������ � ���� � 55724-2013 " & _
    Strings.Chr(171) & "�������� �������������. ���������� �������. ������ ��������������" & Strings.Chr(187) & ", �� �������������."
'    ActiveDocument.Variables("MnNum3").Value = "�����������"
'    ActiveDocument.Variables("MnNum4").Value = "����������"
'    ActiveDocument.Variables("MnNum8").Value = "����������"
    Call VarSoorugenie
'    ActiveDocument.Variables("TechDiagn").Value = "������������"
'    ActiveDocument.Variables("TechDiagnB").Value = "������������"
'    ActiveDocument.Variables("tehdiagnnk").Value = "������������"
'    ActiveDocument.Variables("tehdiagn").Value = "������������"
'    ActiveDocument.Variables("tehdiagn1").Value = "������������"
'    ActiveDocument.Variables("Izgotovitel6").Value = "��������� �����������"
    ActiveDocument.Variables("TechTr0").Value = ", ����������� ������� �� ���������, ������ ���������� ������� ����������� � ����������� � ���������� ������������"
    ActiveDocument.Variables("punkt7-3-1").Value = "����������, ���������, ���������� � ��������� "
    ActiveDocument.Variables("punkt7-3").Value = Strings.ChrW(31)
    ActiveDocument.Variables("SvedOMer0").Value = "26, 27" 'Strings.ChrW(31)
    ActiveDocument.Variables("SvedOMerPFNP").Value = "24, 25, 26, 27"
'    ActiveDocument.Variables("NaznTehUstr").Value = ActiveDocument.Variables("NaznTehUstr").Value & ActiveDocument.Variables("RabSreda").Value
    ActiveDocument.Variables("SostElTrub").Value = "���������������� ��������� ������������ ������������� ���������� �����. ����������� ������� ��������� ������������ ������������� ���������� �����. ���������� ������� �������� ��������� ������������. "
    ActiveDocument.Variables("TTrDop").Value = " ������� ������� ���������������, �������������� � ����������� ����������� �� ����������� ���������� ��� ��������������. ������� ���������� ������������� ���������������� ����� � ��������� ���������� ������������."
    ActiveDocument.Variables("TTrDop3").Value = Strings.Chr(13) & "���������� ����������� ��������� ���������� ������������. �� ��������� ������� ���������������-����������� ������������, ����������� ���������-�������������� � ������ ������� �������������� �������� ������ �����: ����������� ��������� ���������� ������������� ��������� ����������."
    ActiveDocument.Variables("TTrDop2").Value = " ������ ���������� ������� ����������� � ����������� ������������ � ���������� ������������,"
    ActiveDocument.Variables("TTrDop4").Value = ", ����������� ������������ ����������� ��������� ������������ � ����������� ����������� ����������, ��������� �������� � ����������� ��������� � �����, ����������� ����������������� ��������� �����������, �� ����������� ������� � ��������� ����������, ����������� ������� �������� ������������� ���������"
    ActiveDocument.Variables("TTrDop5").Value = " ����������� ������� ������� ���������������, �������������� � ����������� �����������. �������� ���������� ������������� ���������������� ����� � ��������� ���������� " & ActiveDocument.Variables("TechUsrtva").Value & "."
    ActiveDocument.Variables("TTrDop5Som").Value = ";" & Strings.Chr(13) & vbTab & "- ����������� ������� ������� ���������������, �������������� � ����������� �����������;" & Strings.Chr(13) & vbTab & "- �������� ���������� ������������� ���������������� ����� � ��������� ���������� " & ActiveDocument.Variables("TechUsrtva").Value
    ActiveDocument.Variables("TTrDop6").Value = " ����������� ����������� ��������� ���������� � ��������� � ���������� �����������."
    ActiveDocument.Variables("TTrDop6SoM").Value = ";" & Strings.Chr(13) & vbTab & "- ����������� ����������� ��������� ���������� � ��������� � ���������� �����������"
    ActiveDocument.Variables("TTrDop7").Value = " ���������� ������ ����������� � ������ ���������� ��� ������������ ����������, �������� � �����������, ����������� ��������."
    ActiveDocument.Variables("TTrDop7SoM").Value = ";" & Strings.Chr(13) & vbTab & "- ���������� ������ ����������� � ������ ���������� ��� ������������ ����������, �������� � �����������, ����������� ��������"
    ActiveDocument.Variables("TTrDop8").Value = " ������ ���������� ������� ����������� � ����������� ������������ � ���������� ������������."
    ActiveDocument.Variables("TTrDop8SoM").Value = ";" & Strings.Chr(13) & vbTab & "- ������ ���������� ������� ����������� � ����������� ������������ � ���������� ������������"
    ActiveDocument.Variables("P71ProekDok").Value = " �� ��������������� ������, ����������� ����������� � ������������ � ��������� �������������."
    ActiveDocument.Variables.Item("PassatT").Value = Strings.ChrW(31)
    ActiveDocument.Variables("p2pr8RD3417302").Value = Strings.ChrW(31)
    ActiveDocument.Variables("ProbnDavlen").Value = "�� ��������� � ���������,"
    If UF1.CBtt100.Value = True Then ActiveDocument.Variables("ProbnDavlen").Value = ActiveDocument.Variables("ProbnDavlen").Value & " � �������������� ��������� �� �������������,"
    ActiveDocument.Variables("TTPoverRash").Value = " ����������� ������� ���"
    ActiveDocument.Variables("korpusa").Value = Strings.ChrW(31)
    ActiveDocument.Variables("CBp466-1").Value = " � �. 177" & UF1.FNPTehnTrub.Value
    If ActiveDocument.Bookmarks.Exists("Ne_peremesh") = True Then ActiveDocument.Bookmarks("Ne_peremesh").Range.Delete
   
End Sub
Sub TruboprovodKislota()
    Call DeleteBookmarks("Rezervuar") ' ������� ����� ��� ����������
'    ActiveDocument.Bookmarks("KotlObor").Range.Delete ' ����� 7.3. ��� ������������ �����
'    ActiveDocument.Variables("MnNum1").Value = "����������"
'    ActiveDocument.Variables("MnNum2").Value = "�����������"
'    ActiveDocument.Variables("MnNum3").Value = "�����������"
    Call VarSoorugenie
    ActiveDocument.Variables("TechTr0").Value = ", ����������� ������� � ������ ���������� ������� ����������� � ����������� � ���������� ������������"
    ActiveDocument.Variables("punkt7-3-1").Value = "����������, ���������, ���������� � ��������� "
    ActiveDocument.Variables("punkt7-3").Value = Strings.ChrW(31)
    ActiveDocument.Variables("SvedOMer0").Value = "25, 27" 'Strings.ChrW(31)
    ActiveDocument.Variables("SvedOMerPFNP").Value = "24, 25, 26, 27"
'    ActiveDocument.Variables("NaznTehUstr").Value = ActiveDocument.Variables("NaznTehUstr").Value & ActiveDocument.Variables("RabSreda").Value
    ActiveDocument.Variables("SostElTrub").Value = "���������������� ��������� � ����������� ������� ��������� ������������ ������������� ���������� �����. ���������� ������� �������� ��������� ������������. "
    ActiveDocument.Variables("TTrDop").Value = " ������� ������� ���������������, �������������� � ����������� ����������� �� ����������� ���������� ��� ��������������. ������� ���������� ������������� ���������������� ����� � ��������� ���������� ������������."
    ActiveDocument.Variables("TTrDop1").Value = " ����������� ��������� ������������ ������������� ����������� �.�. 30, 169" & UF1.FNPOPVBR.Value & "."
    ActiveDocument.Variables("TTrDop3").Value = Strings.Chr(13) & "���������� ����������� ��������� ���������� ������������. �� ��������� ������� ���������������-����������� ������������, ����������� ���������-�������������� � ������ ������� �������������� �������� ������ �����: ����������� ��������� ���������� ������������� ��������� ����������."
    ActiveDocument.Variables("TTrDop2").Value = " ������ ���������� ������� ����������� � ����������� ������������ � ���������� ������������,"
    ActiveDocument.Variables.Item("PassatT").Value = Strings.ChrW(31)
    ActiveDocument.Variables("p2pr8RD3417302").Value = Strings.ChrW(31)
    ActiveDocument.Variables("ProbnDavlen").Value = "�� ��������� � ���������,"
    ActiveDocument.Variables("korpusa").Value = Strings.ChrW(31)
    ActiveDocument.Variables("CBp466-1").Value = " � �. 177" & UF1.FNPTehnTrub.Value
End Sub
Sub RezervuarMazut()
'    ActiveDocument.Variables("MnNum2").Value = "�����������"
'    ActiveDocument.Variables("MnNum3").Value = "�����������"
'    ActiveDocument.Variables("MnNum1").Value = "����������"
'    ActiveDocument.Variables("MnNum4").Value = "����������"
'    ActiveDocument.Variables("MnNum8").Value = "����������"
    Call VarSoorugenie
'    ActiveDocument.Variables("TechDiagn").Value = "������������"
'    ActiveDocument.Variables("TechDiagnB").Value = "������������"
'    ActiveDocument.Variables("tehdiagnnk").Value = "������������"
'    ActiveDocument.Variables("tehdiagn").Value = "������������"
'    ActiveDocument.Variables("tehdiagn1").Value = "������������"
'    ActiveDocument.Variables("Izgotovitel6").Value = "��������� �����������"
    ActiveDocument.Variables("PrSrSl-PBSNN").Value = Strings.Chr(13) & "��������� ����� ������ ���������� �������������� � ������������ � ������������ �. 261" & UF1.FNPPESNN.Value
End Sub

Sub BallGroUst()
    Call MnogCHislo
'    ActiveDocument.Variables("TechUsrtva").Value = "��������"
'    ActiveDocument.Variables("TechUsrtvo").Value = "������� ��������� ���������"
'    ActiveDocument.Variables("TechUsrtvoB").Value = UCase(Left(ActiveDocument.Variables("TechUsrtvo").Value, 1)) & LCase(Mid(ActiveDocument.Variables("TechUsrtvo").Value, 2)) '������ ������ ����� �������
    ActiveDocument.Variables("GOST34347PMat").Value = UF1.GOST9731.Value
    ActiveDocument.Variables.Item("PassatT").Value = Strings.ChrW(31)
End Sub
Sub Ballon()
'    ActiveDocument.Variables("TechUsrtva").Value = "��������"
'    ActiveDocument.Variables("TechUsrtvo").Value = "������"
'    ActiveDocument.Variables("TechUsrtvoB").Value = UCase(Left(ActiveDocument.Variables("TechUsrtvo").Value, 1)) & LCase(Mid(ActiveDocument.Variables("TechUsrtvo").Value, 2)) '������ ������ ����� �������
    ActiveDocument.Variables("GOST34347PMat").Value = UF1.GOST9731.Value
    ActiveDocument.Variables.Item("PassatT").Value = Strings.ChrW(31)
End Sub

Sub Nasos()
    ActiveDocument.Tables(2).Cell(Row:=2, Column:=1).Range = "�����, �.���.��."
    ActiveDocument.Tables(2).Cell(Row:=2, Column:=2).Range = UF1.RaschetnP.Value
    ActiveDocument.Tables(2).Cell(Row:=3, Column:=1).Range = "������������������, �" & Strings.ChrW(179) & "/�."
    ActiveDocument.Variables("RaschSreda").Value = Replace(ActiveDocument.Variables("RaschSreda").Value, ",", ".")
    ActiveDocument.Variables("PasportPar").Value = "�������� ����������� ��������������"
    ActiveDocument.Variables("RabSredaToplRasch").Value = Strings.Chr(13) & ActiveDocument.Variables("RabSredaToplRasch").Value
    ActiveDocument.Variables("RaschSreda").Value = ActiveDocument.Variables("RaschSreda").Value & Strings.Chr(13) & "�����:"
    ActiveDocument.Variables("RaschetnP").Value = UF1.RaschetnP.Value & " �.���.��." & Strings.Chr(13) & "������������������: "
    ActiveDocument.Variables("Raschetnt").Value = Trim(UF1.Raschetnt.Value) & " �" & Strings.ChrW(179) & "/�."
    ActiveDocument.Variables("VKorp").Value = Strings.Chr(13) & "�������� ����������������: " & UF1.VKorp.Value & " ���"
    If ActiveDocument.Bookmarks.Exists("ORPD10") = True Then ActiveDocument.Bookmarks("ORPD10").Range.Delete
'    ActiveDocument.Variables("NaznTehUstr").Value = ActiveDocument.Variables("NaznTehUstr").Value & ActiveDocument.Variables("RabSreda").Value
'    ActiveDocument.Variables("punkt7-3").Value = Strings.ChrW(31)
    ActiveDocument.Variables("korpusa").Value = " ��������� �������� ���������"
    ActiveDocument.Variables("ISvarSoed").Value = Strings.ChrW(31)
    ActiveDocument.Variables("p191PBETT").Value = " ������� 1 " & UF1.M2_96.Value
    If UF1.CBFNPHOPO.Value = True Then ActiveDocument.Variables("TTrDop3").Value = " ��������� ������ ������������� ����������� �.�. 15, 132" & UF1.FNPHOPO.Value & "."
    ActiveDocument.Variables("RezNKPP").Value = "����������� ��������� " & ActiveDocument.Variables("TechUsrtva").Value & " ����������� ��� ��������������� ��� ����� ����������� ������, ��������� ���������� ������ ��������."
    ActiveDocument.Variables("M2_96").Value = Strings.Chr(13) & UF1.M2_96.Value
    ActiveDocument.Variables("GOST32106").Value = UF1.GOST32106.Value
    ActiveDocument.Variables("OORNasos").Value = Strings.Chr(13) & "�������� ������ ������� ������ � ������������ ����������-���������������� ��������� (������: ���������� ������� ���������� �������; �������������� ��������� ��������� ������)."
    ActiveDocument.Bookmarks("OORAnPr").Range.Delete
    ActiveDocument.Variables("OORNasos").Value = ActiveDocument.Variables("OORNasos").Value & Strings.Chr(13) & "��� ��������������������� ������ �������� ������ " & Strings.Chr(171) & "������" & Strings.Chr(187) & " � " & Strings.Chr(171) & "���������" & Strings.Chr(187) & " (��������� ���������� ���������� ������������), � ������������ � ������������ �. 6.1 � ������� �.1 ���������� � " & UF1.GOST32106.Value
    ActiveDocument.Variables("OORNasos").Value = ActiveDocument.Variables("OORNasos").Value & "; ������������ �. 4.4.2 � ������� �.1 ���������� � " & UF1.SA03_001_05.Value
    ActiveDocument.Variables("OORNasos").Value = ActiveDocument.Variables("OORNasos").Value & "." & Strings.Chr(13) & "� ������������ � ������������ " & UF1.M2_96.Value & ", ���������������� ����, ������������� � ���������� ������������ ������������, ������ ������������ ������ ����� ������������ ���������, ���������� �� ����� 30000 �����. � ������ ������ ������������ ��������� ������, �������� ��������� ����� ������ ������, ��� ����������� ������ ������������, �� "
    ActiveDocument.Variables("OORNasos").Value = ActiveDocument.Variables("OORNasos").Value & ActiveDocument.Variables("NaNLet").Value & "."
    ActiveDocument.Variables("VIKRezKontr").Value = "��� ���������� � ������������� �������� ������, ������������ � ������������ �����������, ����������� ������ � ������ ������� ��������, �������������� ���������� ������������, �� ����������"
End Sub


