Attribute VB_Name = "IndIz"
'�������������� ��������� ��� ������� �������������

Sub ParKotl()
'    ActiveDocument.Variables("punkt7-3NTD").Value = " �." & ActiveDocument.Variables("CBp10").Value & UF1.FNPORPDR.Value
    ActiveDocument.Variables("korpusa").Value = " ��������� �����"
    ActiveDocument.Variables("CBp466-1").Value = " � �. 6.6" & UF1.SO469.Value
    ActiveDocument.Variables("ElemVodKot").Value = " (���������, �����������, ���� ������������ �������, �������������� ���� � �������� ����� � �.�.)"
    ActiveDocument.Variables("VIKrdSO469").Value = "; �.�. 5.4, 5.5, 5.15, 5.16, 5.17, 5.18" & UF1.SO469.Value
'    ActiveDocument.Tables(2).Rows(3).Delete
End Sub

Sub VodgKotl()
'    ActiveDocument.Variables("punkt7-3NTD").Value = " �." & ActiveDocument.Variables("CBp10").Value & UF1.FNPORPDR.Value
    ActiveDocument.Variables("korpusa").Value = " ��������� �����"
    ActiveDocument.Variables("CBp466-1").Value = " � �. 6.6" & UF1.SO469.Value
    ActiveDocument.Variables("ElemVodKot").Value = " (�����������, ���� ������������ �������, �������������� ���� � �������� ����� � �.�.)"
End Sub

Sub ElektroKotel()
    ActiveDocument.Variables.item("TimeGI").Value = "10 �����"
'    ActiveDocument.Variables("punkt7-3NTD").Value = " �.�. 10, 22" & UF1.FNPORPDR.Value
    ActiveDocument.Variables("VIKrdSO469").Value = "; �.�. 5.4, 5.5" & UF1.SO469.Value
    ActiveDocument.Variables.item("PassatT").Value = ", ��� ���������� ����������� ������� �� ��������� ������������ ��������� " & Strings.Chr(171) & "������" & Strings.Chr(187) & ", ������������� ��� " & Strings.Chr(171) & "��� �����������" & Strings.Chr(187)
    ActiveDocument.Variables("CBp466-1").Value = " � �. 6.6" & UF1.SO469.Value
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
    ActiveDocument.Variables("CBp466-1").Value = " � �. 6.6" & UF1.SO469.Value
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
    ActiveDocument.Variables("punkt7-3-1").Value = "���������, ���������� � ��������� "
    ActiveDocument.Variables.item("VnuOsmotr").Value = "�������� ������," & Strings.Chr(13) & "�������������� ���������," & Strings.Chr(13) & "����������� ����������������."
    ActiveDocument.Variables.item("VIKrd").Value = Strings.ChrW(31)
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
    ActiveDocument.Variables("RazreshaemoeVKM").Value = "������ �� " & Format((1 - CDbl(UF1.RazreshaemoeP.Value)) / 10, "0.0#####") & "(" & Format(1 - CDbl(UF1.RazreshaemoeP.Value), "0.0#####") & ")"
'    ActiveDocument.Variables("RazreshaemoeP").Value = "���������� " & ActiveDocument.Variables("RazreshaemoeP").Value
    ActiveDocument.Variables("DavlNeVishe").Value = " ��� ����������� �������� (������ �� " & Format(1 - CDbl(UF1.RazreshaemoeP.Value), "0.0#####") & " ���/��" & Strings.ChrW(178) & ")"
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
    ActiveDocument.Variables("punkt7-3").Value = Strings.ChrW(31)
    ActiveDocument.Variables("KIPiA").Value = "��������� " & ActiveDocument.Variables("TechUsrtva").Value & " �������������"
    ActiveDocument.Variables("VnutrIzbP").Value = "����������������"
    ActiveDocument.Variables("VnutrP").Value = "�����������������"
    ActiveDocument.Variables("punkt7-5-4Mat").Value = Strings.ChrW(31)
    ActiveDocument.Variables.item("PassatT").Value = Strings.ChrW(31)
    ActiveDocument.Variables("ProbnDavlen").Value = "�������"
    ActiveDocument.Tables(1).Cell(Row:=1, Column:=4).Range = "������, ��"
    ActiveDocument.Variables("RabTempP6").Value = ActiveDocument.Variables("RabTempP6").Value & ", ������ ������ - " & Trim(UF1.RabocheePRub.Value) & " ��"
    ActiveDocument.Tables(2).Cell(Row:=3, Column:=1).Range = "������ ������, ��"
    ActiveDocument.Tables(2).Cell(Row:=3, Column:=2).Range = "�� " & UF1.RabocheePRub.Value
End Sub

Sub TruboprovPara()
    ActiveDocument.Variables("Izgotovitel6").Value = "��������� �����������"
    ActiveDocument.Variables("SostElTrub").Value = "������� ������ � ������-��������� ������� �����������, ����������� ��������, � � �������� ������, ������ ���������, �������� �����, �� ����������. "
    If ActiveDocument.Bookmarks.Exists("Ne_peremesh") = True Then ActiveDocument.Bookmarks("Ne_peremesh").Range.Delete
    ActiveDocument.Variables("TempDeform").Value = ", �������������� ������������"
    ActiveDocument.Variables("punkt7-5-4Mat").Value = Strings.ChrW(31)
    ActiveDocument.Variables("DataIzg6").Value = "���� �������"
    ActiveDocument.Variables("VIKRezKontr").Value = "��� ������ ������ ��������� �������������� �������, ����������� �������� � ������� (�������) ���������: " & _
    "����������, �������, ����� ����� �����������, �������� ������-��������� �������, ���������� ������������ �� ����������." & Strings.Chr(13) & _
    "� �������� (���������) ��������� �������������:" & Strings.Chr(13) & _
    "- ����������, �������, ����� ����� �����������, �������� ������-��������� �������, ���������� ������������ �� ����������;" & Strings.Chr(13) & _
    "- ��������� ��������: �������� ��������� �������� � ����������� �����������;" & Strings.Chr(13) & _
    "- ������-��������� �������: � ���������� ������� ������-��������� ������� ������������ �������� �� ����������;" & Strings.Chr(13) & _
    "- ��������: ������������ �������� �� ����������;" & Strings.Chr(13) & _
    "����������� � �������������� �������� ���� �����������:" & Strings.Chr(13) & _
    "- ������������� �������� (������);" & Strings.Chr(13) & "- ��������;" & Strings.Chr(13) & "- ������;" & Strings.Chr(13) & _
    "- ������ �������, ����������� � ������������ ����������� �������� ���������;" & Strings.Chr(13) & "-   �������� � ������� ������� ���������� ���� � ����������� ����������� �������� ����������." & Strings.Chr(13)
    
End Sub

Sub NGUCGUUDH()
'    ActiveDocument.Variables("ovalnrd").Value = ActiveDocument.Variables("ovalnrd").Value & "; �. 5.4.3.2" & UF1.RD26_260.Value
End Sub

Sub BakKislota()
'    Call DeleteBookmarks("zikl") ' �����
'    Call DeleteBookmarks("TehnichUstr") ' ������� ������ ��� �������������
'    ActiveDocument.Bookmarks("VikRezTruboprov").Range.Delete ' ��� ���������� ������� ������������
    ActiveDocument.Variables("punkt7-3").Value = Strings.ChrW(31)
    ActiveDocument.Variables("KIPiA").Value = "��������� " & ActiveDocument.Variables("TechUsrtva").Value & " �������������"
    ActiveDocument.Variables("VnutrIzbP").Value = "����������������"
    ActiveDocument.Variables("punkt7-5-4Mat").Value = Strings.ChrW(31)
    ActiveDocument.Variables.item("PassatT").Value = Strings.ChrW(31)
    ActiveDocument.Variables("ProbnDavlen").Value = "�������"
    ActiveDocument.Variables("SvedOMerPFNP").Value = "26, 27"
End Sub

Sub TehnTruboprovod()
    Call DeleteBookmarks("Rezervuar") ' ������� ����� ��� ����������
'    ActiveDocument.Bookmarks("KotlObor").Range.Delete ' ����� 7.3. ��� ������������ �����
'    ActiveDocument.Bookmarks("R7p4").Range.Delete '����� ��� �����
'    ActiveDocument.Bookmarks("VikRezTruboprov").Range.Delete ' ��� ���������� ������� ������������ � �������� � ������� ���������
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
    ActiveDocument.Variables.item("PassatT").Value = Strings.ChrW(31)
    ActiveDocument.Variables("p2pr8RD3417302").Value = Strings.ChrW(31)
    ActiveDocument.Variables("ProbnDavlen").Value = "�� ��������� � ���������,"
    If UF1.CBtt164.Value = True Then ActiveDocument.Variables("ProbnDavlen").Value = ActiveDocument.Variables("ProbnDavlen").Value & " � �������������� ��������� �� �������������,"
    ActiveDocument.Variables("TTPoverRash").Value = " ����������� ������� ���"
    ActiveDocument.Variables("korpusa").Value = Strings.ChrW(31)
    ActiveDocument.Variables("CBp466-1").Value = " � �. 177" & UF1.FNPTehnTrub.Value
    If ActiveDocument.Bookmarks.Exists("Ne_peremesh") = True Then ActiveDocument.Bookmarks("Ne_peremesh").Range.Delete
    ActiveDocument.Variables("DataIzg6").Value = "���� �������"
   
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
    ActiveDocument.Variables.item("PassatT").Value = Strings.ChrW(31)
    ActiveDocument.Variables("p2pr8RD3417302").Value = Strings.ChrW(31)
    ActiveDocument.Variables("ProbnDavlen").Value = "�� ��������� � ���������,"
    ActiveDocument.Variables("korpusa").Value = Strings.ChrW(31)
    ActiveDocument.Variables("CBp466-1").Value = " � �. 177" & UF1.FNPTehnTrub.Value
End Sub
Sub RezervuarMazut()
    Call VarSoorugenie
'    ActiveDocument.Variables("SposSoedEl").Value = "�������� � ���������� � ��������� ����������:"
    ActiveDocument.Variables("MnNum7").Value = "������� �� ��������������� �����"
    ActiveDocument.Tables(2).Cell(Row:=2, Column:=1).Range = "������ ������, ��"
    ActiveDocument.Tables(2).Cell(Row:=2, Column:=2).Range = UF1.RabocheePRub.Value
    ActiveDocument.Tables(3).Delete
    ActiveDocument.Variables("VKorp").Value = ActiveDocument.Variables("VKorp").Value & ", ������ " & ActiveDocument.Variables("TechUsrtva").Value & " - " & UF1.RaschetnPRub.Value & " ��"
    ActiveDocument.Variables("RabTempP6").Value = ActiveDocument.Variables("RabTempP6").Value & ", ������ ������ - " & Trim(UF1.RabocheePRub.Value) & " ��"
    If UF1.CBFNPPBSNN.Value = True Then ActiveDocument.Variables("PrSrSl-PBSNN").Value = Strings.Chr(13) & "��������� ����� ������ ���������� �������������� � ������������ � ������������ �. 261" & UF1.FNPPESNN.Value
    ActiveDocument.Variables("DataIzg6").Value = "���� �������"
    ActiveDocument.Variables("SvedOMerPFNP").Value = "26, 27"
    ActiveDocument.Variables("VIKRezKontr").Value = "������ ����������." & Strings.Chr(13) _
    & "���������� ������ ����������� ������� ������������ � ���������� � �������� (� ��������� ������) ������. ������ ���������� �� ������ ������ �� " & ActiveDocument.Variables("RaschetntRub").Value & " ������, ������� 1500 �� ������. " & _
    "�� ������� ���������� ����������� 2 ����-����. ��� ���������� ������� ������� ������������ �������� �� ����������. ������������ ��������, ��������� � �����-����� ������������� ����������� ���. " & _
    "������� ��� ������ ����� ������� ������������ ������������ ������� ���� ������� �� ����� ����� 200 ��. ��������� ������� ���������� ������������� ����������� �.�. 8.7, 8.8, 8.13 �� 08-95-95." & _
    Strings.Chr(13) & "����� ����������." & Strings.Chr(13) & _
    "��� ���������� ������� ����� ������������ �������� �� ����������. ��� �������������� �������� ����� ������������ �������� �� ����������. " & _
    "��������� ����� ���������� ������������� ����������� �.�. 8.7, 8.8, 8.15 �� 08-95-95." & _
    Strings.Chr(13) & "������ � ������� �����������." & Strings.Chr(13) & _
    "������ ������ � ������� ����������� ���������� � ���������� � �������� (� ��������� ������) ������ ����������. ��� ���������� ������� ������ ������������ �������� �� ����������. " & _
    "��� ������� ������� ���� ������������ �������� �� ����������. ������ �������� ���� ������� �������� ������� ���� �������� �����, �������� ��� ��������� ������������. " & _
    "��������� ������ ���������� ������������� ����������� �.�. 8.7, 8.8, 8.11 �� 08-95-95." & _
    Strings.Chr(13) & "�������� ����������." & Strings.Chr(13) & _
    "- ���������� ������ ����� ���������� � ����� � ��������� ���� �� ������� ���������� �����������;" & Strings.Chr(13) & _
    "- ����������� �������� ������ �� ��������������;" & Strings.Chr(13) & _
    "- �������� ���������� ����� ����������� ����� (1:10)." & Strings.Chr(13) & _
    "��������� �������� ������������� ����������� �. 5.6 �� 08-95-95." & Strings.Chr(13)
    If UF1.CBRD089595.Value = True And Val(UF1.NaNLet.Value) > 4 Then
        ActiveDocument.Variables("p8Osvid").Value = "��������� ����������� ������������ " & ActiveDocument.Variables("TechUsrtva").Value _
        & ", � ������������ � ������������ �. 3.7.1" & UF1.RD089595.Value & ", ���������� �������� � ���� �� " & Format(DateAdd("yyyy", 4, UF1.AktGID.Value), "dd.mm.yyyy") & "."
    Else
        ActiveDocument.Variables("p8Osvid").Value = Strings.ChrW(31)
    End If
    If UF1.ComboBoxRaschet.ListIndex = 0 Then ActiveDocument.Variables.item("OzOsR").Value = "���� 34233.1-2017, ���-��-03-002-2009"
End Sub

Sub BallGroUst()
    Call MnogCHislo
'    ActiveDocument.Variables("TechUsrtva").Value = "��������"
'    ActiveDocument.Variables("TechUsrtvo").Value = "������� ��������� ���������"
'    ActiveDocument.Variables("TechUsrtvoB").Value = UCase(Left(ActiveDocument.Variables("TechUsrtvo").Value, 1)) & LCase(Mid(ActiveDocument.Variables("TechUsrtvo").Value, 2)) '������ ������ ����� �������
    ActiveDocument.Variables("GOST34347PMat").Value = UF1.GOST9731.Value
    ActiveDocument.Variables.item("PassatT").Value = Strings.ChrW(31)
End Sub
Sub Ballon()
'    ActiveDocument.Variables("TechUsrtva").Value = "��������"
'    ActiveDocument.Variables("TechUsrtvo").Value = "������"
'    ActiveDocument.Variables("TechUsrtvoB").Value = UCase(Left(ActiveDocument.Variables("TechUsrtvo").Value, 1)) & LCase(Mid(ActiveDocument.Variables("TechUsrtvo").Value, 2)) '������ ������ ����� �������
    ActiveDocument.Variables("GOST34347PMat").Value = UF1.GOST9731.Value
    ActiveDocument.Variables.item("PassatT").Value = Strings.ChrW(31)
End Sub

Sub Nasos()
    Dim EdIzP, StansAssoc As String
    EdIzP = " �.���.��."
    If UF1.ComboBoxTipUstroistva.Value = "����������" Then EdIzP = " ���/��" & Strings.ChrW(178) & "."
    ActiveDocument.Tables(2).Cell(Row:=2, Column:=1).Range = "�����," & EdIzP
    ActiveDocument.Tables(2).Cell(Row:=2, Column:=2).Range = UF1.RaschetnP.Value
    ActiveDocument.Tables(2).Cell(Row:=3, Column:=1).Range = "������������������, �" & Strings.ChrW(179) & "/�."
    ActiveDocument.Variables("RaschSreda").Value = Replace(ActiveDocument.Variables("RaschSreda").Value, ",", ".")
    ActiveDocument.Variables("PasportPar").Value = "�������� ����������� ��������������"
    ActiveDocument.Variables("RabSredaToplRasch").Value = Strings.Chr(13) & "������� ����� - " & UF1.RaschSreda.Value & ". " & Strings.Chr(13) & "�����: "
'    ActiveDocument.Variables("RaschSreda").Value = ActiveDocument.Variables("RaschSreda").Value & Strings.Chr(13) & "�����:"
    ActiveDocument.Variables("RaschetnP").Value = UF1.RaschetnP.Value & EdIzP & Strings.Chr(13) & "������������������: "
    ActiveDocument.Variables("Raschetnt").Value = UF1.VKorp.Value & " �" & Strings.ChrW(179) & "/�."
    If UF1.Raschetnt.Value <> "" Then ActiveDocument.Variables("Raschetnt").Value = ActiveDocument.Variables("Raschetnt").Value & Strings.Chr(13) & "����������� ������� �����: " & Trim(UF1.Raschetnt.Value) & Strings.ChrW(176) & "�."
    ActiveDocument.Variables("VKorp").Value = Strings.Chr(13) & "�������� ����������������: " & UF1.RaschetnPRub.Value & " ���." & Strings.Chr(13) & "������� ��������: " & UF1.RaschetntRub.Value & " ��/���"
    If ActiveDocument.Bookmarks.Exists("ORPD10") = True Then ActiveDocument.Bookmarks("ORPD10").Range.Delete
'    ActiveDocument.Variables("NaznTehUstr").Value = ActiveDocument.Variables("NaznTehUstr").Value & ActiveDocument.Variables("RabSreda").Value
    ActiveDocument.Variables("punkt7-3-1").Value = "���������, ���������� � ��������� "
    ActiveDocument.Variables("korpusa").Value = " ��������� �������� ���������"
    ActiveDocument.Variables("ISvarSoed").Value = Strings.ChrW(31)
    ActiveDocument.Variables("p191PBETT").Value = " ������� 1 " & UF1.M2_96.Value
    ActiveDocument.Variables("RezNKPP").Value = "����������� ��������� " & ActiveDocument.Variables("TechUsrtva").Value & " ����������� ��� ��������������� ��� ����� ����������� ������, ��������� ���������� ������ ��������."
    If UF1.CBFNPHOPO.Value = True Then ActiveDocument.Variables("RezNKPP").Value = ActiveDocument.Variables("RezNKPP").Value & " ��������� ������ ������������� ����������� �.�. 15, 132" & UF1.FNPHOPO.Value & "."
    ActiveDocument.Variables("NTDAktVIK").Value = ActiveDocument.Variables("NTDAktVIK").Value & Strings.Chr(13) & UF1.M2_96.Value
    ActiveDocument.Bookmarks("OORAnPr").Range.Delete
    If UF1.CBGOST20816_1.Value = True Then
        ActiveDocument.Variables.item("OzOsR").Value = "�. 6.3.2.3. � ������� �.1 ���������� � " & UF1.GOST20816_1.Value
        ActiveDocument.Variables("OORNasosOzenka").Value = "��� ��������������������� ������ �������� ������ " & Strings.Chr(171) & "���� �" & Strings.Chr(187) & " � " & Strings.Chr(171) & "���� �" & Strings.Chr(187) & " (�������� ��� ���������� ������������), � ������������ � ������������ " & ActiveDocument.Variables.item("OzOsR").Value
        ActiveDocument.Variables("GOST32106").Value = UF1.GOST20816_1.Value
        StansAssoc = ""
    Else
        ActiveDocument.Variables("OORNasosOzenka").Value = "��� ��������������������� ������ �������� ������ " & Strings.Chr(171) & "������" & Strings.Chr(187) & " � " & Strings.Chr(171) & "���������" & Strings.Chr(187) & " (��������� ���������� ���������� ������������), � ������������ � ������������ �. 6.1 � ������� �.1 ���������� � " & UF1.GOST32106.Value
        ActiveDocument.Variables("GOST32106").Value = UF1.GOST32106.Value
        StansAssoc = "; ������������ �. 4.4.2 � ������� �.1 ���������� � " & UF1.SA03_001_05.Value
    End If
    ActiveDocument.Variables("OORNasos").Value = Strings.Chr(13) & "�������� ������ ������� ������ � ������������ ����������-���������������� ��������� (������: ���������� ������� ���������� �������; �������������� ��������� ��������� ������)."
    ActiveDocument.Variables("OORNasos").Value = ActiveDocument.Variables("OORNasos").Value & Strings.Chr(13) & ActiveDocument.Variables("OORNasosOzenka").Value
    ActiveDocument.Variables("OORNasos").Value = ActiveDocument.Variables("OORNasos").Value & StansAssoc
    ActiveDocument.Variables("OORNasos").Value = ActiveDocument.Variables("OORNasos").Value & "." & Strings.Chr(13) & "� ������������ � ������������ " & UF1.M2_96.Value & ", ���������������� ����, ������������� � ���������� ������������ ������������, ������ ������������ ������ ����� ������������ ���������, ���������� �� ����� 30000 �����. � ������ ������ ������������ ��������� ������, �������� ��������� ����� ������ ������, ��� ����������� ������ ������������, �� "
    ActiveDocument.Variables("OORNasos").Value = ActiveDocument.Variables("OORNasos").Value & ActiveDocument.Variables("NaNLet").Value & "."
    ActiveDocument.Variables("VIKRezKontr").Value = "��� ���������� � ������������� �������� ������, ������������ � ������������ �����������, ����������� ������ � ������ ������� ��������, �������������� ���������� ������������, �� ����������."
End Sub

Sub NTDAktVIK()                             '��������� ��� � ����� ��
    tmp = Strings.ChrW(31)
    tmp1 = "����������� ����� � ������� � ������� ������������ ������������ "
    If UF1.CBFNPORPD.Value = True Then      '��� ����
        ActiveDocument.Variables("NTDAktVIK").Value = tmp1 & Mid(UF1.FNPORPDR.Value, 64, 104) & "."
        tmp = Strings.Chr(13)
    End If
    If UF1.CBFNPOPVB.Value = True Then      '��� ����
        ActiveDocument.Variables("NTDAktVIK").Value = ActiveDocument.Variables("NTDAktVIK").Value & tmp & tmp1 & Mid(UF1.FNPOPVBR.Value, 64, 122) & "."
        tmp = Strings.Chr(13)
    End If
    If UF1.CBFNPHOPO.Value = True Then      '��� ����
        ActiveDocument.Variables("NTDAktVIK").Value = ActiveDocument.Variables("NTDAktVIK").Value & tmp & tmp1 & Mid(UF1.FNPHOPO.Value, 64, 66) & "."
        tmp = Strings.Chr(13)
    End If
    If UF1.CBFNPPBETT.Value = True Then     '��� �����
        ActiveDocument.Variables("NTDAktVIK").Value = ActiveDocument.Variables("NTDAktVIK").Value & tmp & tmp1 & Mid(UF1.FNPTehnTrub.Value, 64, 63) & "."
        tmp = Strings.Chr(13)
    End If
    If UF1.CBFNPPBSNN.Value = True Then     '��� ���
        ActiveDocument.Variables("NTDAktVIK").Value = ActiveDocument.Variables("NTDAktVIK").Value & tmp & tmp1 & Mid(UF1.FNPPESNN.Value, 64, 66) & "."
        tmp = Strings.Chr(13)
    End If
    If UF1.CBSO439.Value = True Then        '�� 439 ������
        ActiveDocument.Variables("NTDAktVIK").Value = ActiveDocument.Variables("NTDAktVIK").Value & tmp & Mid(UF1.SO439.Value, 2, 94) & "."
        ActiveDocument.Variables("NTDAktNKPD").Value = Strings.Chr(13) & Mid(UF1.SO439.Value, 2, 94) & "."
        ActiveDocument.Variables("NTDAktNK").Value = Mid(UF1.SO439.Value, 2, 94) & "."
        tmp = Strings.Chr(13)
    End If
    If UF1.CBSO464.Value = True Then        '�� 464 ������������
        ActiveDocument.Variables("NTDAktVIK").Value = ActiveDocument.Variables("NTDAktVIK").Value & tmp & Mid(UF1.SO464.Value, 2, 97) & "."
        ActiveDocument.Variables("NTDAktNKPD").Value = Strings.Chr(13) & Mid(UF1.SO464.Value, 2, 97) & "."
        ActiveDocument.Variables("NTDAktNK").Value = Mid(UF1.SO464.Value, 2, 97) & "."
        tmp = Strings.Chr(13)
    End If
    If UF1.CBSO469.Value = True Then        '�� 469 �����
        ActiveDocument.Variables("NTDAktVIK").Value = ActiveDocument.Variables("NTDAktVIK").Value & tmp & Mid(UF1.SO469.Value, 2, 188) & "."
        ActiveDocument.Variables("NTDAktNKPD").Value = Strings.Chr(13) & Mid(UF1.SO469.Value, 2, 188) & "."
        ActiveDocument.Variables("NTDAktNK").Value = Mid(UF1.SO469.Value, 2, 188) & "."
        tmp = Strings.Chr(13)
    End If
    If UF1.CBGOST34347.Value = True Then        '���� ������ � �������� ��������
        ActiveDocument.Variables("NTDAktVIK").Value = ActiveDocument.Variables("NTDAktVIK").Value & tmp & Mid(UF1.GOST34347.Value, 2, 79) & "."
        tmp = Strings.Chr(13)
    End If
    If UF1.CBRD2626012.Value = True Then        '�� ������� ��2
        ActiveDocument.Variables("NTDAktVIK").Value = ActiveDocument.Variables("NTDAktVIK").Value & tmp & Mid(UF1.GOST34347.Value, 2, 79) & "."
        tmp = Strings.Chr(13)
    End If
    If UF1.CBVM030104.Value = True Then        '�� ��������� �������� ������������
        ActiveDocument.Variables("NTDAktVIK").Value = ActiveDocument.Variables("NTDAktVIK").Value & tmp & Mid(UF1.RDVM03.Value, 2, 229) & "."
        tmp = Strings.Chr(13)
    End If
    If UF1.CBRUA93.Value = True Then            '���-93
        ActiveDocument.Variables("NTDAktVIK").Value = ActiveDocument.Variables("NTDAktVIK").Value & tmp & Mid(UF1.RUA93.Value, 2, 140) & "."
        tmp = Strings.Chr(13)
    End If
    If UF1.CBRD1533413752696.Value = True Then  '���� ������� � ������ �����
        ActiveDocument.Variables("NTDAktVIK").Value = ActiveDocument.Variables("NTDAktVIK").Value & tmp & Mid(UF1.RD1533413752696.Value, 2, 104) & "." & Strings.Chr(13) & Mid(UF1.Snip31875.Value, 2, 42) & "."
        If ActiveDocument.Variables("TechUsrtvo").Value = "���������" Then ActiveDocument.Variables("NTDAktVIK").Value = ActiveDocument.Variables("NTDAktVIK").Value & Strings.Chr(13) & Mid(UF1.Snip31875.Value, 2, 42) & "."
        tmp = Strings.Chr(13)
    End If
    If UF1.CBRD089595.Value = True Then         '�� 08-95-95 ����������
        ActiveDocument.Variables("NTDAktVIK").Value = ActiveDocument.Variables("NTDAktVIK").Value & tmp & Mid(UF1.RD089595.Value, 2, 138) & "."
        tmp = Strings.Chr(13)
    End If
End Sub

Sub OformlenBase()
    UF1.Label8.Caption = "���� ������������"
    UF1.Label5.Caption = "����� ���."
    UF1.Label513.Visible = False
    UF1.MontagOrg.Visible = False
    UF1.Label514.Visible = False
    UF1.DataMontaga.Visible = False
    UF1.CBPodNaliv.Visible = True
    UF1.CBVakuum.Visible = True
    UF1.Label475.Caption = "P="
    UF1.VKorp.Visible = True
    UF1.CBRubashka.Visible = True
    UF1.CBRubashka.Caption = "�������"
    UF1.Label480.Visible = False
    UF1.Label480.Caption = "� ������� �="
    UF1.RaschetnPRub.Visible = False
    UF1.Label479.Visible = False
    UF1.Label479.Caption = "t="
    UF1.RaschetntRub.Visible = False
    UF1.RaschetntRub.ControlTipText = "��������� ����������� � �������"
    UF1.Label481.Visible = False
    UF1.VRub.Visible = False
    UF1.Label486.Visible = False
    UF1.RaschSredaRub.Visible = False
    UF1.Label487.Visible = False
    UF1.Label487.Caption = "� ������� �="
    UF1.RabocheePRub.Visible = False
    UF1.Label482.Visible = False
    UF1.Label482.Caption = "t="
    UF1.RabTempRub.Visible = False
    UF1.RabTempRub.ControlTipText = "������� ����������� � �������"
    UF1.Label484.Visible = False
    UF1.RabSredaRub.Visible = False
    UF1.Label491.Visible = False
    UF1.IspitatPRub.Visible = False
    
    UF1.ProtokolVD.Visible = False
    UF1.ProtokolVDD.Visible = False
    UF1.Label465.Visible = False
    UF1.Label466.Visible = False
    UF1.Label511.Caption = "���. ��. ����� ����."
    
    UF1.Label501.Caption = "��������"
    UF1.Label502.Caption = "�����"
    UF1.Label432.Caption = "�������"
    UF1.CBZikl.Visible = True
    UF1.CBZikl.Value = True
    UF1.Label18.Visible = True
    UF1.Label462.Caption = "����.���."
        
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
        UF1.Label418.Caption = "����."
        UF1.Label485.Caption = "����."
    End If

    If UF1.OptionTruboprovod.Value = True Then
        UF1.VKorp.Visible = False
        UF1.CBRubashka.Caption = "���"
        UF1.Label5.Caption = "����. ���."
        UF1.Label501.Caption = "�����"
        UF1.Label502.Caption = "�����"
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
        UF1.Label475.Caption = "�����"
        UF1.Label479.Caption = "��."
        UF1.Label480.Caption = "����. ���."
        UF1.RaschetnPRub.Visible = True
        UF1.RaschetntRub.Visible = True
    End If
End Sub
