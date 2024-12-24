VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF1 
   Caption         =   "Экспертиза"
   ClientHeight    =   8970.001
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14010
   OleObjectBlob   =   "UF1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UF1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public BasePredp As Variant
Public BasePredpIndex As Variant
Public BaseOPO As Variant
Public BaseOPOIndex As Variant
Public BaseUstr As Variant
Public BaseUstrIndex As Variant
Public UstrIndxs As New Collection
Public BaseElements As Variant
Public BaseEPB As Variant
Public BaseRemont As Variant

Private Sub AktAED_Change()
ActiveDocument.Variables("AktAED").Value = UF1.AktAED.Value
If IsDate(UF1.AktAED.Value) Then ActiveDocument.Variables("AktAEData").Value = FormDat(UF1.AktAED.Value)
End Sub

Private Sub AktVIKD_Change()
If IsDate(UF1.AktVIKD.Value) Then
    UF1.AktVIKMKD.Value = UF1.AktVIKD.Value
    UF1.ProtokTolchD.Value = UF1.AktVIKD.Value
    UF1.ZakUZKD.Value = UF1.AktVIKD.Value
    UF1.ZakMPDZDD.Value = UF1.AktVIKD.Value
    UF1.ProtkTVD.Value = UF1.AktVIKD.Value
    UF1.KartOvalnD.Value = UF1.AktVIKD.Value
    UF1.ProgibD.Value = UF1.AktVIKD.Value
    UF1.KontrGibD.Value = UF1.AktVIKD.Value
    UF1.AktGID.Value = UF1.AktVIKD.Value
    UF1.AktAED.Value = UF1.AktVIKD.Value
    ActiveDocument.Variables("AktVIKD").Value = Format(UF1.AktVIKD.Value, "dd.mm.yyyy")
    ActiveDocument.Variables("AktVIKData").Value = FormDat(UF1.AktVIKD.Value)
End If

End Sub

Private Sub AktVIKD_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not IsDate(UF1.AktVIKD.Value) Then MsgBox ("Неправильный формат даты")
End Sub
 
Private Sub AktVIKMK_Change()
If UF1.AktVIKMK.Enabled = True Then
    UF1.ProtokTolch.Value = Val(UF1.AktVIKMK.Value) + 1
    ActiveDocument.Variables("AktVIKMK").Value = UF1.AktVIKMK.Value
Else
    UF1.ProtokTolch.Value = Val(UF1.AktVIKMK.Value)
End If
End Sub

Private Sub AktVIKMKD_Change()
If IsDate(UF1.AktVIKMKD.Value) Then
    ActiveDocument.Variables("AktVIKMKD").Value = UF1.AktVIKMKD.Value
    ActiveDocument.Variables("AktVIKMKData").Value = FormDat(UF1.AktVIKMKD.Value)
End If
End Sub

Private Sub CBFNPHOPO_Click()
If UF1.CBFNPHOPO.Value = True Then
    UF1.ExpertHim.Value = True
Else
    UF1.ExpertHim.Value = False
End If
End Sub

Private Sub CBFNPOPVB_Click()
If UF1.CBFNPOPVB.Value = True Then
    UF1.ExpertHim.Value = True
Else
    UF1.ExpertHim.Value = False
End If
End Sub

Private Sub CBFNPORPD_Click()
If UF1.CBFNPORPD.Value = True Then UF1.ExpertORPD.Value = True
End Sub

Private Sub CBFNPPBETT_Click()
If UF1.CBFNPPBETT.Value = True Then
    UF1.ExpertZS.Value = True
Else
    UF1.ExpertZS.Value = False
End If

End Sub

Private Sub CBFNPPBSNN_Click()
If UF1.CBFNPPBSNN.Value = True Then
    UF1.ExpertZS.Value = True
    UF1.ExpertSNN.Value = True
Else
    UF1.ExpertZS.Value = False
    UF1.ExpertSNN.Value = False
End If
End Sub

Private Sub CBFNPSUG_Click()
If UF1.CBFNPSUG.Value = True Then
    ActiveDocument.Variables("punkt1-1PBSUG").Value = ";" & Strings.Chr(13) & "п.п. 3, 48" & UF1.FNPSUG.Value
    ActiveDocument.Variables("p7-1SUGProdl").Value = "; п.п. 3, 48" & UF1.FNPSUG.Value
    ActiveDocument.Variables("p-8FNPSUG").Value = ";" & UF1.FNPSUG.Value
Else
    ActiveDocument.Variables("punkt1-1PBSUG").Value = Strings.ChrW(31)
    ActiveDocument.Variables("p7-1SUGProdl").Value = Strings.ChrW(31)
    ActiveDocument.Variables("p-8FNPSUG").Value = Strings.ChrW(31)
End If
End Sub

Private Sub CBGOST34347_Click()
If UF1.CBGOST34347.Value = True Then
    ActiveDocument.Variables.Item("GOST34347").Value = UF1.GOST34347.Value
    ActiveDocument.Variables("GOST34347PMat").Value = "; п.п. 4.1.4, 5.9.1" & UF1.GOST34347.Value
    ActiveDocument.Variables("GOST34347PSosEl").Value = ActiveDocument.Variables("TckZpt").Value & " п. 5.10.2" & UF1.GOST34347.Value
    If UF1.CBVakuum.Value = True Or UF1.CBPodNaliv.Value = True Then
        ActiveDocument.Variables("GOST34347PGiI").Value = Strings.ChrW(31)
    Else
        ActiveDocument.Variables("GOST34347PGiI").Value = "; п.п. 7.11.3, 7.11.5, 7.11.10" & UF1.GOST34347.Value
    End If
    ActiveDocument.Variables.Item("GOST34347PPiI").Value = "; п. 7.11.9" & UF1.GOST34347.Value
    ActiveDocument.Variables("NTDAktNKGOST34347").Value = Strings.Chr(13) & Mid(UF1.GOST34347.Value, 2, 79) & "."
Else
    ActiveDocument.Variables.Item("GOST34347").Value = Strings.ChrW(31)
    ActiveDocument.Variables("GOST34347PSosEl").Value = Strings.ChrW(31)
    ActiveDocument.Variables("GOST34347PMat").Value = Strings.ChrW(31)
    ActiveDocument.Variables("GOST34347PGiI").Value = Strings.ChrW(31)
    ActiveDocument.Variables.Item("GOST34347PPiI").Value = Strings.ChrW(31)
    ActiveDocument.Variables("NTDAktNKGOST34347").Value = Strings.ChrW(31)
End If
End Sub

Private Sub CBp100_Click()
If UF1.CBp100.Value = True Then
ActiveDocument.Variables("p100FNPORPD").Value = " п. 100" & UF1.FNPORPDR.Value
End If
If UF1.CBp100.Value = False Then
ActiveDocument.Variables("p100FNPORPD").Value = Strings.ChrW(31)
End If
End Sub

Private Sub CBp177_Click()
If UF1.CBp177.Value = True Then If UF1.PnIs.Value = True Then UF1.PnIs.Value = False
End Sub

Private Sub CBp178_Click()
If UF1.CBp178.Value = True Then If UF1.PnIs.Value = True Then UF1.PnIs.Value = False
End Sub

Private Sub CBp179_Click()
If UF1.CBp179.Value = True Then If UF1.PnIs.Value = True Then UF1.PnIs.Value = False
End Sub

Private Sub CBp185_Click()
If UF1.CBp185.Value = True Then If UF1.PnIs.Value = True Then UF1.PnIs.Value = False
End Sub

Private Sub CBp186_Click()
If UF1.CBp186.Value = True Then If UF1.PnIs.Value = True Then UF1.PnIs.Value = False
End Sub

Private Sub CBp187_Click()
If UF1.CBp187.Value = True Then If UF1.PnIs.Value = True Then UF1.PnIs.Value = False
End Sub

Private Sub CBp188_Click()
If UF1.CBp188.Value = True Then If UF1.PnIs.Value = True Then UF1.PnIs.Value = False
End Sub

Private Sub CBp348_Click()
If UF1.CBp348.Value = True Then
    ActiveDocument.Variables("PredKlNet").Value = " Предохранительный клапан на сосуде не установлен, его установка не обязательна, так как рабочее давление в сосуде больше давления питающего источника, что соответствует требованиям п. 348" & UF1.FNPORPDR.Value & "."
    UF1.CBp338.Value = False
    UF1.CBp339.Value = False
    UF1.CBp340.Value = False
    UF1.CBp341.Value = False
    UF1.CBp343.Value = False
    UF1.CBp353.Value = False
Else
    ActiveDocument.Variables("PredKlNet").Value = Strings.ChrW(31)
End If
End Sub

Private Sub CBp466_Click()
If UF1.CBp466.Value = True Then
    ActiveDocument.Variables("CBp466-2").Value = " в соответствии с п. 466" & UF1.FNPORPDR.Value
    ActiveDocument.Variables("CBp466-1").Value = " и п. 466" & UF1.FNPORPDR.Value
    ActiveDocument.Variables("CBp466data").Value = "30.09." & Year(Date)
Else
    ActiveDocument.Variables("CBp466-1").Value = Strings.ChrW(31)
End If
End Sub

Private Sub CBPodNaliv_Click()
    BaseUstr(BaseUstrIndex, 15) = Trim$(UF1.CBPodNaliv.Value)
    If UF1.CBPodNaliv.Value = True And UF1.CBVakuum.Value = True Then UF1.CBVakuum.Value = False
    If UF1.CBPodNaliv.Value = True Then
        UF1.CBSO439.Value = False
        ActiveDocument.Variables("p2-102RUA").Value = Strings.ChrW(31)
        ActiveDocument.Variables("p2-111RUA").Value = " 2.111,"
        UF1.CBRUA93.Value = True
        ActiveDocument.Variables("GOST34347PGiI").Value = Strings.ChrW(31)
        ActiveDocument.Variables("TipTehUstr").Value = "сосуд, работающий без давления (под налив)"
        UF1.NaznTehUstr.Value = "емкость"
        ActiveDocument.Variables("IspitatP").Value = "полный налив"
'        UF1.RabocheeP.Value = "под налив"
        UF1.RazreshaemoeP.Value = "под налив"
        UF1.IspitatP.Value = "наливом"
    Else
        ActiveDocument.Variables("TipTehUstr").Value = "сосуд, работающий под давлением"
'        ActiveDocument.Variables("NaznTehUstr").Value = "сосуд для "
        ActiveDocument.Variables("IndxP").Value = Strings.ChrW(31)
        UF1.CBSO439.Value = True
    End If
End Sub

Private Sub CBR10_Click()
    If UF1.CBR10.Value = True Then
        UF1.R10C1.Enabled = True
        UF1.R10C2.Enabled = True
        UF1.R10C3.Enabled = True
        UF1.R10C4.Enabled = True
        UF1.R10C5.Enabled = True
        UF1.R10C6.Enabled = True
        UF1.R10C7.Enabled = True
    Else
        UF1.R10C1.Enabled = False
        UF1.R10C2.Enabled = False
        UF1.R10C3.Enabled = False
        UF1.R10C4.Enabled = False
        UF1.R10C5.Enabled = False
        UF1.R10C6.Enabled = False
        UF1.R10C7.Enabled = False
    End If
    BaseElements(BaseUstrIndex, 75 + (8 * (UF1.SpButElm.Value))) = Trim$(UF1.CBR10.Value)
End Sub

Private Sub CBR2_Change()
    If UF1.CBR2.Value = True Then
        UF1.R2C1.Enabled = True
        UF1.R2C2.Enabled = True
        UF1.R2C3.Enabled = True
        UF1.R2C4.Enabled = True
        UF1.R2C5.Enabled = True
        UF1.R2C6.Enabled = True
        UF1.R2C7.Enabled = True
    Else
        UF1.R2C1.Enabled = False
        UF1.R2C2.Enabled = False
        UF1.R2C3.Enabled = False
        UF1.R2C4.Enabled = False
        UF1.R2C5.Enabled = False
        UF1.R2C6.Enabled = False
        UF1.R2C7.Enabled = False
    End If
    If IsNull(UF1.CBR2.Value) Then
        BaseElements(BaseUstrIndex, 11 + (8 * (UF1.SpButElm.Value))) = UF1.CBR2.Value
    Else
        BaseElements(BaseUstrIndex, 11 + (8 * (UF1.SpButElm.Value))) = Trim$(UF1.CBR2.Value)
    End If
End Sub

Private Sub CBR3_Click()
    If UF1.CBR3.Value = True Then
        UF1.R3C1.Enabled = True
        UF1.R3C2.Enabled = True
        UF1.R3C3.Enabled = True
        UF1.R3C4.Enabled = True
        UF1.R3C5.Enabled = True
        UF1.R3C6.Enabled = True
        UF1.R3C7.Enabled = True
    Else
        UF1.R3C1.Enabled = False
        UF1.R3C2.Enabled = False
        UF1.R3C3.Enabled = False
        UF1.R3C4.Enabled = False
        UF1.R3C5.Enabled = False
        UF1.R3C6.Enabled = False
        UF1.R3C7.Enabled = False
    End If
    BaseElements(BaseUstrIndex, 19 + (8 * (UF1.SpButElm.Value))) = Trim$(UF1.CBR3.Value)
End Sub

Private Sub CBR4_Click()
    If UF1.CBR4.Value = True Then
        UF1.R4C1.Enabled = True
        UF1.R4C2.Enabled = True
        UF1.R4C3.Enabled = True
        UF1.R4C4.Enabled = True
        UF1.R4C5.Enabled = True
        UF1.R4C6.Enabled = True
        UF1.R4C7.Enabled = True
    Else
        UF1.R4C1.Enabled = False
        UF1.R4C2.Enabled = False
        UF1.R4C3.Enabled = False
        UF1.R4C4.Enabled = False
        UF1.R4C5.Enabled = False
        UF1.R4C6.Enabled = False
        UF1.R4C7.Enabled = False
    End If
    BaseElements(BaseUstrIndex, 27 + (8 * (UF1.SpButElm.Value))) = Trim$(UF1.CBR4.Value)
End Sub

Private Sub CBR5_Click()
    If UF1.CBR5.Value = True Then
        UF1.R5C1.Enabled = True
        UF1.R5C2.Enabled = True
        UF1.R5C3.Enabled = True
        UF1.R5C4.Enabled = True
        UF1.R5C5.Enabled = True
        UF1.R5C6.Enabled = True
        UF1.R5C7.Enabled = True
    Else
        UF1.R5C1.Enabled = False
        UF1.R5C2.Enabled = False
        UF1.R5C3.Enabled = False
        UF1.R5C4.Enabled = False
        UF1.R5C5.Enabled = False
        UF1.R5C6.Enabled = False
        UF1.R5C7.Enabled = False
    End If
    BaseElements(BaseUstrIndex, 35 + (8 * (UF1.SpButElm.Value))) = Trim$(UF1.CBR5.Value)
End Sub

Private Sub CBR6_Click()
    If UF1.CBR6.Value = True Then
        UF1.R6C1.Enabled = True
        UF1.R6C2.Enabled = True
        UF1.R6C3.Enabled = True
        UF1.R6C4.Enabled = True
        UF1.R6C5.Enabled = True
        UF1.R6C6.Enabled = True
        UF1.R6C7.Enabled = True
    Else
        UF1.R6C1.Enabled = False
        UF1.R6C2.Enabled = False
        UF1.R6C3.Enabled = False
        UF1.R6C4.Enabled = False
        UF1.R6C5.Enabled = False
        UF1.R6C6.Enabled = False
        UF1.R6C7.Enabled = False
    End If
    BaseElements(BaseUstrIndex, 43 + (8 * (UF1.SpButElm.Value))) = Trim$(UF1.CBR6.Value)
End Sub

Private Sub CBR7_Click()
    If UF1.CBR7.Value = True Then
        UF1.R7C1.Enabled = True
        UF1.R7C2.Enabled = True
        UF1.R7C3.Enabled = True
        UF1.R7C4.Enabled = True
        UF1.R7C5.Enabled = True
        UF1.R7C6.Enabled = True
        UF1.R7C7.Enabled = True
    Else
        UF1.R7C1.Enabled = False
        UF1.R7C2.Enabled = False
        UF1.R7C3.Enabled = False
        UF1.R7C4.Enabled = False
        UF1.R7C5.Enabled = False
        UF1.R7C6.Enabled = False
        UF1.R7C7.Enabled = False
    End If
    BaseElements(BaseUstrIndex, 51 + (8 * (UF1.SpButElm.Value))) = Trim$(UF1.CBR7.Value)
End Sub

Private Sub CBR8_Click()
    If UF1.CBR8.Value = True Then
        UF1.R8C1.Enabled = True
        UF1.R8C2.Enabled = True
        UF1.R8C3.Enabled = True
        UF1.R8C4.Enabled = True
        UF1.R8C5.Enabled = True
        UF1.R8C6.Enabled = True
        UF1.R8C7.Enabled = True
    Else
        UF1.R8C1.Enabled = False
        UF1.R8C2.Enabled = False
        UF1.R8C3.Enabled = False
        UF1.R8C4.Enabled = False
        UF1.R8C5.Enabled = False
        UF1.R8C6.Enabled = False
        UF1.R8C7.Enabled = False
    End If
    BaseElements(BaseUstrIndex, 59 + (8 * (UF1.SpButElm.Value))) = Trim$(UF1.CBR8.Value)
End Sub

Private Sub CBR9_Click()
    If UF1.CBR9.Value = True Then
        UF1.R9C1.Enabled = True
        UF1.R9C2.Enabled = True
        UF1.R9C3.Enabled = True
        UF1.R9C4.Enabled = True
        UF1.R9C5.Enabled = True
        UF1.R9C6.Enabled = True
        UF1.R9C7.Enabled = True
    Else
        UF1.R9C1.Enabled = False
        UF1.R9C2.Enabled = False
        UF1.R9C3.Enabled = False
        UF1.R9C4.Enabled = False
        UF1.R9C5.Enabled = False
        UF1.R9C6.Enabled = False
        UF1.R9C7.Enabled = False
    End If
    BaseElements(BaseUstrIndex, 67 + (8 * (UF1.SpButElm.Value))) = Trim$(UF1.CBR9.Value)
End Sub

Private Sub CBRD089595_Click()
If UF1.CBRD089595.Value = True Then
    ActiveDocument.Variables("SostKorpRD089595").Value = " п.п. 8.7, 8.8, 8.13" & UF1.RD089595.Value
    ActiveDocument.Variables("DnichRezHlop").Value = " п. 8.15" & UF1.RD089595.Value
    ActiveDocument.Variables("DnichRez").Value = " п.п. 8.7, 8.8, 8.15" & UF1.RD089595.Value
    ActiveDocument.Variables("KrovlRez").Value = " п.п. 8.7, 8.8, 8.11" & UF1.RD089595.Value
    ActiveDocument.Variables("OtmRez").Value = " п. 5.6" & UF1.RD089595.Value
Else
    ActiveDocument.Variables("SostKorpRD089595").Value = Strings.ChrW(31)
    ActiveDocument.Variables("DnichRezHlop").Value = Strings.ChrW(31)
    ActiveDocument.Variables("DnichRez").Value = Strings.ChrW(31)
    ActiveDocument.Variables("KrovlRez").Value = Strings.ChrW(31)
    ActiveDocument.Variables("OtmRez").Value = Strings.ChrW(31)
End If
End Sub

Private Sub CBRD1533413752696_Change()
If UF1.CBRD1533413752696.Value = True Then
    ActiveDocument.Variables("NTDAktNKRD1533413752696").Value = Strings.Chr(13) & Mid(UF1.RD1533413752696.Value, 2, 104) & "." & Strings.Chr(13) & Mid(UF1.Snip31875.Value, 2, 42) & "."
    ActiveDocument.Variables("SostKorpRDKisl").Value = " п.п. 11.2, 11.3, 11.9, 11.10 " & UF1.RD1533413752696.Value
Else
    ActiveDocument.Variables("NTDAktNKRD1533413752696").Value = Strings.ChrW(31)
    ActiveDocument.Variables("SostKorpRDKisl").Value = Strings.ChrW(31)
End If
End Sub

Private Sub CBRD2626012_Click()
If UF1.CBRD2626012.Value = True Then
    ActiveDocument.Variables("ovalnrd").Value = ActiveDocument.Variables("ovalnrd").Value & "; п. 5.4.3.2" & UF1.RD26_260.Value
    ActiveDocument.Variables("NTDAktNKRD2626012").Value = Strings.Chr(13) & Mid(UF1.RD26_260.Value, 2, 104) & "."
Else
    ActiveDocument.Variables("NTDAktNKRD2626012").Value = Strings.ChrW(31)
End If
End Sub

Private Sub CBRekpoTT_Click()
If UF1.CBRekpoTT.Value = True Then
    ActiveDocument.Variables("VikTT").Value = "; п.п. 332, 334, 435" & UF1.RekpoTT.Value
    ActiveDocument.Variables("UZKRekTT").Value = " п. 343" & UF1.RekpoTT.Value
    ActiveDocument.Variables("ZDRekTT").Value = " п. 344" & UF1.RekpoTT.Value
    ActiveDocument.Variables("GIRekTT").Value = "; п.п. 374, 375, 379, 384" & UF1.RekpoTT.Value
Else
    ActiveDocument.Variables("VikTT").Value = Strings.ChrW(31)
    ActiveDocument.Variables("UZKRekTT").Value = Strings.ChrW(31)
    ActiveDocument.Variables("ZDRekTT").Value = Strings.ChrW(31)
    ActiveDocument.Variables("GIRekTT").Value = Strings.ChrW(31)
End If
End Sub

Private Sub CBRUA93_Click()
If UF1.CBRUA93.Value = True Then
    ActiveDocument.Variables("p4-5RUA").Value = "; п.п. 2.51, 2.53, 2.54" & UF1.RUA93.Value
    ActiveDocument.Variables("GIRUA93").Value = "; п.п." & ActiveDocument.Variables("p2-102RUA").Value & ActiveDocument.Variables("p2-111RUA").Value & " 2.113" & UF1.RUA93.Value
    ActiveDocument.Variables("UZKRUA").Value = "; п. 2.84" & UF1.RUA93.Value
    ActiveDocument.Variables("NTDAktNKRUA93").Value = Strings.Chr(13) & Mid(UF1.RUA93.Value, 2, 140) & "."
Else
    ActiveDocument.Variables("NTDAktNKRUA93").Value = Strings.ChrW(31)
    ActiveDocument.Variables("p4-5RUA").Value = Strings.ChrW(31)
    ActiveDocument.Variables("GIRUA93").Value = Strings.ChrW(31)
    ActiveDocument.Variables("UZKRUA").Value = Strings.ChrW(31)
End If
End Sub

Private Sub CBRubashka_Change()
    If IsNull(UF1.CBRubashka.Value) Then
        BaseUstr(BaseUstrIndex, 21) = UF1.CBRubashka.Value
    Else
        BaseUstr(BaseUstrIndex, 21) = Trim$(UF1.CBRubashka.Value)
    End If
    If UF1.CBRubashka.Value = False Then
        UF1.RaschetnPRub.Enabled = False
        UF1.RaschetntRub.Enabled = False
        UF1.VRub.Enabled = False
        UF1.RaschSredaRub.Enabled = False
        UF1.RabocheePRub.Enabled = False
        UF1.RabTempRub.Enabled = False
        UF1.RabSredaRub.Enabled = False
        UF1.IspitatPRub.Enabled = False
    Else
        If UF1.CBRubashka.Value = True Then
            UF1.CBRubashka.Caption = "Рубашка"
        Else
            UF1.CBRubashka.Caption = "Труб.сис."
        End If
        If UF1.OptionTruboprovod.Value = True Then
            UF1.CBRubashka.Caption = "РОУ"
            UF1.Label480.Caption = "после РОУ Р="
            UF1.Label487.Caption = "после РОУ Р="
        End If
        UF1.RaschetnPRub.Enabled = True
        UF1.RaschetntRub.Enabled = True
        If UF1.OptionSosud.Value = True Then UF1.VRub.Enabled = True
        UF1.RaschSredaRub.Enabled = True
        UF1.RabocheePRub.Enabled = True
        UF1.RabTempRub.Enabled = True
        UF1.RabSredaRub.Enabled = True
        UF1.IspitatPRub.Enabled = True
    End If
End Sub

Private Sub CBSO439_Click()
If UF1.CBSO439.Value = True Then
    ActiveDocument.Variables("VIKrdSO439").Value = "; п.п. 5.4, 5.7, 5.8, 5.10" & UF1.SO439.Value
    ActiveDocument.Variables("tverdSO439").Value = " п. 5.9" & UF1.SO439.Value
    ActiveDocument.Variables("UZKSO439").Value = "; п. 5.14" & UF1.SO439.Value
    ActiveDocument.Variables("ovalnrd").Value = "При измерениях овальности обечайки " & ActiveDocument.Variables("TechUsrtva").Value & " установлено, что овальность соответствует требованиям п. 5.6." & UF1.SO439.Value
    ActiveDocument.Variables("NTDAktNKPD").Value = Strings.Chr(13) & Mid(UF1.SO439.Value, 2, 94) & "."
    ActiveDocument.Variables("NTDAktNK").Value = Mid(UF1.SO439.Value, 2, 94) & "."
Else
    ActiveDocument.Variables("NTDAktNKPD").Value = Strings.ChrW(31)
    ActiveDocument.Variables("NTDAktNK").Value = Strings.ChrW(31)
    ActiveDocument.Variables("VIKrdSO439").Value = Strings.ChrW(31)
    ActiveDocument.Variables("UZKSO439").Value = Strings.ChrW(31)
    ActiveDocument.Variables("tverdSO439").Value = Strings.ChrW(31)
    ActiveDocument.Variables("ovalnrd").Value = Strings.ChrW(31)
End If
End Sub

Private Sub CBSO464_Click()
If UF1.CBSO464.Value = True Then
    ActiveDocument.Variables("VIKrdSO464").Value = "; п.п. 5.1, 5.2, 5.3, 5.6, 5.10, 5.12, 5.13, 5.15, 5.18, 5.20" & UF1.SO464.Value
    ActiveDocument.Variables("TolSte464").Value = "; п. 5.7" & UF1.SO464.Value
    ActiveDocument.Variables("UZKSO469").Value = "; п. 5.19" & UF1.SO464.Value
    ActiveDocument.Variables("tverdSO464").Value = " п. 5.14" & UF1.SO464.Value
    ActiveDocument.Variables("GISO464").Value = "; раздела 4.8" & UF1.SO464.Value
    ActiveDocument.Variables("NTDAktNKPD").Value = Strings.Chr(13) & Mid(UF1.SO464.Value, 2, 97) & "."
    ActiveDocument.Variables("NTDAktNK").Value = Mid(UF1.SO464.Value, 2, 97) & "."
Else
    ActiveDocument.Variables("VIKrdSO464").Value = Strings.ChrW(31)
    ActiveDocument.Variables("TolSte464").Value = Strings.ChrW(31)
    ActiveDocument.Variables("UZKSO469").Value = Strings.ChrW(31)
    ActiveDocument.Variables("tverdSO464").Value = Strings.ChrW(31)
    ActiveDocument.Variables("GISO464").Value = Strings.ChrW(31)
    ActiveDocument.Variables("NTDAktNKPD").Value = Strings.ChrW(31)
    ActiveDocument.Variables("NTDAktNK").Value = Strings.ChrW(31)
End If
End Sub

Private Sub CBSO469_Click()
If UF1.CBSO469.Value = True Then
    ActiveDocument.Variables("VIKrdSO469").Value = "; п.п. 5.4, 5.5, 5.15" & UF1.SO469.Value
    ActiveDocument.Variables("tverdSO469").Value = " п. 5.29" & UF1.SO469.Value
    ActiveDocument.Variables("ovalnKG").Value = " п. 5.11" & UF1.SO469.Value
    ActiveDocument.Variables("ProgibCO").Value = " п. 5.13" & UF1.SO469.Value
    ActiveDocument.Variables("ovalnrd").Value = "При измерениях овальности барабанов котла установлено, что овальность соответствует требованиям п. 5.10." & UF1.SO469.Value
    ActiveDocument.Variables("NTDAktNKPD").Value = Strings.Chr(13) & Mid(UF1.SO469.Value, 2, 188) & "."
    ActiveDocument.Variables("NTDAktNK").Value = Mid(UF1.SO469.Value, 2, 188) & "."
Else
    ActiveDocument.Variables("VIKrdSO469").Value = Strings.ChrW(31)
    ActiveDocument.Variables("tverdSO469").Value = Strings.ChrW(31)
    ActiveDocument.Variables("ovalnKG").Value = Strings.ChrW(31)
    ActiveDocument.Variables("ProgibCO").Value = Strings.ChrW(31)
    ActiveDocument.Variables("ovalnrd").Value = Strings.ChrW(31)
    ActiveDocument.Variables("NTDAktNKPD").Value = Strings.ChrW(31)
    ActiveDocument.Variables("NTDAktNK").Value = Strings.ChrW(31)
End If
End Sub

Private Sub CBVakuum_Click()
    BaseUstr(BaseUstrIndex, 16) = Trim$(UF1.CBVakuum.Value)
    If UF1.CBPodNaliv.Value = True And UF1.CBVakuum.Value = True Then UF1.CBPodNaliv.Value = False
    If UF1.CBVakuum.Value = True Then
        UF1.CBSO439.Value = False
        ActiveDocument.Variables("p2-102RUA").Value = " 2.102,"
        ActiveDocument.Variables("p2-111RUA").Value = Strings.ChrW(31)
        If UF1.OptionSosud.Value = True Then UF1.CBRUA93.Value = True
        ActiveDocument.Variables("GOST34347PGiI").Value = Strings.ChrW(31)
        ActiveDocument.Variables("TipTehUstr").Value = "сосуд, работающий без давления (под вакуумом)"
'        ActiveDocument.Variables("NaznTehUstr").Value = "вакуумная емкость"
'        UF1.Label14.Caption = "Раб. абс."
'        uf1.Label15.Caption = "Разреш. абс."
        ActiveDocument.Variables("VnutrIzbP").Value = "наружным"
        ActiveDocument.Variables("VnutrP").Value = "наружного"
        If UF1.ComboBoxTipUstroistva.Value = "технологический трубопровод" Then
            UF1.IspitatP.Value = "2,0"
'            ActiveDocument.Variables("NaznTehUstr").Value = "транспортировать "
            ActiveDocument.Variables("TipTehUstr").Value = "технологический трубопровод"
        Else
            UF1.IspitatP.Value = "1,25"
        End If
        UF1.CBtt144.Value = True
    Else
'        UF1.Label14.Caption = "Рабочее P"
'        uf1.Label15.Caption = "Разреш. P"
        ActiveDocument.Variables("TipTehUstr").Value = "сосуд, работающий под давлением"
'        ActiveDocument.Variables("NaznTehUstr").Value = "сосуд для "
'        ActiveDocument.Variables("IndxP").Value = Strings.ChrW(31)
        UF1.CBSO439.Value = True
        UF1.CBtt144.Value = False
        ActiveDocument.Variables("VnutrIzbP").Value = "внутренним избыточным"
        ActiveDocument.Variables("VnutrP").Value = "внутреннего"
    End If
End Sub

Private Sub CBVM030104_Click()
If UF1.CBVM030104.Value = True Then
    ActiveDocument.Variables("p4-5VM").Value = "; п. 4.5" & UF1.RDVM03.Value
    ActiveDocument.Variables("p4-7VM").Value = "; п. 4.7" & UF1.RDVM03.Value
    ActiveDocument.Variables("NTDAktNKVM030104").Value = Strings.Chr(13) & Mid(UF1.RDVM03.Value, 2, 229) & "."
Else
    ActiveDocument.Variables("NTDAktNKVM030104").Value = Strings.ChrW(31)
    ActiveDocument.Variables("p4-5VM").Value = Strings.ChrW(31)
    ActiveDocument.Variables("p4-7VM").Value = Strings.ChrW(31)
End If
End Sub

Private Sub CBZikl_Change()
    If IsNull(UF1.CBZikl.Value) Then
        BaseElements(BaseUstrIndex, 163) = "False"
    Else
        BaseElements(BaseUstrIndex, 163) = Trim$(UF1.CBZikl.Value)
    End If
    If UF1.CBZikl.Value = True Then
        UF1.KolZicl.Enabled = True
    Else
        UF1.KolZicl.Enabled = False
    End If
End Sub

Private Sub ClassOpasOPO_Change()
    If UF1.ClassOpasOPO.ListIndex <> -1 And UF1.ClassOpasOPO.ListCount = UF1.RegNOPO.ListCount Then UF1.RegNOPO.ListIndex = UF1.ClassOpasOPO.ListIndex
    ActiveDocument.Variables("ClassOpasOPO").Value = ", " & Trim(UF1.ClassOpasOPO.Value)
    If Trim(UF1.ClassOpasOPO.Value) = "" Then ActiveDocument.Variables("ClassOpasOPO").Value = Strings.ChrW(31)
End Sub

Private Sub ClassOpasOPO_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    BaseOPO(BaseOPOIndex, 3) = Trim(UF1.ClassOpasOPO.Value)
End Sub

Private Sub ComboBoxRaschet_Change()
If UF1.ComboBoxRaschet.ListIndex = 0 Then
    ActiveDocument.Variables.Item("PassatT").Value = ", для проведения поверочного расчета на прочность использована программа " & Strings.Chr(171) & "Пассат" & Strings.Chr(187) & ", разработанная ООО " & Strings.Chr(171) & "НТП Трубопровод" & Strings.Chr(187)
    ActiveDocument.Variables.Item("OzOsR").Value = "ГОСТ 34233.1-2017, ГОСТ 34233.2-2017, ГОСТ 34233.6-2017"
End If
If UF1.ComboBoxRaschet.ListIndex = 1 Then
    ActiveDocument.Variables.Item("PassatT").Value = Strings.ChrW(31)
    ActiveDocument.Variables.Item("OzOsR").Value = Trim(UF1.RD1024998.Value)
End If
If UF1.ComboBoxRaschet.ListIndex = 2 Then
    ActiveDocument.Variables.Item("PassatT").Value = Strings.ChrW(31)
    ActiveDocument.Variables.Item("OzOsR").Value = "п. 11.10" & UF1.RD1533413752696.Value
End If
If UF1.ComboBoxRaschet.ListIndex = 3 Then
    ActiveDocument.Variables.Item("PassatT").Value = Strings.ChrW(31)
    ActiveDocument.Variables("OzOsR").Value = Trim(UF1.GOST32388.Value)
End If
If UF1.ComboBoxRaschet.ListIndex = 4 Then
    ActiveDocument.Variables.Item("PassatT").Value = Strings.ChrW(31)
    ActiveDocument.Variables.Item("OzOsR").Value = UF1.GOST25215.Value
End If
If UF1.ComboBoxRaschet.ListIndex = 5 Then
    ActiveDocument.Variables.Item("PassatT").Value = Strings.ChrW(31)
    ActiveDocument.Variables.Item("OzOsR").Value = " справочника " & Strings.Chr(171) & "Основы конструирования и расчета химической аппаратуры" & Strings.Chr(187)
End If
If UF1.ComboBoxRaschet.ListIndex = 6 Then
    ActiveDocument.Variables.Item("PassatT").Value = Strings.ChrW(31)
    ActiveDocument.Variables.Item("OzOsR").Value = "п. 6.1 и таблицы А.1 приложения А " & UF1.GOST32106.Value
End If
    Call RaschOstRes
End Sub

Private Sub ComboBoxTechUsrtvo_Change()
If UF1.ComboBoxTechUsrtvo.ListIndex = 0 Then
        ActiveDocument.Variables("TechUsrtva").Value = "котла"
        ActiveDocument.Variables("TechUsrtvo").Value = "котел"
        ActiveDocument.Variables("TechUsrtvoB").Value = UCase(Left(ActiveDocument.Variables("TechUsrtvo").Value, 1)) & LCase(Mid(ActiveDocument.Variables("TechUsrtvo").Value, 2)) 'Делаем первую букву большой
End If
If UF1.ComboBoxTechUsrtvo.ListIndex = 1 Then
        ActiveDocument.Variables("TechUsrtva").Value = "экономайзера"
        ActiveDocument.Variables("TechUsrtvo").Value = "экономайзер"
        ActiveDocument.Variables("TechUsrtvoB").Value = UCase(Left(ActiveDocument.Variables("TechUsrtvo").Value, 1)) & LCase(Mid(ActiveDocument.Variables("TechUsrtvo").Value, 2)) 'Делаем первую букву большой
End If
If UF1.ComboBoxTechUsrtvo.ListIndex = 2 Then
        ActiveDocument.Variables("TechUsrtva").Value = "сосуда"
        ActiveDocument.Variables("TechUsrtvo").Value = "сосуд"
        ActiveDocument.Variables("TechUsrtvoB").Value = UCase(Left(ActiveDocument.Variables("TechUsrtvo").Value, 1)) & LCase(Mid(ActiveDocument.Variables("TechUsrtvo").Value, 2)) 'Делаем первую букву большой
End If
If UF1.ComboBoxTechUsrtvo.ListIndex = 3 Then
        ActiveDocument.Variables("TechUsrtva").Value = "трубопровода"
        ActiveDocument.Variables("TechUsrtvo").Value = "трубопровод"
        ActiveDocument.Variables("TechUsrtvoB").Value = UCase(Left(ActiveDocument.Variables("TechUsrtvo").Value, 1)) & LCase(Mid(ActiveDocument.Variables("TechUsrtvo").Value, 2)) 'Делаем первую букву большой
End If
If UF1.ComboBoxTechUsrtvo.ListIndex = 4 Then
        ActiveDocument.Variables("TechUsrtva").Value = "воздухосборника"
        ActiveDocument.Variables("TechUsrtvo").Value = "воздухосборник"
        ActiveDocument.Variables("TechUsrtvoB").Value = UCase(Left(ActiveDocument.Variables("TechUsrtvo").Value, 1)) & LCase(Mid(ActiveDocument.Variables("TechUsrtvo").Value, 2)) 'Делаем первую букву большой
End If
If UF1.ComboBoxTechUsrtvo.ListIndex = 5 Then
        ActiveDocument.Variables("TechUsrtva").Value = "газификатора"
        ActiveDocument.Variables("TechUsrtvo").Value = "газификатор"
        ActiveDocument.Variables("TechUsrtvoB").Value = UCase(Left(ActiveDocument.Variables("TechUsrtvo").Value, 1)) & LCase(Mid(ActiveDocument.Variables("TechUsrtvo").Value, 2)) 'Делаем первую букву большой
End If
If UF1.ComboBoxTechUsrtvo.ListIndex = 6 Then
        ActiveDocument.Variables("TechUsrtva").Value = "подогревателя"
        ActiveDocument.Variables("TechUsrtvo").Value = "подогреватель"
        ActiveDocument.Variables("TechUsrtvoB").Value = UCase(Left(ActiveDocument.Variables("TechUsrtvo").Value, 1)) & LCase(Mid(ActiveDocument.Variables("TechUsrtvo").Value, 2)) 'Делаем первую букву большой
End If
If UF1.ComboBoxTechUsrtvo.ListIndex = 7 Then
        ActiveDocument.Variables("TechUsrtva").Value = "автоклава"
        ActiveDocument.Variables("TechUsrtvo").Value = "автоклав"
        ActiveDocument.Variables("TechUsrtvoB").Value = UCase(Left(ActiveDocument.Variables("TechUsrtvo").Value, 1)) & LCase(Mid(ActiveDocument.Variables("TechUsrtvo").Value, 2)) 'Делаем первую букву большой
End If
If UF1.ComboBoxTechUsrtvo.ListIndex = 8 Then
        ActiveDocument.Variables("TechUsrtva").Value = "емкости"
        ActiveDocument.Variables("TechUsrtvo").Value = "емкость"
        ActiveDocument.Variables("TechUsrtvoB").Value = UCase(Left(ActiveDocument.Variables("TechUsrtvo").Value, 1)) & LCase(Mid(ActiveDocument.Variables("TechUsrtvo").Value, 2)) 'Делаем первую букву большой
End If
If UF1.ComboBoxTechUsrtvo.ListIndex = 9 Then
        ActiveDocument.Variables("TechUsrtva").Value = "баллона"
        ActiveDocument.Variables("TechUsrtvo").Value = "баллон"
        ActiveDocument.Variables("TechUsrtvoB").Value = UCase(Left(ActiveDocument.Variables("TechUsrtvo").Value, 1)) & LCase(Mid(ActiveDocument.Variables("TechUsrtvo").Value, 2)) 'Делаем первую букву большой
End If
If UF1.ComboBoxTechUsrtvo.ListIndex = 10 Then
        ActiveDocument.Variables("TechUsrtva").Value = "баллонов"
        ActiveDocument.Variables("TechUsrtvo").Value = "баллоны"
        ActiveDocument.Variables("TechUsrtvoB").Value = UCase(Left(ActiveDocument.Variables("TechUsrtvo").Value, 1)) & LCase(Mid(ActiveDocument.Variables("TechUsrtvo").Value, 2)) 'Делаем первую букву большой
End If
If UF1.ComboBoxTechUsrtvo.ListIndex = 11 Then
        ActiveDocument.Variables("TechUsrtva").Value = "бака"
        ActiveDocument.Variables("TechUsrtvo").Value = "бак"
        ActiveDocument.Variables("TechUsrtvoB").Value = UCase(Left(ActiveDocument.Variables("TechUsrtvo").Value, 1)) & LCase(Mid(ActiveDocument.Variables("TechUsrtvo").Value, 2)) 'Делаем первую букву большой
End If
If UF1.ComboBoxTechUsrtvo.ListIndex = 12 Then
        ActiveDocument.Variables("TechUsrtva").Value = "резервуара"
        ActiveDocument.Variables("TechUsrtvo").Value = "резервуар"
        ActiveDocument.Variables("TechUsrtvoB").Value = UCase(Left(ActiveDocument.Variables("TechUsrtvo").Value, 1)) & LCase(Mid(ActiveDocument.Variables("TechUsrtvo").Value, 2)) 'Делаем первую букву большой
End If
If UF1.ComboBoxTechUsrtvo.ListIndex = 13 Then
        ActiveDocument.Variables("TechUsrtva").Value = "технологического трубопровода"
        ActiveDocument.Variables("TechUsrtvo").Value = "технологический трубопровод"
        ActiveDocument.Variables("TechUsrtvoB").Value = UCase(Left(ActiveDocument.Variables("TechUsrtvo").Value, 1)) & LCase(Mid(ActiveDocument.Variables("TechUsrtvo").Value, 2)) 'Делаем первую букву большой
End If
If UF1.ComboBoxTechUsrtvo.ListIndex = 14 Then
        ActiveDocument.Variables("TechUsrtva").Value = "теплообменника"
        ActiveDocument.Variables("TechUsrtvo").Value = "теплообменник"
        ActiveDocument.Variables("TechUsrtvoB").Value = UCase(Left(ActiveDocument.Variables("TechUsrtvo").Value, 1)) & LCase(Mid(ActiveDocument.Variables("TechUsrtvo").Value, 2)) 'Делаем первую букву большой
End If
If UF1.ComboBoxTechUsrtvo.ListIndex = 15 Then
        ActiveDocument.Variables("TechUsrtva").Value = "насоса"
        ActiveDocument.Variables("TechUsrtvo").Value = "насос"
        ActiveDocument.Variables("TechUsrtvoB").Value = UCase(Left(ActiveDocument.Variables("TechUsrtvo").Value, 1)) & LCase(Mid(ActiveDocument.Variables("TechUsrtvo").Value, 2)) 'Делаем первую букву большой
End If
End Sub

Private Sub ComboBoxTechUsrtvo_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    BaseUstr(BaseUstrIndex, 13) = UF1.ComboBoxTechUsrtvo.Value
End Sub

Private Sub ComboBoxTipUstroistva_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    BaseUstr(BaseUstrIndex, 12) = UF1.ComboBoxTipUstroistva.Value
End Sub

Private Sub ConfigSave_Click()
'    ReDim artmp(1 To UBound(BaseRemont, 2))
'    For i = 1 To UBound(BaseRemont, 2)
'        artmp(i) = BaseRemont(UstrIndx, i)
'    Next i
'    DateBase.Workbooks("data_base.xls").Worksheets("Remont").Range("B" & UBound(BaseRemont, 1) + 2 & ":R" & UBound(BaseRemont, 1) + 2).Value = artmp 'сохранить одну строку
If UF1.Predpriyatie.ListIndex = -1 Then
    DateBase.Workbooks("data_base.xls").Worksheets("Predpriyatiya").Range("B2:I" & UBound(BasePredp, 1) + 1).Value = BasePredp
Else
    DateBase.Workbooks("data_base.xls").Worksheets("Predpriyatiya").Range("B2:I" & UBound(BasePredp, 1)).Value = BasePredp
End If
If UF1.RegNOPO.ListIndex = -1 Then
    DateBase.Workbooks("data_base.xls").Worksheets("OPO").Range("B2:D" & UBound(BaseOPO, 1) + 1).Value = BaseOPO
Else
    DateBase.Workbooks("data_base.xls").Worksheets("OPO").Range("B2:D" & UBound(BaseOPO, 1)).Value = BaseOPO
End If
If UF1.RegN.ListIndex = -1 Then
    If UF1.RegN.Value <> "" Then
        BaseUstr(BaseUstrIndex, 1) = UF1.RegNOPO.Value
        If UF1.OptionKotel.Value = True Then BaseUstr(BaseUstrIndex, 2) = "Котел"
        If UF1.OptionSosud.Value = True Then BaseUstr(BaseUstrIndex, 2) = "Сосуд"
        If UF1.OptionTruboprovod.Value = True Then BaseUstr(BaseUstrIndex, 2) = "Трубопровод"
        If UF1.OptionSoorugenie.Value = True Then BaseUstr(BaseUstrIndex, 2) = "Сооружение"
        If UF1.OptionOstalnoe.Value = True Then BaseUstr(BaseUstrIndex, 2) = "Остальное"
        BaseElements(BaseUstrIndex, 1) = UF1.ZavN.Value
        BaseElements(BaseUstrIndex, 2) = UF1.RegN.Value
        BaseElements(BaseUstrIndex, 3) = "True"
        BaseEPB(BaseUstrIndex, 1) = UF1.ZavN.Value
        BaseEPB(BaseUstrIndex, 2) = UF1.RegN.Value
        BaseRemont(BaseUstrIndex, 1) = UF1.ZavN.Value
        BaseRemont(BaseUstrIndex, 2) = UF1.RegN.Value
    End If
    DateBase.Workbooks("data_base.xls").Worksheets("Ustroystva").Range("B2:AM" & UBound(BaseUstr, 1) + 1).Value = BaseUstr
    DateBase.Workbooks("data_base.xls").Worksheets("Elements").Range("B2:FI" & UBound(BaseUstr, 1) + 1).Value = BaseElements
    DateBase.Workbooks("data_base.xls").Worksheets("EPB").Range("B2:R" & UBound(BaseUstr, 1) + 1).Value = BaseEPB
    DateBase.Workbooks("data_base.xls").Worksheets("Remont").Range("B2:R" & UBound(BaseUstr, 1) + 1).Value = BaseRemont
Else
    DateBase.Workbooks("data_base.xls").Worksheets("Ustroystva").Range("B2:AM" & UBound(BaseUstr, 1)).Value = BaseUstr
    DateBase.Workbooks("data_base.xls").Worksheets("Elements").Range("B2:FI" & UBound(BaseUstr, 1)).Value = BaseElements
    DateBase.Workbooks("data_base.xls").Worksheets("EPB").Range("B2:R" & UBound(BaseUstr, 1)).Value = BaseEPB
    DateBase.Workbooks("data_base.xls").Worksheets("Remont").Range("B2:R" & UBound(BaseUstr, 1)).Value = BaseRemont
End If
'    Call Save_File("Page1", UF1.OpenFile.TabIndex - 1, "")'
End Sub

Private Sub DataAktPoRez_Change()
If IsDate(UF1.DataAktPoRez.Value) Then
    ActiveDocument.Variables("DataAktPoRez").Value = Trim(UF1.DataAktPoRez.Value)
    ActiveDocument.Variables("DateAktPoRez").Value = FormDat(Trim(UF1.DataAktPoRez.Value))
End If
End Sub

Private Sub DataIzg_Change()
    BaseUstr(BaseUstrIndex, 8) = UF1.DataIzg.Value
    If IsNumeric(Right(UF1.DataIzg.Value, 4)) Then
        ActiveDocument.Variables("DataIzg").Value = Trim(UF1.DataIzg.Value) & " г."
    Else
        ActiveDocument.Variables("DataIzg").Value = Trim(UF1.DataIzg.Value)
    End If
    If UF1.DataIzg.Value = "-" Or UF1.DataIzg.Value = "" Then ActiveDocument.Variables("DataIzg").Value = Strings.ChrW(31)
End Sub

Private Sub DataVvoda_Change()
    BaseUstr(BaseUstrIndex, 9) = UF1.DataVvoda.Value
    If IsNumeric(Val(Right(UF1.DataVvoda.Value, 4))) Then
        UF1.DataVvoda.Value = Trim(UF1.DataVvoda.Value)
        ActiveDocument.Variables("DataVvoda").Value = UF1.DataVvoda.Value & " г."
        ActiveDocument.Variables("SrokSlugb").Value = Year(Date) - Val(Right(UF1.DataVvoda.Value, 4))
        UF1.Inform.Caption = "Срок службы: " & ActiveDocument.Variables("SrokSlugb").Value & Strings.Chr(13) & "Скорость коррозии: " & ActiveDocument.Variables("oSkorKorroz").Value & " мм/год" & Strings.Chr(13) & "Скорость коррозии: " & ActiveDocument.Variables("dSkorKorroz").Value & " мм/год" & Strings.Chr(13)
    End If
End Sub

Private Sub ddiam_Change()
UF1.ddiam.Value = Trim(UF1.ddiam.Value)
ActiveDocument.Variables("ddiam").Value = UF1.ddiam.Value
End Sub

Private Sub Dogovor_Change()
    ActiveDocument.Variables("Dogovor").Value = Trim(UF1.Dogovor.Value)
    BasePredp(BasePredpIndex, 7) = UF1.Dogovor.Value
End Sub

Private Sub DogovorData_Change()
    ActiveDocument.Variables("DogovorData").Value = Trim(UF1.DogovorData.Value)
    BasePredp(BasePredpIndex, 8) = UF1.DogovorData.Value
End Sub

Private Sub DopuskNaprd_Change()
    UF1.DopuskNaprd.Value = Replace(UF1.DopuskNaprd.Value, ".", ",")
    ActiveDocument.Variables("DopuskNaprd").Value = UF1.DopuskNapro.Value
End Sub

Private Sub DopuskNaprd_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Call RaschOstRes
End Sub

Private Sub DopuskNapro_Change()
    UF1.DopuskNapro.Value = Replace(UF1.DopuskNapro.Value, ".", ",")
    ActiveDocument.Variables("DopuskNapro").Value = UF1.DopuskNapro.Value
End Sub

Private Sub DopuskNapro_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Call RaschOstRes
End Sub

Private Sub dtolsh_Change()
    UF1.dtolsh.Value = Trim(UF1.dtolsh.Value)
    UF1.dtolsh.Value = Replace(UF1.dtolsh.Value, ".", ",")
    ActiveDocument.Variables("dtolsh").Value = UF1.dtolsh.Value
End Sub

Private Sub dtolshfakt_Change()
    UF1.dtolshfakt.Value = Trim(UF1.dtolshfakt.Value)
    UF1.dtolshfakt.Value = Replace(UF1.dtolshfakt.Value, ".", ",")
    ActiveDocument.Variables("dtolshfakt").Value = UF1.dtolshfakt.Value
    Call RaschOstRes
'    If UF1.dtolshfakt.Value <> "" Then ActiveDocument.Variables("dSkorKorroz").Value = Format((UF1.dtolsh.Value - UF1.dtolshfakt.Value) / Val(ActiveDocument.Variables("SrokSlugb").Value), "#0.0##")
'    UF1.Inform.Caption = "Срок службы: " & ActiveDocument.Variables("SrokSlugb").Value & Strings.Chr(13) & "Скорость коррозии: " & ActiveDocument.Variables("oSkorKorroz").Value & " мм/год" & Strings.Chr(13) & "Скорость коррозии: " & ActiveDocument.Variables("dSkorKorroz").Value & " мм/год" & Strings.Chr(13)
End Sub

Private Sub ExpertGas_Change()
    If Date > CDate("16.12.2027") Then MsgBox ("Проверь срок действия удостоверения эксперта Э11 ТУ(по газу)")
End Sub

Private Sub ExpertHim_Change()
    If UF1.ExpertORPD.Value = True And UF1.ExpertHim.Value = True Then
        ActiveDocument.Variables("UdostExp").Value = "АЭ.20.01562.003, область аттестации Э12 ТУ, категория эксперта 1, срок действия до 01.03.2026 г." 'Удостоверения эксперта
        ActiveDocument.Variables("UdostExpHim").Value = "; АЭ.21.01562.001, область аттестации Э7 ТУ, категория эксперта 1, срок действия до 16.04.2026 г."
    Else
        If UF1.ExpertHim.Value = True Then ActiveDocument.Variables("UdostExpHim").Value = "АЭ.21.01562.001, область аттестации Э7 ТУ, категория эксперта 1, срок действия до 16.04.2026 г."
        If UF1.ExpertORPD.Value = True Then ActiveDocument.Variables("UdostExp").Value = "АЭ.20.01562.003, область аттестации Э12 ТУ, категория эксперта 1, срок действия до 01.03.2026 г."
    End If
    If UF1.ExpertHim.Value = False Then ActiveDocument.Variables("UdostExpHim").Value = Strings.ChrW(31)
    If Date > CDate("16.04.2026") Then MsgBox ("Проверь срок действия удостоверения эксперта Э7 ТУ")
End Sub

Private Sub ExpertORPD_Change()
    If UF1.ExpertORPD.Value = True And UF1.ExpertHim.Value = True Then
        ActiveDocument.Variables("UdostExp").Value = "АЭ.20.01562.003, область аттестации Э12 ТУ, категория эксперта 1, срок действия до 01.03.2026 г." 'Удостоверения эксперта
        ActiveDocument.Variables("UdostExpHim").Value = "; АЭ.21.01562.001, область аттестации Э7 ТУ, категория эксперта 1, срок действия до 16.04.2026 г."
    Else
        If UF1.ExpertHim.Value = True Then ActiveDocument.Variables("UdostExpHim").Value = "АЭ.21.01562.001, область аттестации Э7 ТУ, категория эксперта 1, срок действия до 16.04.2026 г."
        If UF1.ExpertORPD.Value = True Then ActiveDocument.Variables("UdostExp").Value = "АЭ.20.01562.003, область аттестации Э12 ТУ, категория эксперта 1, срок действия до 01.03.2026 г."
    End If
    If UF1.ExpertORPD.Value = False Then ActiveDocument.Variables("UdostExp").Value = Strings.ChrW(31)
    If Date > CDate("01.03.2026") Then MsgBox ("Проверь срок действия удостоверения эксперта Э12 ТУ")
End Sub

Private Sub ExpertSNN_Change()
    If Date > CDate("04.12.2025") Then MsgBox ("Проверь срок действия удостоверения эксперта по складам НН")
End Sub

Private Sub ExpertZS_Change()
    If Date > CDate("04.12.2025") Then MsgBox ("Проверь срок действия удостоверения экспертов по ЗС")
End Sub

Private Sub FlanzSoed_Change()
    BaseUstr(BaseUstrIndex, 34) = UF1.FlanzSoed.Value
    ActiveDocument.Variables("FlanzSoed").Value = Trim(UF1.FlanzSoed.Value)
End Sub

Private Sub FormSobstv_Change()
    ActiveDocument.Variables("FormSobstv").Value = UF1.FormSobstv.Value
    BasePredp(BasePredpIndex, 3) = UF1.FormSobstv.Value
End Sub

Private Sub IspitatPRub_Change()
    BaseUstr(BaseUstrIndex, 33) = UF1.IspitatPRub.Value
End Sub

Private Sub KartOvalnD_Change()
If IsDate(UF1.KartOvalnD.Value) Then
    ActiveDocument.Variables("KartOvalnD").Value = UF1.KartOvalnD.Value
    ActiveDocument.Variables("KartOvalnData").Value = FormDat(UF1.KartOvalnD.Value)
End If
End Sub

Private Sub KontrGib_Change()
If UF1.KontrGib.Enabled = True Then
    UF1.AktGI.Value = Val(UF1.KontrGib.Value) + 1
    ActiveDocument.Variables("KontrGib").Value = UF1.KontrGib.Value
Else
    UF1.AktGI.Value = Val(UF1.KontrGib.Value)
End If
End Sub

Private Sub KontrGibCh_Change()
If UF1.KontrGibCh.Value = True Then
    UF1.KontrGib.Enabled = True
    UF1.KontrGibD.Enabled = True
    UF1.KontrGibCh.Caption = "Есть"
Else
    UF1.KontrGib.Enabled = False
    UF1.KontrGibD.Enabled = False
    UF1.KontrGibCh.Caption = "Нет"
End If
End Sub

Private Sub KontrGibD_Change()
If IsDate(UF1.KontrGibD.Value) Then
    ActiveDocument.Variables("KontrGibD").Value = UF1.KontrGibD.Value
    ActiveDocument.Variables("KontrGibData").Value = FormDat(UF1.KontrGibD.Value)
End If
End Sub


Private Sub Koof_fid_Change()
    UF1.Koof_fid.Value = Replace(UF1.Koof_fid.Value, ".", ",")
    ActiveDocument.Variables("Koof_fid").Value = UF1.Koof_fid.Value
End Sub

Private Sub Koof_fid_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Call RaschOstRes
End Sub

Private Sub Koof_fio_Change()
    UF1.Koof_fio.Value = Replace(UF1.Koof_fio.Value, ".", ",")
    ActiveDocument.Variables("Koof_fio").Value = UF1.Koof_fio.Value
End Sub

Private Sub Koof_fio_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Call RaschOstRes
End Sub

Private Sub NaNLet_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not IsNumeric(UF1.NaNLet.Value) Then MsgBox ("Нужно ввести число")
End Sub

Private Sub NaznTehUstr_Change()
    BaseUstr(BaseUstrIndex, 14) = UF1.NaznTehUstr.Value
    ActiveDocument.Variables("NaznTehUstr").Value = Trim(UF1.NaznTehUstr.Value)
End Sub

Private Sub NazvOPO_Change()
    If UF1.NazvOPO.ListIndex <> -1 And UF1.NazvOPO.ListCount = UF1.RegNOPO.ListCount Then UF1.RegNOPO.ListIndex = UF1.NazvOPO.ListIndex
    ActiveDocument.Variables("NazvOPO").Value = Trim(UF1.NazvOPO.Value)
    If Trim(UF1.NazvOPO.Value) = "" Then ActiveDocument.Variables("NazvOPO").Value = Strings.ChrW(31)
    If UF1.NazvOPO.Value Like "*[Гг]аз*" Then UF1.ExpertGas.Value = True
End Sub

Private Sub NazvOPO_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    BaseOPO(BaseOPOIndex, 2) = Trim(UF1.NazvOPO.Value)
End Sub

Private Sub NazvTehUstr_Change()
    BaseUstr(BaseUstrIndex, 11) = UF1.NazvTehUstr.Value
    ActiveDocument.Variables("NazvTehUstr").Value = Trim(UF1.NazvTehUstr.Value)
End Sub

Private Sub odiam_Change()
UF1.odiam.Value = Trim(UF1.odiam.Value)
UF1.ddiam.Value = UF1.odiam.Value
ActiveDocument.Variables("odiam").Value = UF1.odiam.Value
End Sub

Private Sub odlina_Change()
UF1.odlina.Value = Trim(UF1.odlina.Value)
ActiveDocument.Variables("odlina").Value = UF1.odlina.Value
End Sub

Private Sub OPOData_Change()
    ActiveDocument.Variables("OPOData").Value = UF1.OPOData.Value
    BasePredp(BasePredpIndex, 5) = UF1.OPOData.Value
End Sub

Private Sub OPOKolStr_Change()
    ActiveDocument.Variables("OPOKolStr").Value = UF1.OPOKolStr.Value
    BasePredp(BasePredpIndex, 6) = UF1.OPOKolStr.Value
End Sub

Private Sub OPONum_Change()
    ActiveDocument.Variables("OPONum").Value = UF1.OPONum.Value
    BasePredp(BasePredpIndex, 4) = UF1.OPONum.Value
End Sub

Private Sub OptionOstalnoe_Click()
If UF1.OptionOstalnoe.Value = True Then
    With UF1.ComboBoxTipUstroistva
        .Clear
        .AddItem "Мазутный насос"
    End With
    UF1.ComboBoxTechUsrtvo.ListIndex = 15
    UF1.KolZicl.Visible = False
    UF1.CBZikl.Visible = False
    UF1.CBZikl.Value = False
    UF1.Label18.Visible = False
    ActiveDocument.Variables("P7RabSredaTopl").Value = "рабочая среда"
    UF1.VikMK.Value = False
    
    Set UstrIndxs = Nothing
    UF1.RegN.Clear
    UF1.ZavN.Clear
    For i = 1 To UBound(BaseUstr, 1) - 1
        If UF1.RegNOPO.Value = BaseUstr(i, 1) And BaseUstr(i, 2) = "Остальное" Then
            UF1.ZavN.AddItem BaseUstr(i, 7)
            UF1.RegN.AddItem BaseUstr(i, 4)
            UstrIndxs.Add i
        End If
    Next i
    If UF1.RegN.ListCount <> 0 Then
        UF1.RegN.ListIndex = 0
    Else
        BaseUstrIndex = UBound(BaseUstr, 1)
    End If
    Call FillOutElements(BaseUstrIndex, BaseElements)
End If
End Sub

Private Sub OptionSoorugenie_Change()
If UF1.OptionSoorugenie.Value = True Then
    With UF1.ComboBoxTipUstroistva
        .Clear
        .AddItem "бак кислоты"
        .AddItem "технологический трубопровод"
        .AddItem "трубопровод кислота"
        .AddItem "резервуар мазутный"
    End With
    UF1.ComboBoxTechUsrtvo.ListIndex = 11
    UF1.KolZicl.Visible = False
    UF1.CBZikl.Visible = False
    UF1.CBZikl.Value = False
    UF1.Label18.Visible = False
    UF1.Label5.Caption = "Монт. орг."
    UF1.ExpertZS.Value = True
    ActiveDocument.Variables("P7RabSredaTopl").Value = "рабочая среда"
    ActiveDocument.Variables("PunktPril8ORPD").Value = ", п.п. 3, 5 Приложения №8"
    UF1.VikMK.Value = False
    
    Set UstrIndxs = Nothing
    UF1.RegN.Clear
    UF1.ZavN.Clear
    For i = 1 To UBound(BaseUstr, 1) - 1
        If UF1.RegNOPO.Value = BaseUstr(i, 1) And BaseUstr(i, 2) = "Сооружение" Then
            UF1.ZavN.AddItem BaseUstr(i, 7)
            UF1.RegN.AddItem BaseUstr(i, 4)
            UstrIndxs.Add i
        End If
    Next i
    If UF1.RegN.ListCount <> 0 Then
        UF1.RegN.ListIndex = 0
    Else
        BaseUstrIndex = UBound(BaseUstr, 1)
    End If
    Call FillOutElements(BaseUstrIndex, BaseElements)
End If
End Sub

Private Sub otolshfakt_Change()
    UF1.otolshfakt.Value = Trim(UF1.otolshfakt.Value)
    UF1.otolshfakt.Value = Replace(UF1.otolshfakt.Value, ".", ",")
    ActiveDocument.Variables("otolshfakt").Value = UF1.otolshfakt.Value
'    UF1.Inform.Caption = "Срок службы: " & ActiveDocument.Variables("SrokSlugb").Value & Strings.Chr(13) & "Скорость коррозии: " & ActiveDocument.Variables("oSkorKorroz").Value & " мм/год" & Strings.Chr(13) & "Скорость коррозии: " & ActiveDocument.Variables("dSkorKorroz").Value & " мм/год" & Strings.Chr(13)
End Sub

Private Sub otolshfakt_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Call RaschOstRes
End Sub

Private Sub PasportKolStr_Change()
    BaseUstr(BaseUstrIndex, 37) = UF1.PasportKolStr.Value
    If UF1.PasportKolStr.Value = "" Then
        ActiveDocument.Variables("PasportKolStr").Value = Trim(UF1.PasportKolStr.Value)
    Else
        ActiveDocument.Variables("PasportKolStr").Value = " - " & Trim(UF1.PasportKolStr.Value) & " л."
    End If
End Sub

Private Sub poleRegNum_Change()
    If BaseUstrIndex = UBound(BaseUstr, 1) Then BaseUstr(BaseUstrIndex, 3) = UF1.poleRegNum.Value
    If UF1.poleRegNum.Value = "-" Or UF1.poleRegNum.Value = "" Then
        ActiveDocument.Variables("No").Value = Strings.ChrW(31)
    Else
        ActiveDocument.Variables("No").Value = UF1.poleRegNum.Value & "№"
    End If
End Sub

Private Sub Predpriyatie_Change()
    ActiveDocument.Variables("Predpriyatie").Value = Trim(UF1.Predpriyatie.Value)
    Dim indks As Long
    Dim sText As String
    
    sText = Trim(UF1.Predpriyatie.Value)
    indks = Predpriyatie.ListIndex
    TT = UF1.Predpriyatie.List
    If Len(sText) <> 0 And indks = -1 Then
        UF1.Predpriyatie.Clear
        For i = 1 To UBound(BasePredp) - 1
            If BasePredp(i, 1) Like "*" & sText & "*" Or BasePredp(i, 2) Like "*" & sText & "*" Then UF1.Predpriyatie.AddItem BasePredp(i, 1)
        Next i
'    If indks = -1 Then cmbSource.Filter = "names Like '*" & sText & "*'"
    ElseIf indks = -1 Then
        UF1.Predpriyatie.Clear
        For i = 1 To UBound(BasePredp, 1) - 1
            Me.Predpriyatie.AddItem BasePredp(i, 1)
        Next i
    Else
        For i = 1 To UBound(BasePredp, 1)
            If BasePredp(i, 1) = Trim(UF1.Predpriyatie.Value) Then
                BasePredpIndex = i
                UF1.PredpriyatieKrNaimen.Value = BasePredp(BasePredpIndex, 2)
                UF1.FormSobstv.Value = BasePredp(BasePredpIndex, 3)
                UF1.OPONum.Value = BasePredp(BasePredpIndex, 4)
                UF1.OPOData.Value = BasePredp(BasePredpIndex, 5)
                UF1.OPOKolStr.Value = BasePredp(BasePredpIndex, 6)
                UF1.Dogovor.Value = BasePredp(BasePredpIndex, 7)
                UF1.DogovorData.Value = BasePredp(BasePredpIndex, 8)
            End If
        Next i
    End If
'
'    If cmbSource.RecordCount = 0 Then
'        Predpriyatie.List = Array("[не найдено соответствия]")
'        Exit Sub
'    End If
'
'    cmbSource.MoveFirst
'    Predpriyatie.Column = cmbSource.GetRows
    Predpriyatie.DropDown
End Sub

Private Sub Predpriyatie_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    UF1.RegNOPO.Clear
    UF1.ClassOpasOPO.Clear
    UF1.NazvOPO.Clear
If UF1.Predpriyatie.ListIndex = -1 Then
    BasePredpIndex = UBound(BasePredp, 1)
    BasePredp(BasePredpIndex, 1) = UF1.Predpriyatie.Value
Else
    For i = 1 To UBound(BasePredp, 1)
        If BasePredp(i, 1) = Trim(UF1.Predpriyatie.Value) Then
            BasePredpIndex = i
            For n = 1 To UBound(BaseOPO, 1) - 1
                If BaseOPO(n, 1) Like BasePredp(i, 4) & "*" Then
                    UF1.NazvOPO.AddItem BaseOPO(n, 2)
                    UF1.ClassOpasOPO.AddItem BaseOPO(n, 3)
                    UF1.RegNOPO.AddItem BaseOPO(n, 1)
                End If
            Next n
            Exit For
        End If
    Next i
    If UF1.RegNOPO.ListCount <> 0 Then UF1.RegNOPO.ListIndex = 0
End If
    UF1.PredpriyatieKrNaimen.Value = BasePredp(BasePredpIndex, 2)
    UF1.FormSobstv.Value = BasePredp(BasePredpIndex, 3)
    UF1.OPONum.Value = BasePredp(BasePredpIndex, 4)
    UF1.OPOData.Value = BasePredp(BasePredpIndex, 5)
    UF1.OPOKolStr.Value = BasePredp(BasePredpIndex, 6)
    UF1.Dogovor.Value = BasePredp(BasePredpIndex, 7)
    UF1.DogovorData.Value = BasePredp(BasePredpIndex, 8)
End Sub

Private Sub PredpriyatieKrNaimen_Change()
    ActiveDocument.Variables("PredpriyatieKrNaimen").Value = UF1.PredpriyatieKrNaimen.Value
    BasePredp(BasePredpIndex, 2) = UF1.PredpriyatieKrNaimen.Value
End Sub

Private Sub PribNaKorrd_Change()
    UF1.PribNaKorrd.Value = Replace(UF1.PribNaKorrd.Value, ".", ",")
    ActiveDocument.Variables("PribNaKorrd").Value = UF1.PribNaKorrd.Value
End Sub

Private Sub PribNaKorrd_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Call RaschOstRes
End Sub

Private Sub PribNaKorro_Change()
    UF1.PribNaKorro.Value = Replace(UF1.PribNaKorro.Value, ".", ",")
    ActiveDocument.Variables("PribNaKorro").Value = UF1.PribNaKorro.Value
End Sub

Private Sub PribNaKorro_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Call RaschOstRes
End Sub

Private Sub PrimSvMat_Change()
    BaseUstr(BaseUstrIndex, 35) = UF1.PrimSvMat.Value
    ActiveDocument.Variables("PrimSvMat").Value = Trim(UF1.PrimSvMat.Value)
End Sub

Private Sub Progib_Change()
ActiveDocument.Variables("Progib").Value = UF1.Progib.Value
If UF1.Progib.Enabled = True Then
    UF1.KontrGib.Value = Val(UF1.Progib.Value) + 1
Else
    UF1.KontrGib.Value = Val(UF1.Progib.Value)
End If

End Sub

Private Sub ProgibCb_Change()
If UF1.ProgibCb.Value = True Then
    UF1.Progib.Enabled = True
    UF1.ProgibD.Enabled = True
    UF1.ProgibCb.Caption = "Есть"
Else
    UF1.Progib.Enabled = False
    UF1.ProgibD.Enabled = False
    UF1.ProgibCb.Caption = "Нет"
End If
End Sub

Private Sub ProgibD_Change()
If IsDate(UF1.ProgibD.Value) Then
    ActiveDocument.Variables("ProgibD").Value = UF1.ProgibD.Value
    ActiveDocument.Variables("ProgibData").Value = FormDat(UF1.ProgibD.Value)
End If
End Sub

Private Sub ProtkTVD_Change()
If IsDate(UF1.ProtkTVD.Value) Then
    ActiveDocument.Variables("ProtkTVD").Value = UF1.ProtkTVD.Value
    ActiveDocument.Variables("ProtkTVData").Value = FormDat(UF1.ProtkTVD.Value)
End If
End Sub

Private Sub ProtokolVD_Change()
    ActiveDocument.Variables("ProtokolVD").Value = UF1.ProtokolVD.Value
End Sub

Private Sub ProtokolVDD_Change()
    ActiveDocument.Variables("ProtokolVDD").Value = UF1.ProtokolVDD.Value
    If IsDate(UF1.ProtokolVDD.Value) Then ActiveDocument.Variables("ProtokolVDData").Value = FormDat(UF1.ProtokolVDD.Value)
End Sub

Private Sub ProtokTolchD_Change()
If IsDate(UF1.ProtokTolchD.Value) Then
    ActiveDocument.Variables("ProtokTolchD").Value = UF1.ProtokTolchD.Value
    ActiveDocument.Variables("ProtokTolchData").Value = FormDat(UF1.ProtokTolchD.Value)
End If
End Sub

Private Sub AktGID_Change()
If IsDate(UF1.AktGID.Value) Then
    ActiveDocument.Variables("DoNgoda").Value = Format(DateAdd("yyyy", Val(UF1.NaNLet.Value), UF1.AktGID.Value), "dd.mm.yyyy")
    ActiveDocument.Variables("DoNgodaOsv").Value = Format(DateAdd("yyyy", 3, UF1.AktGID.Value), "dd.mm.yyyy")
    ActiveDocument.Variables("AktGID").Value = UF1.AktGID.Value
    ActiveDocument.Variables("AktGIData").Value = FormDat(UF1.AktGID.Value)
    UF1.AktAED.Value = UF1.AktGID.Value
    UF1.ProtokolVDD.Value = UF1.AktGID.Value
End If
End Sub

Private Sub AktAE_Change()
    ActiveDocument.Variables("AktAE").Value = UF1.AktAE.Value
End Sub

Private Sub AktGI_Change()
If IsNumeric(UF1.AktGI.Value) Then
    ActiveDocument.Variables("AktGI").Value = "№" & UF1.AktGI.Value
Else
    ActiveDocument.Variables("AktGI").Value = UF1.AktGI.Value
End If
If UF1.AktGI.Enabled = True Then
    UF1.AktAE.Value = Val(UF1.AktGI.Value) + 1
    UF1.ProtokolVD.Value = Val(UF1.AktGI.Value) + 1
Else
    UF1.AktAE.Value = Val(UF1.AktGI.Value)
    UF1.ProtokolVD.Value = Val(UF1.AktGI.Value)
End If
End Sub

Private Sub OpenFile_Click()
Dim fName As String
Dim lRetVal As Long
Dim objWrd As Object
Dim objDoc As Object
    With Application.Dialogs(wdDialogFileOpen)
        lRetVal = .Display
        fName = .Name
        fName = Options.DefaultFilePath(wdCurrentFolderPath) & "\" & Replace(fName, """", "")
    End With
    If lRetVal <> -1 Then Exit Sub
Set objWrd = CreateObject("Word.Application")
Set objDoc = objWrd.Documents.Open(fName)
objWrd.Visible = True
End Sub

Private Sub dp1_Change()
ActiveDocument.Variables("pribor1").Value = UF1.pribor1.Value & UF1.dp1.Value
End Sub
Private Sub dp2_Change()
ActiveDocument.Variables("pribor2").Value = UF1.pribor2.Value & UF1.dp2.Value
End Sub
Private Sub dp3_Change()
ActiveDocument.Variables("pribor3").Value = UF1.pribor3.Value & UF1.dp3.Value
End Sub
Private Sub dp4_Change()
ActiveDocument.Variables("pribor4").Value = UF1.pribor4.Value & UF1.dp4.Value
End Sub
Private Sub dp5_Change()
ActiveDocument.Variables("pribor5").Value = UF1.pribor5.Value & UF1.dp5.Value
End Sub
Private Sub dp6_Change()
ActiveDocument.Variables("pribor6").Value = UF1.pribor6.Value & UF1.dp6.Value
End Sub
Private Sub dp7_Change()
ActiveDocument.Variables("pribor7").Value = UF1.pribor7.Value & UF1.dp7.Value
End Sub
Private Sub dp8_Change()
ActiveDocument.Variables("pribor8").Value = UF1.pribor8.Value & UF1.dp8.Value
End Sub
Private Sub dp9_Change()
ActiveDocument.Variables("pribor9").Value = UF1.pribor9.Value & UF1.dp9.Value
End Sub
Private Sub dp10_Change()
ActiveDocument.Variables("pribor10").Value = UF1.pribor10.Value & UF1.dp10.Value
End Sub
Private Sub dp11_Change()
ActiveDocument.Variables("pribor11").Value = UF1.pribor11.Value & UF1.dp11.Value
End Sub
Private Sub dp12_Change()
ActiveDocument.Variables("pribor12").Value = UF1.pribor12.Value & UF1.dp12.Value
End Sub
Private Sub dp13_Change()
ActiveDocument.Variables("pribor13").Value = UF1.pribor13.Value & UF1.dp13.Value
End Sub
Private Sub dp14_Change()
ActiveDocument.Variables("pribor14").Value = UF1.pribor14.Value & UF1.dp14.Value
End Sub
Private Sub dp15_Change()
ActiveDocument.Variables("pribor15").Value = UF1.pribor15.Value & UF1.dp15.Value
End Sub
Private Sub dp16_Change()
ActiveDocument.Variables("pribor16").Value = UF1.pribor16.Value & UF1.dp16.Value
End Sub
Private Sub dp17_Change()
ActiveDocument.Variables("pribor17").Value = UF1.pribor17.Value & UF1.dp17.Value
End Sub

Private Sub IspitatP_Change()
    UF1.IspitatP.Value = Trim(UF1.IspitatP.Value)
    UF1.IspitatP.Value = Replace(UF1.IspitatP.Value, ".", ",")
    BaseUstr(BaseUstrIndex, 32) = UF1.IspitatP.Value
    If IsNumeric(UF1.RazreshaemoeP.Value) And IsNumeric(UF1.IspitatP.Value) Then
        ActiveDocument.Variables("IspitatP").Value = Format(UF1.IspitatP.Value, "###0.0#####") & " кгс/см" & Strings.ChrW(178)
        ActiveDocument.Variables("IspitatPMP").Value = Format(CDbl(UF1.IspitatP.Value) / 10, "###0.0#####") & " МПа"
    Else
        ActiveDocument.Variables("IspitatP").Value = UF1.IspitatP.Value
        ActiveDocument.Variables("IspitatPMP").Value = Strings.ChrW(31)
    End If
End Sub

Private Sub KartOvaln_Change()
ActiveDocument.Variables("KartOvaln").Value = UF1.KartOvaln.Value
If UF1.KartOvaln.Enabled = True Then
    UF1.Progib.Value = Val(UF1.KartOvaln.Value) + 1
Else
    UF1.Progib.Value = Val(UF1.KartOvaln.Value)
End If
End Sub

Private Sub NaNLet_Change()
UF1.NaNLet.Value = Trim(UF1.NaNLet.Value)
If IsNumeric(UF1.NaNLet.Value) Then 'Если в поле число то добавляем год, года или лет
    If (Val(UF1.NaNLet.Value) = 1) Then ActiveDocument.Variables("NaNLet").Value = UF1.NaNLet.Value & " год"
    If (Val(UF1.NaNLet.Value) > 1 And Val(UF1.NaNLet.Value) < 5) Then ActiveDocument.Variables("NaNLet").Value = UF1.NaNLet & " года"
    If (Val(UF1.NaNLet.Value) > 4) Then ActiveDocument.Variables("NaNLet").Value = UF1.NaNLet.Value & " лет"
'    ActiveDocument.Variables("DoNgoda").Value = Format(DateAdd("yyyy", Val(UF1.NaNLet.Value), UF1.AktGID.Value), "dd.mm.yyyy")
Else
    MsgBox ("Нужно ввести число")
End If
End Sub

Private Sub DataAktPoRez_Exit(ByVal Cancel As MSForms.ReturnBoolean)
If Not IsDate(UF1.DataAktPoRez.Value) Then
    MsgBox ("Неправильный формат даты")
End If
End Sub

Private Sub DataRegistracii_Change()
    BaseUstr(BaseUstrIndex, 5) = UF1.DataRegistracii.Value
    If UF1.DataRegistracii.Value = "-" Or UF1.DataRegistracii.Value = "" Then
        ActiveDocument.Variables("DataRegistracii").Value = Strings.ChrW(31)
    Else
        ActiveDocument.Variables("DataRegistracii").Value = UF1.DataRegistracii.Value
        If IsDate(Trim(UF1.DataRegistracii.Value)) Then ActiveDocument.Variables("DataRegistracii").Value = ActiveDocument.Variables("DataRegistracii").Value & " г."
    End If
End Sub

Private Sub DataVvoda_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Val(ActiveDocument.Variables("SrokSlugb").Value) > 100 Or UF1.DataVvoda.Value = "" Then MsgBox ("Неправильный формат даты")
    Call RaschOstRes
End Sub

Private Sub ddlina_Change()
UF1.ddlina.Value = Trim(UF1.ddlina.Value)
ActiveDocument.Variables("ddlina").Value = UF1.ddlina.Value
End Sub

Private Sub dgost_Change()
UF1.dgost.Value = Trim(UF1.dgost.Value)
ActiveDocument.Variables("dgost").Value = UF1.dgost.Value
End Sub

'Private Sub A13_Exit(ByVal Cancel As MSForms.ReturnBoolean)
'uf1.A13 = Replace(uf1.A13, ".", ",")
'End Sub
'
Private Sub AktVIK_Change()
ActiveDocument.Variables("AktVIK").Value = UF1.AktVIK.Value
UF1.AktVIKMK.Value = Val(UF1.AktVIK.Value) + 1
End Sub

Private Sub Button1_Click()
Application.ScreenUpdating = False

If UF1.OptionKotel.Value = True Then ' Если составляем заключение для котлов
    Call DeleteBookmarks("Rezervuar") ' удаляем часть про резервуары и сооружения
    If ActiveDocument.Bookmarks.Exists("P7p1mat") = True Then ActiveDocument.Bookmarks("P7p1mat").Range.Delete ' пункт 7.1. материалы
    If ActiveDocument.Bookmarks.Exists("VikRezTruboprov") = True Then ActiveDocument.Bookmarks("VikRezTruboprov").Range.Delete ' ВИК результаты осмотра трубопровода
    If ActiveDocument.Bookmarks.Exists("TipSvS") = True Then ActiveDocument.Bookmarks("TipSvS").Range.Delete ' пункт в УЗК - тип сварного соединения
    ActiveDocument.Variables.Item("PassatT").Value = Strings.ChrW(31)
    ActiveDocument.Variables("p12-1pril2").Value = " п. 12.1. Приложения №2, п.п. 2, 3, 4 Приложения №8" & UF1.FNPORPDR.Value
    ActiveDocument.Variables("VIKrd").Value = " п.п. 5.4, 5.5, 5.15"
'    ActiveDocument.Variables("ovalnrd").Value = "При измерениях овальности барабанов котла установлено, что овальность соответствует требованиям п. 5.10." & UF1.SO469.Value
    ActiveDocument.Variables("RabSredaTopl").Value = "Вид топлива"
    ActiveDocument.Variables("RabSredaToplRasch").Value = "Вид топлива"
    ActiveDocument.Tables(2).Cell(Row:=4, Column:=1).Range = "Вид топлива"
    ActiveDocument.Variables("punkt7-3-1").Value = "Установка и оснащение "
End If
If UF1.OptionSosud.Value = True Then ' Если составляем заключение для сосудов
    Call DeleteBookmarks("Rezervuar") ' удаляем часть про резервуары и сооружения
'    If ActiveDocument.Bookmarks.Exists("KotlObor") = True Then ActiveDocument.Bookmarks("KotlObor").Range.Delete ' пункт 7.3. оборудование котла
    If ActiveDocument.Bookmarks.Exists("TipSvS") = True Then ActiveDocument.Bookmarks("TipSvS").Range.Delete ' пункт в УЗК - тип сварного соединения
    If ActiveDocument.Bookmarks.Exists("VikRezTruboprov") = True Then ActiveDocument.Bookmarks("VikRezTruboprov").Range.Delete ' ВИК результаты осмотра трубопровода
    If UF1.Ovalnost.Value = True And UF1.CBSO439.Value = True Then ActiveDocument.Variables("VIKrdSO439").Value = "; п.п. 5.4, 5.6, 5.7, 5.8, 5.10" & UF1.SO439.Value
    Call DeleteBookmarks("MetKonstrKotla") ' Металлоконструкции котла
End If
If UF1.OptionTruboprovod.Value = True Then ' Если составляем заключение для трубопровода
    Call DeleteBookmarks("Rezervuar") ' удаляем часть про резервуары и сооружения
    If ActiveDocument.Bookmarks.Exists("P7p1mat") = True Then ActiveDocument.Bookmarks("P7p1mat").Range.Delete ' пункт 7.1. материалы
'    If ActiveDocument.Bookmarks.Exists("KotlObor") = True Then ActiveDocument.Bookmarks("KotlObor").Range.Delete ' пункт 7.3. оборудование котла
    If ActiveDocument.Bookmarks.Exists("VikRezKontr") = True Then ActiveDocument.Bookmarks("VikRezKontr").Range.Delete ' ВИК результаты осмотра
    ActiveDocument.Variables("punkt7-3-1").Value = "Прокладка и оснащение "
'    ActiveDocument.Variables("punkt7-3").Value = Strings.ChrW(31)
    ActiveDocument.Variables.Item("PassatT").Value = Strings.ChrW(31)
    ActiveDocument.Variables("korpusa").Value = Strings.ChrW(31)
End If
If UF1.OptionOstalnoe.Value = True Then
    Call DeleteBookmarks("Rezervuar") ' удаляем часть про резервуары и сооружения
'    If ActiveDocument.Bookmarks.Exists("KotlObor") = True Then ActiveDocument.Bookmarks("KotlObor").Range.Delete ' пункт 7.3. оборудование котла
    If ActiveDocument.Bookmarks.Exists("VikRezTruboprov") = True Then ActiveDocument.Bookmarks("VikRezTruboprov").Range.Delete ' ВИК результаты осмотра трубопровода
End If
If UF1.CBZikl.Value = True Then
    If Val(UF1.KolZicl.Value) > 1000 Then ActiveDocument.Bookmarks("ziklmen1000").Range.Delete
Else
    Call DeleteBookmarks("zikl")
End If

Call OsnovnPunkt
'Индивидуальные изменения для каждого техустройства
If UF1.ComboBoxTipUstroistva.Value = "Воздухосборник" Then Call Vozduhosbornik
If UF1.ComboBoxTipUstroistva.Value = "НЖУ, ЦЖУ, УДХ" Then Call NGUCGUUDH
If UF1.ComboBoxTipUstroistva.Value = "Баллоны групповой установки" Then Call BallGroUst
If UF1.ComboBoxTipUstroistva.Value = "Баллон" Then Call Ballon
If UF1.ComboBoxTipUstroistva.Value = "Автоцистерна для СУГ" Then Call Avtozisterna
If UF1.ComboBoxTipUstroistva.Value = "Подогреватель" Then Call Podogrevatel
If UF1.ComboBoxTipUstroistva.Value = "Газификатор ГХК" Then Call Gasifikator
If UF1.ComboBoxTipUstroistva.Value = "Сосуд с вакуумом" Then Call VakuumSosud
If UF1.ComboBoxTipUstroistva.Value = "Сосуд под налив" Then Call SosudPodNaliv
If UF1.ComboBoxTipUstroistva.Value = "Бак кислотный" Then Call SosudHOPO
If UF1.ComboBoxTipUstroistva.Value = "водогрейный котел" Then Call VodgKotl
If UF1.ComboBoxTipUstroistva.Value = "экономайзер" Then Call Ekonomayzer
If UF1.ComboBoxTipUstroistva.Value = "электрический котел" Then Call ElektroKotel
If UF1.ComboBoxTipUstroistva.Value = "паровой котел" Then Call ParKotl
If UF1.ComboBoxTipUstroistva.Value = "трубопровод пара" Then Call TruboprovPara
If UF1.ComboBoxTipUstroistva.Value = "технологический трубопровод" Then Call TehnTruboprovod
If UF1.ComboBoxTipUstroistva.Value = "трубопровод кислота" Then Call TruboprovodKislota
If UF1.ComboBoxTipUstroistva.Value = "резервуар мазутный" Then Call RezervuarMazut
If UF1.ComboBoxTipUstroistva.Value = "Мазутный насос" Then Call Nasos

'Эксперты по газу и сооружениям
If UF1.ExpertGas.Value = False Then
Call DeleteBookmarks("ExprtPG")
End If
If UF1.ExpertSNN.Value = False Then
Call DeleteBookmarks("ExprtPSNN")
End If
If UF1.ExpertZS.Value = False Then
Call DeleteBookmarks("ExprtPZ")
End If

'ВИК МК да/нет
If (UF1.VikMK.Value = False) Then
    Call DeleteBookmarks("MetKonstrKotla")
End If
'Толщинометрия да/нет
If (UF1.Tolshin.Value = False) Then
    Call DeleteBookmarks("Tolshin")
End If
'УЗК да/нет
If (UF1.UZK.Value = False) Then
    Call DeleteBookmarks("UZK")
End If
'МПД/ЦД/нет
If (UF1.MPDZD.Value = True) Then
    Call DeleteBookmarks("ZD")
Else
    If (UF1.MPDZD.Value = False) Then
        Call DeleteBookmarks("ZD")
        Call DeleteBookmarks("MPD")
    Else
        Call DeleteBookmarks("MPD")
    End If
End If
'Твердость да/нет
If (UF1.Tverdost.Value = False) Then
    Call DeleteBookmarks("Tverdost")
End If
'Овальность да/нет
If (UF1.Ovalnost.Value = False) Then
    Call DeleteBookmarks("Ovaln")
End If
'Прогиб да/нет
If (UF1.ProgibCb.Value = False) Then
    Call DeleteBookmarks("Progib")
End If
'Контроль гибов да/нет
If (UF1.KontrGibCh.Value = False) Then
    Call DeleteBookmarks("KontrGib")
End If
'Вибродиагностика да/нет
If (UF1.ProtokolVD.Visible = False) Then
    Call DeleteBookmarks("ProtokolVD")
End If
'Подготавливаем вариант гидравлического или пневматического испытания
If (UF1.PnIs.Value = False) Then
    Call DeleteBookmarks("PnevmatIsp")
Else
    Call DeleteBookmarks("GidroIsp")
End If
'Индивидуальные изменения для резервуаров
If UF1.ComboBoxTipUstroistva.Value = "бак кислоты" Then Call BakKislota
Call TckZpt 'Расставление запятых

'Заполняем таблицу элементов
Z = 3
For i = 1 To 20
    If BaseElements(BaseUstrIndex, Z) = "True" Then
        If i > 1 Then ActiveDocument.Tables(1).Rows.Add
        Z = Z + 1
        X = ActiveDocument.Tables(1).Rows.Count
        For n = 1 To 6
            If n = 5 Then Z = Z + 1
            ActiveDocument.Tables(1).Cell(Row:=X, Column:=n + 1).Range = BaseElements(BaseUstrIndex, Z)
            Z = Z + 1
        Next n
    Else
        Z = Z + 8
    End If
Next i

'Заполняем сведения о ремонтах и экспертизах
ActiveDocument.Variables("SvedORemonte").Value = BaseRemont(BaseUstrIndex, 3)
ActiveDocument.Variables("SvedOEPB").Value = BaseEPB(BaseUstrIndex, 3)
For i = 4 To 17
    If BaseRemont(BaseUstrIndex, i) <> "" Then ActiveDocument.Variables("SvedORemonte").Value = ActiveDocument.Variables("SvedORemonte").Value & "." & Strings.Chr(13) & BaseRemont(BaseUstrIndex, i)
    If BaseEPB(BaseUstrIndex, i) <> "" Then ActiveDocument.Variables("SvedOEPB").Value = ActiveDocument.Variables("SvedOEPB").Value & "." & Strings.Chr(13) & BaseEPB(BaseUstrIndex, i)
Next i

'Собираем заголовок для сохранения файла
ActiveDocument.BuiltInDocumentProperties("Title").Value = Replace(UF1.NazvTehUstr.Value & " рег.№" & UF1.RegN.Value & "(" & Year(Date) & ")", "/", "-")
ActiveDocument.BuiltInDocumentProperties("Title").Value = Replace(ActiveDocument.BuiltInDocumentProperties("Title").Value, "\", "-")
'Unload Me
UF1.hide
End Sub

Private Sub dstal_Change()
UF1.dstal.Value = Trim(UF1.dstal.Value)
ActiveDocument.Variables("dstal").Value = UF1.dstal.Value
End Sub

Private Sub KolZicl_Change()
    BaseElements(BaseUstrIndex, 164) = Trim$(UF1.KolZicl.Value)
    ActiveDocument.Variables("KolZicl").Value = Val(Trim(UF1.KolZicl.Value))
End Sub

Private Sub ComboBoxTipUstroistva_Change()
Dim SelectComboBox()
If UF1.OptionSosud.Value = True Then
    If UF1.ComboBoxTipUstroistva.ListIndex = 0 Then
        Call ClearAllF
        UF1.ExpertORPD.Value = True
        UF1.ComboBoxTechUsrtvo.ListIndex = 2
        UF1.ComboBoxRaschet.ListIndex = 0
        ActiveDocument.Variables("TipTehUstr").Value = "сосуд, работающий под давлением"
        If BaseUstrIndex = UBound(BaseUstr, 1) Then UF1.NaznTehUstr.Value = "сосуд для хранения "
        UF1.CBFNPORPD.Value = True
        UF1.CBSO439.Value = True
        UF1.CBGOST34347.Value = True
        Call UstSosAktNK(0, 1, 1, 1, 1, 0, 0, 0, 0)
    End If
    If UF1.ComboBoxTipUstroistva.ListIndex = 1 Then
        Call ClearAllF
        UF1.ExpertORPD.Value = True
        UF1.ComboBoxTechUsrtvo.ListIndex = 4
        UF1.ComboBoxRaschet.ListIndex = 0
        ActiveDocument.Variables("TipTehUstr").Value = "сосуд, работающий под давлением"
        If BaseUstrIndex = UBound(BaseUstr, 1) Then UF1.NaznTehUstr.Value = "сосуд для накопления и хранения сжатого воздуха"
        UF1.CBFNPORPD.Value = True
        UF1.CBSO439.Value = True
        UF1.CBGOST34347.Value = True
        Call UstSosAktNK(0, 1, 1, 1, 1, 0, 0, 0, 0)
    End If
    If UF1.ComboBoxTipUstroistva.ListIndex = 2 Then
        Call ClearAllF
        UF1.ExpertORPD.Value = True
        UF1.ComboBoxTechUsrtvo.ListIndex = 2
        UF1.ComboBoxRaschet.ListIndex = 0
        ActiveDocument.Variables("TipTehUstr").Value = "сосуд, работающий под давлением"
        If BaseUstrIndex = UBound(BaseUstr, 1) Then UF1.NaznTehUstr.Value = "сосуд для хранения СО2"
        UF1.CBFNPORPD.Value = True
        UF1.CBSO439.Value = True
        UF1.CBGOST34347.Value = True
        UF1.CBRD2626012.Value = True
        Call UstSosAktNK(0, 1, 1, 1, 1, 0, 0, 0, 1)
    End If
    If UF1.ComboBoxTipUstroistva.ListIndex = 3 Then
        Call ClearAllF
        SelectComboBox = Array(2, 3, 10, 175, 179, 185, 186, 187, 188, 465, 468, 469, 471, 538, 539, 540, 577, 589)
        Call SetComboBox(SelectComboBox, "CBp")
        UF1.ExpertORPD.Value = True
        UF1.ComboBoxTechUsrtvo.ListIndex = 2
        UF1.ComboBoxRaschet.ListIndex = 4
        ActiveDocument.Variables("TipTehUstr").Value = "баллоны групповой установки"
        If BaseUstrIndex = UBound(BaseUstr, 1) Then UF1.NaznTehUstr.Value = "для подъема давления"
        UF1.CBFNPORPD.Value = True
        UF1.CBSO439.Value = True
        UF1.CBGOST34347.Value = True
        Call UstSosAktNK(0, 1, 0, 1, 1, 0, 0, 0, 0)
    End If
    If UF1.ComboBoxTipUstroistva.ListIndex = 4 Then
        Call ClearAllF
        SelectComboBox = Array(2, 3, 10, 65, 68, 69, 100, 154, 175, 178, 185, 187, 188, 338, 339, 340, 343, 353, 465, 468, 519, 521, 523)
        Call SetComboBox(SelectComboBox, "CBp")
        UF1.ExpertORPD.Value = True
        UF1.ExpertGas.Value = True
        UF1.ComboBoxTechUsrtvo.ListIndex = 2
        UF1.ComboBoxRaschet.ListIndex = 0
        ActiveDocument.Variables("TipTehUstr").Value = "сосуд, работающий под давлением"
        If BaseUstrIndex = UBound(BaseUstr, 1) Then UF1.NaznTehUstr.Value = "для хранения и транспортировки СУГ"
        UF1.CBFNPORPD.Value = True
        UF1.CBFNPSUG.Value = True
        UF1.CBSO439.Value = True
        UF1.CBGOST34347.Value = True
'        UF1.CBRD2626012.Value = True
        Call UstSosAktNK(0, 1, 1, 0, 0, 0, 0, 0, 0)
    End If
    If UF1.ComboBoxTipUstroistva.ListIndex = 5 Then
        Call ClearAllF
        SelectComboBox = Array(2, 3, 10, 68, 69, 100, 175, 178, 185, 186, 187, 188, 338, 339, 340, 341, 343, 353, 465, 468)
        Call SetComboBox(SelectComboBox, "CBp")
        UF1.ExpertORPD.Value = True
        UF1.ComboBoxTechUsrtvo.ListIndex = 6
        UF1.ComboBoxRaschet.ListIndex = 0
        ActiveDocument.Variables("TipTehUstr").Value = "сосуд, работающий под давлением"
        If BaseUstrIndex = UBound(BaseUstr, 1) Then UF1.NaznTehUstr.Value = "подогрев сетевой воды"
        UF1.CBFNPORPD.Value = True
        UF1.CBSO439.Value = True
        UF1.CBGOST34347.Value = True
    End If
    If UF1.ComboBoxTipUstroistva.ListIndex = 6 Then
        Call ClearAllF
        SelectComboBox = Array(2, 3, 10, 65, 68, 69, 100, 154, 175, 190, 338, 339, 340, 341, 343, 353, 465, 468, 471)
        Call SetComboBox(SelectComboBox, "CBp")
        UF1.ExpertORPD.Value = True
        UF1.ComboBoxTechUsrtvo.ListIndex = 5
        UF1.ComboBoxRaschet.ListIndex = 0
        ActiveDocument.Variables("TipTehUstr").Value = "сосуд, работающий под давлением"
        If BaseUstrIndex = UBound(BaseUstr, 1) Then UF1.NaznTehUstr.Value = "хранение сжиженных криопродуктов"
        UF1.CBFNPORPD.Value = True
        UF1.CBVM030104.Value = True
        Call UstSosAktNK(0, 0, 0, 0, 0, 0, 0, 0, 1)
    End If
    If UF1.ComboBoxTipUstroistva.ListIndex = 7 Then
        Call ClearAllF
        SelectComboBox = Array(161, 164, 169, 177, 178, 179)
        Call SetComboBox(SelectComboBox, "CBvb")
        UF1.ExpertHim.Value = True
        UF1.ComboBoxTechUsrtvo.ListIndex = 2
        UF1.ComboBoxRaschet.ListIndex = 0
        ActiveDocument.Variables("TipTehUstr").Value = "сосуд, работающий без давления (под вакуумом)"
        If BaseUstrIndex = UBound(BaseUstr, 1) Then UF1.NaznTehUstr.Value = "вакуумная емкость"
        UF1.CBRUA93.Value = True
        ActiveDocument.Variables("p2-102RUA").Value = " 2.102,"
        ActiveDocument.Variables("p2-111RUA").Value = Strings.ChrW(31)
        UF1.CBFNPOPVB.Value = True
        UF1.CBGOST34347.Value = True
    End If
    If UF1.ComboBoxTipUstroistva.ListIndex = 8 Then
        Call ClearAllF
        SelectComboBox = Array(11, 12, 15, 132, 135, 136, 137, 140)
        Call SetComboBox(SelectComboBox, "CBho")
        SelectComboBox = Array(9, 161, 164, 169, 177, 178, 179)
        Call SetComboBox(SelectComboBox, "CBvb")
        UF1.ExpertHim.Value = True
        UF1.CBZikl.Value = False
        UF1.ComboBoxTechUsrtvo.ListIndex = 2
        UF1.ComboBoxRaschet.ListIndex = 0
        ActiveDocument.Variables("TipTehUstr").Value = "сосуд, работающий без давления (под налив)"
        If BaseUstrIndex = UBound(BaseUstr, 1) Then UF1.NaznTehUstr.Value = "емкость под налив"
        UF1.CBRUA93.Value = True
        ActiveDocument.Variables("p2-102RUA").Value = Strings.ChrW(31)
        ActiveDocument.Variables("p2-111RUA").Value = " 2.111,"
        UF1.CBFNPOPVB.Value = True
        UF1.CBGOST34347.Value = True
    End If
    If UF1.ComboBoxTipUstroistva.ListIndex = 9 Then
        Call ClearAllF
        SelectComboBox = Array(11, 12, 15, 132, 135, 137, 140, 233, 234, 255, 267)
        Call SetComboBox(SelectComboBox, "CBho")
        UF1.ExpertHim.Value = True
        UF1.ComboBoxTechUsrtvo.ListIndex = 11
        UF1.ComboBoxRaschet.ListIndex = 2
        ActiveDocument.Variables("TipTehUstr").Value = "емкость"
        UF1.NaznTehUstr.Value = "емкость под налив"
        ActiveDocument.Variables("p2-102RUA").Value = Strings.ChrW(31)
        ActiveDocument.Variables("p2-111RUA").Value = " 2.111,"
        UF1.CBZikl.Value = False
        UF1.CBFNPHOPO.Value = True
        UF1.CBRUA93.Value = True
    End If
    If UF1.ComboBoxTipUstroistva.ListIndex = 10 Then
        Call ClearAllF
        SelectComboBox = Array(2, 3, 10, 175, 179, 185, 186, 187, 188, 465, 468, 469, 471, 538, 539, 540, 577, 589)
        Call SetComboBox(SelectComboBox, "CBp")
        UF1.ExpertORPD.Value = True
        UF1.ComboBoxTechUsrtvo.ListIndex = 9
        UF1.ComboBoxRaschet.ListIndex = 4
        UF1.CBZikl.Value = False
        UF1.CBFNPORPD.Value = True
        UF1.CBSO439.Value = True
        UF1.CBGOST34347.Value = True
    End If
End If
If UF1.OptionKotel.Value = True Then
    If UF1.ComboBoxTipUstroistva.ListIndex = 0 Then
        Call ClearAllF
        SelectComboBox = Array(2, 3, 10, 38, 39, 43, 45, 47, 49, 50, 100, 154, 175, 177, 185, 186, 187, 188, 257, 258, 260, 267, 271, 465, 468, 469, 471)
        Call SetComboBox(SelectComboBox, "CBp")
        UF1.ExpertORPD.Value = True
        UF1.ComboBoxTechUsrtvo.ListIndex = 0
        UF1.ComboBoxRaschet.ListIndex = 1
        ActiveDocument.Variables("TipTehUstr").Value = "паровой котел"
        If BaseUstrIndex = UBound(BaseUstr, 1) Then UF1.NaznTehUstr.Value = "выработка пара"
        UF1.CBFNPORPD.Value = True
        UF1.CBSO469.Value = True
        Call UstSosAktNK(1, 1, 1, 1, 1, 1, 1, 1, 0)
    End If
    If UF1.ComboBoxTipUstroistva.ListIndex = 1 Then
        Call ClearAllF
        SelectComboBox = Array(2, 3, 10, 38, 39, 45, 46, 47, 50, 154, 175, 177, 185, 186, 187, 188, 267, 271, 465, 468, 469, 471)
        Call SetComboBox(SelectComboBox, "CBp")
        UF1.ExpertORPD.Value = True
        UF1.ComboBoxTechUsrtvo.ListIndex = 0
        UF1.ComboBoxRaschet.ListIndex = 1
        ActiveDocument.Variables("TipTehUstr").Value = "водогрейный котел"
        If BaseUstrIndex = UBound(BaseUstr, 1) Then UF1.NaznTehUstr.Value = "нагрев сетевой воды"
        UF1.CBFNPORPD.Value = True
        UF1.CBSO469.Value = True
        Call UstSosAktNK(1, 1, 1, 1, 1, 0, 0, 1, 0)
    End If
    If UF1.ComboBoxTipUstroistva.ListIndex = 2 Then
        Call ClearAllF
        SelectComboBox = Array(2, 3, 10, 22, 39, 43, 50, 100, 154, 175, 178, 185, 186, 187, 188, 257, 258, 260, 267, 271, 465, 468, 471, 500, 502, 503, 505, 506)
        Call SetComboBox(SelectComboBox, "CBp")
        UF1.ExpertORPD.Value = True
        UF1.ComboBoxTechUsrtvo.ListIndex = 0
        UF1.ComboBoxRaschet.ListIndex = 0
        ActiveDocument.Variables("TipTehUstr").Value = "электрокотел"
        If BaseUstrIndex = UBound(BaseUstr, 1) Then UF1.NaznTehUstr.Value = "выработка пара"
        UF1.CBFNPORPD.Value = True
        UF1.CBSO469.Value = True
    End If
    If UF1.ComboBoxTipUstroistva.ListIndex = 3 Then
        Call ClearAllF
        SelectComboBox = Array(2, 3, 10, 45, 61, 64, 175, 177, 185, 186, 187, 188, 267, 268, 269, 270, 271, 465, 468, 469, 471)
        Call SetComboBox(SelectComboBox, "CBp")
        UF1.ExpertORPD.Value = True
        UF1.ComboBoxTechUsrtvo.ListIndex = 1
        UF1.ComboBoxRaschet.ListIndex = 1
        ActiveDocument.Variables("TipTehUstr").Value = "водяной экономайзер"
        If BaseUstrIndex = UBound(BaseUstr, 1) Then UF1.NaznTehUstr.Value = "подогрев питательной воды"
        UF1.CBFNPORPD.Value = True
        UF1.CBSO469.Value = True
    End If
End If
If UF1.OptionTruboprovod.Value = True Then
    If UF1.ComboBoxTipUstroistva.ListIndex = 0 Then
        Call ClearAllF
'        SelectComboBox = Array(2, 3, 10, 71, 80, 81, 85, 86, 90, 91, 154, 156, 175, 184, 185, 187, 188, 372, 373, 374, 394, 465, 468, 469, 471)
'        Call SetComboBox(SelectComboBox, "CBp")
        UF1.ExpertORPD.Value = True
        UF1.ComboBoxTechUsrtvo.ListIndex = 3
        UF1.ComboBoxRaschet.ListIndex = 1
        ActiveDocument.Variables("TipTehUstr").Value = "трубопровод пара"
        If BaseUstrIndex = UBound(BaseUstr, 1) Then UF1.NaznTehUstr.Value = "транспортировка пара"
        UF1.CBFNPORPD.Value = True
        UF1.CBSO464.Value = True
        Call UstSosAktNK(0, 1, 1, 0, 1, 0, 0, 1, 0)
    End If
    If UF1.ComboBoxTipUstroistva.ListIndex = 1 Then
        Call ClearAllF
        SelectComboBox = Array(2, 3, 10, 71, 80, 81, 85, 86, 90, 91, 154, 156, 175, 184, 185, 187, 188, 372, 373, 374, 394, 465, 468, 469, 471)
        Call SetComboBox(SelectComboBox, "CBp")
        UF1.ExpertORPD.Value = True
        UF1.ComboBoxTechUsrtvo.ListIndex = 3
        UF1.ComboBoxRaschet.ListIndex = 1
        ActiveDocument.Variables("TipTehUstr").Value = "трубопровод горячего водоснабжения"
        If BaseUstrIndex = UBound(BaseUstr, 1) Then UF1.NaznTehUstr.Value = "транспортировка сетевой воды"
        UF1.CBFNPORPD.Value = True
        UF1.CBSO464.Value = True
    End If
End If
If UF1.OptionSoorugenie.Value = True Then
    If UF1.ComboBoxTipUstroistva.ListIndex = 0 Then
        Call ClearAllF
        SelectComboBox = Array(11, 12, 15, 132, 135, 137, 140, 233, 234, 255, 267)
        Call SetComboBox(SelectComboBox, "CBho")
        UF1.ExpertHim.Value = True
        UF1.ComboBoxTechUsrtvo.ListIndex = 11
        UF1.ComboBoxRaschet.ListIndex = 0
        ActiveDocument.Variables("TipTehUstr").Value = "резервуар"
        If BaseUstrIndex = UBound(BaseUstr, 1) Then UF1.NaznTehUstr.Value = "емкость для хранения"
        UF1.CBFNPHOPO.Value = True
        UF1.CBRD1533413752696.Value = True
    End If
    If UF1.ComboBoxTipUstroistva.ListIndex = 1 Then
        SelectComboBox = Array(30, 119, 161, 164, 169, 177, 178, 179, 193, 196, 197, 198, 199, 203)
        Call SetComboBox(SelectComboBox, "CBvb")
        SelectComboBox = Array(15, 132, 135, 136, 137, 149, 150, 151, 152)
        Call SetComboBox(SelectComboBox, "CBho")
        SelectComboBox = Array(27, 29, 35, 36, 59, 65, 85, 94, 141, 145, 148, 190, 191)
        Call SetComboBox(SelectComboBox, "CBtt")
        SelectComboBox = Array(137, 141, 142, 144, 146, 147, 148, 149, 150)
        Call SetComboBox(SelectComboBox, "CBsn")
        UF1.ExpertHim.Value = True
        UF1.ComboBoxTechUsrtvo.ListIndex = 13
        UF1.ComboBoxRaschet.ListIndex = 3
        UF1.Label501.Caption = "Труба"
        UF1.Label502.Caption = "Труба"
        UF1.ZavN.Value = "б/н"
        ActiveDocument.Variables("TipTehUstr").Value = "технологический трубопровод"
        If BaseUstrIndex = UBound(BaseUstr, 1) Then UF1.NaznTehUstr.Value = "транспортировка "
        UF1.CBFNPHOPO.Value = True
        UF1.CBFNPPBETT.Value = True
'        UF1.CBRekpoTT.Value = True
    End If
    If UF1.ComboBoxTipUstroistva.ListIndex = 2 Then
        SelectComboBox = Array(15, 132, 135, 136, 137, 140, 149, 150, 151, 152, 234, 238, 240, 241, 242, 244, 246, 247, 248)
        Call SetComboBox(SelectComboBox, "CBho")
        SelectComboBox = Array(30, 161, 164, 169, 177, 178, 179, 193, 196, 197, 198, 199, 203)
        Call SetComboBox(SelectComboBox, "CBvb")
        SelectComboBox = Array(27, 29, 35, 36, 44, 59, 65, 85, 94, 100, 141, 145, 148)
        Call SetComboBox(SelectComboBox, "CBtt")
        UF1.ExpertHim.Value = True
        UF1.ComboBoxTechUsrtvo.ListIndex = 13
        UF1.ComboBoxRaschet.ListIndex = 3
        ActiveDocument.Variables("TipTehUstr").Value = "технологический трубопровод"
        If BaseUstrIndex = UBound(BaseUstr, 1) Then UF1.NaznTehUstr.Value = "транспортировка "
        UF1.Label501.Caption = "Труба"
        UF1.Label502.Caption = "Труба"
        UF1.CBFNPHOPO.Value = True
        UF1.CBFNPPBETT.Value = True
        UF1.CBRekpoTT.Value = True
    End If
    If UF1.ComboBoxTipUstroistva.ListIndex = 3 Then
        Call ClearAllF
        SelectComboBox = Array(120, 121, 164, 177, 178, 179)
        Call SetComboBox(SelectComboBox, "CBvb")
        SelectComboBox = Array(77, 81, 87, 94, 98, 102, 104, 105)
        Call SetComboBox(SelectComboBox, "CBsn")
        UF1.ExpertHim.Value = True
        UF1.ComboBoxTechUsrtvo.ListIndex = 12
        UF1.ComboBoxRaschet.ListIndex = 0
        ActiveDocument.Variables("TipTehUstr").Value = "резервуар"
        If BaseUstrIndex = UBound(BaseUstr, 1) Then UF1.NaznTehUstr.Value = "резервуар для хранения"
        UF1.CBFNPPBSNN.Value = True
        UF1.CBFNPOPVB.Value = True
        UF1.CBRD089595.Value = True
    End If
End If
If UF1.OptionOstalnoe.Value = True Then
    If UF1.ComboBoxTipUstroistva.ListIndex = 0 Then
        ActiveDocument.Variables("TipTehUstr").Value = "насос"
        ActiveDocument.Variables("NaznTehUstr").Value = "предназначен для перекачки "
        ActiveDocument.Variables("MnNum7").Value = "Позиция по технологической схеме"
        UF1.Label476.Caption = "P"
        UF1.Label475.Caption = "напор"
        UF1.Label474.Caption = "Q"
        UF1.ComboBoxRaschet.ListIndex = 6
        UF1.ComboBoxTechUsrtvo.ListIndex = 15
        UF1.ProtokolVD.Visible = True
        UF1.ProtokolVDD.Visible = True
        UF1.Label465.Visible = True
        UF1.Label466.Visible = True
        Call ClearAllF
        SelectComboBox = Array(43, 47, 48, 53, 161, 164, 177, 178, 179, 184, 185, 186, 189, 190)
        Call SetComboBox(SelectComboBox, "CBvb")
        SelectComboBox = Array(156, 157, 159, 160, 167, 168, 171)
        Call SetComboBox(SelectComboBox, "CBsn")
        SelectComboBox = Array(11, 12, 15, 126, 132, 135, 136, 137, 140, 142, 144, 145)
        Call SetComboBox(SelectComboBox, "CBho")
    Else
        UF1.ProtokolVD.Visible = False
        UF1.ProtokolVDD.Visible = False
        UF1.Label465.Visible = False
        UF1.Label466.Visible = False
    End If
End If

End Sub

Private Sub MPDZD_Change()
If UF1.MPDZD.Value = True Then
    UF1.ZakMPDZD.Enabled = True
    UF1.ZakMPDZDD.Enabled = True
    UF1.MPDZD.Caption = "МПД"
Else
    If UF1.MPDZD.Value = False Then
        UF1.ZakMPDZD.Enabled = False
        UF1.ZakMPDZDD.Enabled = False
        UF1.MPDZD.Caption = "Нет"
    Else
        UF1.ZakMPDZD.Enabled = True
        UF1.ZakMPDZDD.Enabled = True
        UF1.MPDZD.Caption = "ЦД"
    End If
End If

End Sub

Private Sub ogost_Change()
UF1.ogost.Value = Trim(UF1.ogost.Value)
UF1.dgost.Value = UF1.ogost.Value
ActiveDocument.Variables("ogost").Value = UF1.ogost.Value
End Sub

Private Sub OptionKotel_Change()
If UF1.OptionKotel.Value = True Then
    Me.ComboBoxTipUstroistva.Clear
    Me.ComboBoxTipUstroistva.List = Array("паровой котел", "водогрейный котел", "электрический котел", "экономайзер")
    UF1.CBVakuum.Enabled = False
    UF1.CBPodNaliv.Enabled = False
    UF1.VikMK.Value = True
    UF1.KontrGibCh.Value = True
    UF1.ComboBoxTechUsrtvo.ListIndex = 0
    UF1.ComboBoxRaschet.ListIndex = 1
    UF1.Label418.Caption = "Вид топлива"
    UF1.KolZicl.Visible = False
    UF1.CBZikl.Visible = False
    UF1.CBZikl.Value = False
    UF1.Label18.Visible = False
    UF1.Label5.Caption = "Завод изг."
    UF1.Label501.Caption = "Обечайка"
    UF1.Label502.Caption = "Днище"
    ActiveDocument.Variables("TimeGI").Value = "20 минут"
    ActiveDocument.Variables("P7RabSredaTopl").Value = "вид топлива"
    SelectComboBox = Array(2, 3, 10, 38, 39, 43, 45, 47, 49, 50, 100, 154, 175, 177, 185, 186, 187, 188, 257, 258, 260, 267, 271, 465, 468, 471)
    Call SetComboBox(SelectComboBox, "CBp")
    ActiveDocument.Variables("PunktPril8ORPD").Value = ", п.п. 2, 3, 4, 5 Приложения №8"
    
    Set UstrIndxs = Nothing
    UF1.RegN.Clear
    UF1.ZavN.Clear
    For i = 1 To UBound(BaseUstr, 1) - 1
        If UF1.RegNOPO.Value = BaseUstr(i, 1) And BaseUstr(i, 2) = "Котел" Then
            UF1.ZavN.AddItem BaseUstr(i, 7)
            UF1.RegN.AddItem BaseUstr(i, 4)
            UstrIndxs.Add i
        End If
    Next i
    If UF1.RegN.ListCount <> 0 Then
        UF1.RegN.ListIndex = 0
    Else
        BaseUstrIndex = UBound(BaseUstr, 1)
    End If
    Call FillOutElements(BaseUstrIndex, BaseElements)

End If
End Sub

Private Sub OptionSosud_Change()
If UF1.OptionSosud.Value = True Then
    With UF1.ComboBoxTipUstroistva
        .Clear
        .AddItem "Сосуд под давлением"
        .AddItem "Воздухосборник"
        .AddItem "НЖУ, ЦЖУ, УДХ"
        .AddItem "Баллоны групповой установки"
        .AddItem "Автоцистерна для СУГ"
        .AddItem "Подогреватель"
        .AddItem "Газификатор ГХК"
        .AddItem "Сосуд с вакуумом"
        .AddItem "Сосуд под налив"
        .AddItem "Бак кислотный"
        .AddItem "Баллон"
    End With
    UF1.ComboBoxTechUsrtvo.ListIndex = 2
    UF1.ComboBoxRaschet.ListIndex = 0
    UF1.CBVakuum.Enabled = True
    UF1.CBPodNaliv.Enabled = True
    UF1.KolZicl.Visible = True
    UF1.CBZikl.Visible = True
    UF1.CBZikl.Value = True
    UF1.VKorp.Enabled = True
    UF1.Label18.Visible = True
    UF1.Label5.Caption = "Завод изг."
    UF1.Label501.Caption = "Обечайка"
    UF1.Label502.Caption = "Днище"
    ActiveDocument.Variables("TimeGI").Value = "10 минут"
    ActiveDocument.Variables("P7RabSredaTopl").Value = "рабочая среда"
    SelectComboBox = Array(30, 161, 164, 169, 177, 178, 179)
    Call SetComboBox(SelectComboBox, "CBvb")
    SelectComboBox = Array(11, 12, 15, 126, 132, 135, 136, 137, 140)
    Call SetComboBox(SelectComboBox, "CBho")
    SelectComboBox = Array(2, 3, 10, 65, 68, 69, 154, 175, 178, 185, 187, 188, 338, 339, 340, 341, 343, 353, 465, 468, 471)
    Call SetComboBox(SelectComboBox, "CBp")
    ActiveDocument.Variables("PunktPril8ORPD").Value = ", п.п. 3, 5 Приложения №8"
    UF1.VikMK.Value = False
    UF1.Ovalnost.Value = False
    UF1.ProgibCb.Value = False
    UF1.KontrGibCh.Value = False
    
    Set UstrIndxs = Nothing
    UF1.RegN.Clear
    UF1.ZavN.Clear
    For i = 1 To UBound(BaseUstr, 1) - 1
        If UF1.RegNOPO.Value = BaseUstr(i, 1) And BaseUstr(i, 2) = "Сосуд" Then
            UF1.ZavN.AddItem BaseUstr(i, 7)
            UF1.RegN.AddItem BaseUstr(i, 4)
            UstrIndxs.Add i
        End If
    Next i
    If UF1.RegN.ListCount <> 0 Then
        UF1.RegN.ListIndex = 0
    Else
        BaseUstrIndex = UBound(BaseUstr, 1)
    End If
    Call FillOutElements(BaseUstrIndex, BaseElements)
End If
End Sub

Private Sub OptionTruboprovod_Change()
If UF1.OptionTruboprovod.Value = True Then
    With UF1.ComboBoxTipUstroistva
        .Clear
        .AddItem "трубопровод пара"
        .AddItem "трубопровод горячей воды"
    End With
    UF1.ComboBoxTechUsrtvo.ListIndex = 3
    UF1.ComboBoxRaschet.ListIndex = 1
    UF1.VKorp.Enabled = False
    UF1.VRub.Enabled = False
    UF1.KolZicl.Visible = False
    UF1.CBZikl.Visible = False
    UF1.CBZikl.Value = False
    UF1.CBRubashka.Caption = "РОУ"
    UF1.Label18.Visible = False
    UF1.Label5.Caption = "Монт. орг."
    UF1.Label501.Caption = "Труба"
    UF1.Label502.Caption = "Труба"
    ActiveDocument.Variables("TimeGI").Value = "10 минут"
    ActiveDocument.Variables("P7RabSredaTopl").Value = "рабочая среда"
    SelectComboBox = Array(2, 3, 10, 71, 80, 81, 85, 86, 90, 91, 154, 156, 175, 184, 185, 187, 188, 372, 373, 374, 394, 465, 468, 469, 471)
    Call SetComboBox(SelectComboBox, "CBp")
    ActiveDocument.Variables("PunktPril8ORPD").Value = ", п.п. 3, 5 Приложения №8"
    UF1.VikMK.Value = False
    UF1.Ovalnost.Value = False
    UF1.ProgibCb.Value = False
    UF1.KontrGibCh.Value = False
    
    Set UstrIndxs = Nothing
    UF1.RegN.Clear
    UF1.ZavN.Clear
    For i = 1 To UBound(BaseUstr, 1) - 1
        If UF1.RegNOPO.Value = BaseUstr(i, 1) And BaseUstr(i, 2) = "Трубопровод" Then
            UF1.ZavN.AddItem BaseUstr(i, 7)
            UF1.RegN.AddItem BaseUstr(i, 4)
            UstrIndxs.Add i
        End If
    Next i
    If UF1.RegN.ListCount <> 0 Then
        UF1.RegN.ListIndex = 0
    Else
        BaseUstrIndex = UBound(BaseUstr, 1)
    End If
    Call FillOutElements(BaseUstrIndex, BaseElements)
    UF1.VKorp.Value = ""
    UF1.VRub.Value = ""
End If
End Sub

Private Sub ostal_Change()
UF1.ostal.Value = Trim(UF1.ostal.Value)
UF1.dstal.Value = UF1.ostal.Value
ActiveDocument.Variables("ostal").Value = UF1.ostal.Value
End Sub

Private Sub otolsh_Change()
    UF1.otolsh.Value = Trim(UF1.otolsh.Value)
    UF1.otolsh.Value = Replace(UF1.otolsh.Value, ".", ",")
    UF1.dtolsh.Value = UF1.otolsh.Value
    ActiveDocument.Variables("otolsh").Value = UF1.otolsh.Value
End Sub

Private Sub Ovalnost_Change()
If UF1.Ovalnost.Value = True Then
    UF1.KartOvaln.Enabled = True
    UF1.KartOvalnD.Enabled = True
    UF1.Ovalnost.Caption = "Есть"
Else
    UF1.KartOvaln.Enabled = False
    UF1.KartOvalnD.Enabled = False
    UF1.Ovalnost.Caption = "Нет"
End If
End Sub

Private Sub ProtkTV_Change()
ActiveDocument.Variables("ProtkTV").Value = UF1.ProtkTV.Value
If UF1.ProtkTV.Enabled = True Then
    UF1.KartOvaln.Value = Val(UF1.ProtkTV.Value) + 1
Else
    UF1.KartOvaln.Value = Val(UF1.ProtkTV.Value)
End If
End Sub

Private Sub ProtokTolch_Change()
ActiveDocument.Variables("ProtokTolch").Value = UF1.ProtokTolch.Value
If UF1.ProtokTolch.Enabled = True Then
    UF1.ZakUZK.Value = Val(UF1.ProtokTolch.Value) + 1
Else
    UF1.ZakUZK.Value = Val(UF1.ProtokTolch.Value)
End If
End Sub

Private Sub R10C1_Change()
    BaseElements(BaseUstrIndex, 76 + (8 * (UF1.SpButElm.Value))) = UF1.R10C1.Value
End Sub

Private Sub R10C2_Change()
    BaseElements(BaseUstrIndex, 77 + (8 * (UF1.SpButElm.Value))) = UF1.R10C2.Value
End Sub

Private Sub R10C3_Change()
    BaseElements(BaseUstrIndex, 78 + (8 * (UF1.SpButElm.Value))) = UF1.R10C3.Value
End Sub

Private Sub R10C4_Change()
    BaseElements(BaseUstrIndex, 79 + (8 * (UF1.SpButElm.Value))) = UF1.R10C4.Value
End Sub

Private Sub R10C5_Change()
    If IsNumeric(UF1.R10C4.Value) And IsNumeric(UF1.R10C5.Value) And Val(ActiveDocument.Variables("SrokSlugb").Value) <> 0 Then
    UF1.LabelR10.Caption = "Скор.корроз. = " & Format((UF1.R10C4.Value - UF1.R10C5.Value) / Val(ActiveDocument.Variables("SrokSlugb").Value), "0.0###")
    Else
        UF1.LabelR10.Caption = "" '"Скор.корроз. = " & Format((UF1.R1C4.Value - UF1.R1C5.Value) / Val(ActiveDocument.Variables("SrokSlugb").Value), "0.0###")
    End If
    BaseElements(BaseUstrIndex, 80 + (8 * (UF1.SpButElm.Value))) = UF1.R10C5.Value
End Sub

Private Sub R10C6_Change()
    BaseElements(BaseUstrIndex, 81 + (8 * (UF1.SpButElm.Value))) = UF1.R10C6.Value
End Sub

Private Sub R10C7_Change()
    BaseElements(BaseUstrIndex, 82 + (8 * (UF1.SpButElm.Value))) = UF1.R10C7.Value
End Sub

Private Sub R1C1_Change()
'    ActiveDocument.Tables(1).Cell(Row:=2, Column:=2).Range = UF1.R1C1.Value
    BaseElements(BaseUstrIndex, 4 + (8 * (UF1.SpButElm.Value))) = UF1.R1C1.Value
End Sub

Private Sub R1C2_Change()
    UF1.odlina.Value = UF1.R1C2.Value
'    ActiveDocument.Tables(1).Cell(Row:=2, Column:=3).Range = UF1.R1C2.Value
    BaseElements(BaseUstrIndex, 5 + (8 * (UF1.SpButElm.Value))) = UF1.R1C2.Value
End Sub

Private Sub R1C3_Change()
    UF1.odiam.Value = UF1.R1C3.Value
'    ActiveDocument.Tables(1).Cell(Row:=2, Column:=4).Range = UF1.R1C3.Value
    BaseElements(BaseUstrIndex, 6 + (8 * (UF1.SpButElm.Value))) = UF1.R1C3.Value
End Sub

Private Sub R1C4_Change()
    UF1.otolsh.Value = UF1.R1C4.Value
'    ActiveDocument.Tables(1).Cell(Row:=2, Column:=5).Range = UF1.R1C4.Value
    BaseElements(BaseUstrIndex, 7 + (8 * (UF1.SpButElm.Value))) = UF1.R1C4.Value
End Sub

Private Sub R1C5_Change()
    UF1.otolshfakt.Value = UF1.R1C5.Value
    If IsNumeric(UF1.R1C4.Value) And IsNumeric(UF1.R1C5.Value) And Val(ActiveDocument.Variables("SrokSlugb").Value) <> 0 Then
        UF1.LabelR1.Caption = "Скор.корроз. = " & Format((UF1.R1C4.Value - UF1.R1C5.Value) / Val(ActiveDocument.Variables("SrokSlugb").Value), "0.0###")
    Else
        UF1.LabelR1.Caption = "" '"Скор.корроз. = " & Format((UF1.R1C4.Value - UF1.R1C5.Value) / Val(ActiveDocument.Variables("SrokSlugb").Value), "0.0###")
    End If
    BaseElements(BaseUstrIndex, 8 + (8 * (UF1.SpButElm.Value))) = UF1.R1C5.Value
End Sub

Private Sub R1C6_Change()
    UF1.ostal.Value = UF1.R1C6.Value
'    ActiveDocument.Tables(1).Cell(Row:=2, Column:=6).Range = UF1.R1C6.Value
    BaseElements(BaseUstrIndex, 9 + (8 * (UF1.SpButElm.Value))) = UF1.R1C6.Value
End Sub

Private Sub R1C7_Change()
    UF1.ogost.Value = UF1.R1C7.Value
'    ActiveDocument.Tables(1).Cell(Row:=2, Column:=7).Range = UF1.R1C7.Value
    BaseElements(BaseUstrIndex, 10 + (8 * (UF1.SpButElm.Value))) = UF1.R1C7.Value
End Sub

Private Sub R2C1_Change()
    BaseElements(BaseUstrIndex, 12 + (8 * (UF1.SpButElm.Value))) = UF1.R2C1.Value
End Sub

Private Sub R2C2_Change()
    UF1.ddlina.Value = UF1.R2C2.Value
    BaseElements(BaseUstrIndex, 13 + (8 * (UF1.SpButElm.Value))) = UF1.R2C2.Value
End Sub

Private Sub R2C3_Change()
    UF1.ddiam.Value = UF1.R2C3.Value
    BaseElements(BaseUstrIndex, 14 + (8 * (UF1.SpButElm.Value))) = UF1.R2C3.Value
End Sub

Private Sub R2C4_Change()
    UF1.dtolsh.Value = UF1.R2C4.Value
    BaseElements(BaseUstrIndex, 15 + (8 * (UF1.SpButElm.Value))) = UF1.R2C4.Value
End Sub

Private Sub R2C5_Change()
    UF1.dtolshfakt.Value = UF1.R2C5.Value
    If IsNumeric(UF1.R2C4.Value) And IsNumeric(UF1.R2C5.Value) And Val(ActiveDocument.Variables("SrokSlugb").Value) <> 0 Then
        UF1.LabelR2.Caption = "Скор.корроз. = " & Format((UF1.R2C4.Value - UF1.R2C5.Value) / Val(ActiveDocument.Variables("SrokSlugb").Value), "0.0###")
    Else
        UF1.LabelR2.Caption = "" '"Скор.корроз. = " & Format((UF1.R1C4.Value - UF1.R1C5.Value) / Val(ActiveDocument.Variables("SrokSlugb").Value), "0.0###")
    End If
    BaseElements(BaseUstrIndex, 16 + (8 * (UF1.SpButElm.Value))) = UF1.R2C5.Value
End Sub

Private Sub R2C6_Change()
    UF1.dstal.Value = UF1.R2C6.Value
    BaseElements(BaseUstrIndex, 17 + (8 * (UF1.SpButElm.Value))) = UF1.R2C6.Value
End Sub

Private Sub R2C7_Change()
    UF1.dgost.Value = UF1.R2C7.Value
    BaseElements(BaseUstrIndex, 18 + (8 * (UF1.SpButElm.Value))) = UF1.R2C7.Value
End Sub

Private Sub R3C1_Change()
    BaseElements(BaseUstrIndex, 20 + (8 * (UF1.SpButElm.Value))) = UF1.R3C1.Value
End Sub

Private Sub R3C2_Change()
    BaseElements(BaseUstrIndex, 21 + (8 * (UF1.SpButElm.Value))) = UF1.R3C2.Value
End Sub

Private Sub R3C3_Change()
    BaseElements(BaseUstrIndex, 22 + (8 * (UF1.SpButElm.Value))) = UF1.R3C3.Value
End Sub

Private Sub R3C4_Change()
    BaseElements(BaseUstrIndex, 23 + (8 * (UF1.SpButElm.Value))) = UF1.R3C4.Value
End Sub

Private Sub R3C5_Change()
    If IsNumeric(UF1.R3C4.Value) And IsNumeric(UF1.R3C5.Value) And Val(ActiveDocument.Variables("SrokSlugb").Value) <> 0 Then
        UF1.LabelR3.Caption = "Скор.корроз. = " & Format((UF1.R3C4.Value - UF1.R3C5.Value) / Val(ActiveDocument.Variables("SrokSlugb").Value), "0.0###")
    Else
        UF1.LabelR3.Caption = "" '"Скор.корроз. = " & Format((UF1.R1C4.Value - UF1.R1C5.Value) / Val(ActiveDocument.Variables("SrokSlugb").Value), "0.0###")
    End If
    BaseElements(BaseUstrIndex, 24 + (8 * (UF1.SpButElm.Value))) = UF1.R3C5.Value
End Sub

Private Sub R3C6_Change()
    BaseElements(BaseUstrIndex, 25 + (8 * (UF1.SpButElm.Value))) = UF1.R3C6.Value
End Sub

Private Sub R3C7_Change()
    BaseElements(BaseUstrIndex, 26 + (8 * (UF1.SpButElm.Value))) = UF1.R3C7.Value
End Sub

Private Sub R4C1_Change()
    BaseElements(BaseUstrIndex, 28 + (8 * (UF1.SpButElm.Value))) = UF1.R4C1.Value
End Sub

Private Sub R4C2_Change()
    BaseElements(BaseUstrIndex, 29 + (8 * (UF1.SpButElm.Value))) = UF1.R4C2.Value
End Sub

Private Sub R4C3_Change()
    BaseElements(BaseUstrIndex, 30 + (8 * (UF1.SpButElm.Value))) = UF1.R4C3.Value
End Sub

Private Sub R4C4_Change()
    BaseElements(BaseUstrIndex, 31 + (8 * (UF1.SpButElm.Value))) = UF1.R4C4.Value
End Sub

Private Sub R4C5_Change()
    If IsNumeric(UF1.R4C4.Value) And IsNumeric(UF1.R4C5.Value) And Val(ActiveDocument.Variables("SrokSlugb").Value) <> 0 Then
        UF1.LabelR4.Caption = "Скор.корроз. = " & Format((UF1.R4C4.Value - UF1.R4C5.Value) / Val(ActiveDocument.Variables("SrokSlugb").Value), "0.0###")
    Else
        UF1.LabelR4.Caption = "" '"Скор.корроз. = " & Format((UF1.R1C4.Value - UF1.R1C5.Value) / Val(ActiveDocument.Variables("SrokSlugb").Value), "0.0###")
    End If
    BaseElements(BaseUstrIndex, 32 + (8 * (UF1.SpButElm.Value))) = UF1.R4C5.Value
End Sub

Private Sub R4C6_Change()
    BaseElements(BaseUstrIndex, 33 + (8 * (UF1.SpButElm.Value))) = UF1.R4C6.Value
End Sub

Private Sub R4C7_Change()
    BaseElements(BaseUstrIndex, 34 + (8 * (UF1.SpButElm.Value))) = UF1.R4C7.Value
End Sub

Private Sub R5C1_Change()
    BaseElements(BaseUstrIndex, 36 + (8 * (UF1.SpButElm.Value))) = UF1.R5C1.Value
End Sub

Private Sub R5C2_Change()
    BaseElements(BaseUstrIndex, 37 + (8 * (UF1.SpButElm.Value))) = UF1.R5C2.Value
End Sub

Private Sub R5C3_Change()
    BaseElements(BaseUstrIndex, 38 + (8 * (UF1.SpButElm.Value))) = UF1.R5C3.Value
End Sub

Private Sub R5C4_Change()
    BaseElements(BaseUstrIndex, 39 + (8 * (UF1.SpButElm.Value))) = UF1.R5C4.Value
End Sub

Private Sub R5C5_Change()
    If IsNumeric(UF1.R5C4.Value) And IsNumeric(UF1.R5C5.Value) And Val(ActiveDocument.Variables("SrokSlugb").Value) <> 0 Then
        UF1.LabelR5.Caption = "Скор.корроз. = " & Format((UF1.R5C4.Value - UF1.R5C5.Value) / Val(ActiveDocument.Variables("SrokSlugb").Value), "0.0###")
    Else
        UF1.LabelR5.Caption = "" '"Скор.корроз. = " & Format((UF1.R1C4.Value - UF1.R1C5.Value) / Val(ActiveDocument.Variables("SrokSlugb").Value), "0.0###")
    End If
    BaseElements(BaseUstrIndex, 40 + (8 * (UF1.SpButElm.Value))) = UF1.R5C5.Value
End Sub

Private Sub R5C6_Change()
    BaseElements(BaseUstrIndex, 41 + (8 * (UF1.SpButElm.Value))) = UF1.R5C6.Value
End Sub

Private Sub R5C7_Change()
    BaseElements(BaseUstrIndex, 42 + (8 * (UF1.SpButElm.Value))) = UF1.R5C7.Value
End Sub

Private Sub R6C1_Change()
    BaseElements(BaseUstrIndex, 44 + (8 * (UF1.SpButElm.Value))) = UF1.R6C1.Value
End Sub

Private Sub R6C2_Change()
    BaseElements(BaseUstrIndex, 45 + (8 * (UF1.SpButElm.Value))) = UF1.R6C2.Value
End Sub

Private Sub R6C3_Change()
    BaseElements(BaseUstrIndex, 46 + (8 * (UF1.SpButElm.Value))) = UF1.R6C3.Value
End Sub

Private Sub R6C4_Change()
    BaseElements(BaseUstrIndex, 47 + (8 * (UF1.SpButElm.Value))) = UF1.R6C4.Value
End Sub

Private Sub R6C5_Change()
    If IsNumeric(UF1.R6C4.Value) And IsNumeric(UF1.R6C5.Value) And Val(ActiveDocument.Variables("SrokSlugb").Value) <> 0 Then
        UF1.LabelR6.Caption = "Скор.корроз. = " & Format((UF1.R6C4.Value - UF1.R6C5.Value) / Val(ActiveDocument.Variables("SrokSlugb").Value), "0.0###")
    Else
        UF1.LabelR6.Caption = "" '"Скор.корроз. = " & Format((UF1.R1C4.Value - UF1.R1C5.Value) / Val(ActiveDocument.Variables("SrokSlugb").Value), "0.0###")
    End If
    BaseElements(BaseUstrIndex, 48 + (8 * (UF1.SpButElm.Value))) = UF1.R6C5.Value
End Sub

Private Sub R6C6_Change()
    BaseElements(BaseUstrIndex, 49 + (8 * (UF1.SpButElm.Value))) = UF1.R6C6.Value
End Sub

Private Sub R6C7_Change()
    BaseElements(BaseUstrIndex, 50 + (8 * (UF1.SpButElm.Value))) = UF1.R6C7.Value
End Sub

Private Sub R7C1_Change()
    BaseElements(BaseUstrIndex, 52 + (8 * (UF1.SpButElm.Value))) = UF1.R7C1.Value
End Sub

Private Sub R7C2_Change()
    BaseElements(BaseUstrIndex, 53 + (8 * (UF1.SpButElm.Value))) = UF1.R7C2.Value
End Sub

Private Sub R7C3_Change()
    BaseElements(BaseUstrIndex, 54 + (8 * (UF1.SpButElm.Value))) = UF1.R7C3.Value
End Sub

Private Sub R7C4_Change()
    BaseElements(BaseUstrIndex, 55 + (8 * (UF1.SpButElm.Value))) = UF1.R7C4.Value
End Sub

Private Sub R7C5_Change()
    If IsNumeric(UF1.R7C4.Value) And IsNumeric(UF1.R7C5.Value) And Val(ActiveDocument.Variables("SrokSlugb").Value) <> 0 Then
        UF1.LabelR7.Caption = "Скор.корроз. = " & Format((UF1.R7C4.Value - UF1.R7C5.Value) / Val(ActiveDocument.Variables("SrokSlugb").Value), "0.0###")
    Else
        UF1.LabelR7.Caption = "" '"Скор.корроз. = " & Format((UF1.R1C4.Value - UF1.R1C5.Value) / Val(ActiveDocument.Variables("SrokSlugb").Value), "0.0###")
    End If
    BaseElements(BaseUstrIndex, 56 + (8 * (UF1.SpButElm.Value))) = UF1.R7C5.Value
End Sub

Private Sub R7C6_Change()
    BaseElements(BaseUstrIndex, 57 + (8 * (UF1.SpButElm.Value))) = UF1.R7C6.Value
End Sub

Private Sub R7C7_Change()
    BaseElements(BaseUstrIndex, 58 + (8 * (UF1.SpButElm.Value))) = UF1.R7C7.Value
End Sub

Private Sub R8C1_Change()
    BaseElements(BaseUstrIndex, 60 + (8 * (UF1.SpButElm.Value))) = UF1.R8C1.Value
End Sub

Private Sub R8C2_Change()
    BaseElements(BaseUstrIndex, 61 + (8 * (UF1.SpButElm.Value))) = UF1.R8C2.Value
End Sub

Private Sub R8C3_Change()
    BaseElements(BaseUstrIndex, 62 + (8 * (UF1.SpButElm.Value))) = UF1.R8C3.Value
End Sub

Private Sub R8C4_Change()
    BaseElements(BaseUstrIndex, 63 + (8 * (UF1.SpButElm.Value))) = UF1.R8C4.Value
End Sub

Private Sub R8C5_Change()
    If IsNumeric(UF1.R8C4.Value) And IsNumeric(UF1.R8C5.Value) And Val(ActiveDocument.Variables("SrokSlugb").Value) <> 0 Then
        UF1.LabelR8.Caption = "Скор.корроз. = " & Format((UF1.R8C4.Value - UF1.R8C5.Value) / Val(ActiveDocument.Variables("SrokSlugb").Value), "0.0###")
    Else
        UF1.LabelR8.Caption = "" '"Скор.корроз. = " & Format((UF1.R1C4.Value - UF1.R1C5.Value) / Val(ActiveDocument.Variables("SrokSlugb").Value), "0.0###")
    End If
    BaseElements(BaseUstrIndex, 64 + (8 * (UF1.SpButElm.Value))) = UF1.R8C5.Value
End Sub

Private Sub R8C6_Change()
    BaseElements(BaseUstrIndex, 65 + (8 * (UF1.SpButElm.Value))) = UF1.R8C6.Value
End Sub

Private Sub R8C7_Change()
    BaseElements(BaseUstrIndex, 66 + (8 * (UF1.SpButElm.Value))) = UF1.R8C7.Value
End Sub

Private Sub R9C1_Change()
    BaseElements(BaseUstrIndex, 68 + (8 * (UF1.SpButElm.Value))) = UF1.R9C1.Value
End Sub

Private Sub R9C2_Change()
    BaseElements(BaseUstrIndex, 69 + (8 * (UF1.SpButElm.Value))) = UF1.R9C2.Value
End Sub

Private Sub R9C3_Change()
    BaseElements(BaseUstrIndex, 70 + (8 * (UF1.SpButElm.Value))) = UF1.R9C3.Value
End Sub

Private Sub R9C4_Change()
    BaseElements(BaseUstrIndex, 71 + (8 * (UF1.SpButElm.Value))) = UF1.R9C4.Value
End Sub

Private Sub R9C5_Change()
    If IsNumeric(UF1.R9C4.Value) And IsNumeric(UF1.R9C5.Value) And Val(ActiveDocument.Variables("SrokSlugb").Value) <> 0 Then
        UF1.LabelR9.Caption = "Скор.корроз. = " & Format((UF1.R9C4.Value - UF1.R9C5.Value) / Val(ActiveDocument.Variables("SrokSlugb").Value), "0.0###")
    Else
        UF1.LabelR9.Caption = "" '"Скор.корроз. = " & Format((UF1.R1C4.Value - UF1.R1C5.Value) / Val(ActiveDocument.Variables("SrokSlugb").Value), "0.0###")
    End If
    BaseElements(BaseUstrIndex, 72 + (8 * (UF1.SpButElm.Value))) = UF1.R9C5.Value
End Sub

Private Sub R9C6_Change()
    BaseElements(BaseUstrIndex, 73 + (8 * (UF1.SpButElm.Value))) = UF1.R9C6.Value
End Sub

Private Sub R9C7_Change()
    BaseElements(BaseUstrIndex, 74 + (8 * (UF1.SpButElm.Value))) = UF1.R9C7.Value
End Sub

Private Sub RabocheePRub_Change()
    BaseUstr(BaseUstrIndex, 29) = UF1.RabocheePRub.Value
End Sub

Private Sub RabSreda_Change()
    BaseUstr(BaseUstrIndex, 28) = UF1.RabSreda.Value
    ActiveDocument.Variables("RabSreda").Value = Trim(UF1.RabSreda.Value)
End Sub

Private Sub RabSreda_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If UF1.RabSreda.Value Like "[Вв]одород*" Then
        UF1.CBp354.Value = True
    End If
    If UF1.RabSreda.Value Like "*[Мм]азут*" Then
        UF1.CBFNPOPVB.Value = True
        UF1.CBFNPPBSNN.Value = True
    End If
    If UF1.RabSreda.Value Like "*СУГ*" Then
        MsgBox ("СУГ")
    End If
    If UF1.RabSreda.Value Like "*[Кк]ислород*" Then MsgBox ("Кислород")
    If UF1.RabSreda.Value Like "*[Щщ][её]лочь*" Then MsgBox ("Щелочь")
    If UF1.RabSreda.Value Like "*[Кк]ислота*" Then MsgBox ("Кислота")
End Sub

Private Sub RabSredaRub_Change()
    BaseUstr(BaseUstrIndex, 31) = UF1.RabSredaRub.Value
End Sub

Private Sub RabTemp_Change()
    BaseUstr(BaseUstrIndex, 27) = UF1.RabTemp.Value
    If UF1.RabTemp.Value = "" Then
        ActiveDocument.Variables("RabTemp").Value = Strings.ChrW(31)
        ActiveDocument.Variables("RabTempP6").Value = Strings.ChrW(31)
    Else
        ActiveDocument.Variables("RabTemp").Value = UF1.RabTemp.Value
        ActiveDocument.Variables("RabTempP6").Value = ", t=" & Trim(UF1.RabTemp.Value) & Strings.ChrW(176) & "С"
    End If
End Sub

Private Sub RabTempRub_Change()
    BaseUstr(BaseUstrIndex, 30) = UF1.RabTempRub.Value
End Sub

Private Sub RaschetnP_Change()
UF1.RaschetnP.Value = Trim(UF1.RaschetnP.Value)
UF1.RaschetnP.Value = Replace(UF1.RaschetnP.Value, ".", ",")
BaseUstr(BaseUstrIndex, 17) = UF1.RaschetnP.Value
If IsNumeric(UF1.RazreshaemoeP.Value) Then
    ActiveDocument.Variables("RaschetnP").Value = Format(UF1.RaschetnP.Value, "###0.0#####") & " кгс/см" & Strings.ChrW(178)
Else
    ActiveDocument.Variables("RaschetnP").Value = UF1.RaschetnP.Value
End If
End Sub

Private Sub RaschetnP_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    UF1.RazreshaemoeP.Value = UF1.RaschetnP.Value
End Sub

Private Sub RaschetnPRub_Change()
    UF1.RaschetnPRub.Value = Replace(UF1.RaschetnPRub.Value, ".", ",")
    BaseUstr(BaseUstrIndex, 22) = UF1.RaschetnPRub.Value
    UF1.RabocheePRub.Value = UF1.RaschetnPRub.Value
End Sub

Private Sub Raschetnt_Change()
    BaseUstr(BaseUstrIndex, 18) = UF1.Raschetnt.Value
    UF1.RabTemp.Value = UF1.Raschetnt.Value
    If UF1.Raschetnt.Value = "" Then
        ActiveDocument.Variables("Raschetnt").Value = Strings.ChrW(31)
    Else
        ActiveDocument.Variables("Raschetnt").Value = ", t=" & Trim(UF1.Raschetnt.Value) & Strings.ChrW(176) & "С"
    End If
End Sub

Private Sub RaschetntRub_Change()
    BaseUstr(BaseUstrIndex, 23) = UF1.RaschetntRub.Value
    UF1.RabTempRub.Value = UF1.RaschetntRub.Value
End Sub

Private Sub RaschSreda_Change()
    BaseUstr(BaseUstrIndex, 20) = UF1.RaschSreda.Value
    ActiveDocument.Variables("RaschSreda").Value = UF1.RaschSreda.Value & ","
    UF1.RabSreda.Value = UF1.RaschSreda.Value
End Sub

Private Sub RaschSredaRub_Change()
    BaseUstr(BaseUstrIndex, 25) = UF1.RaschSredaRub.Value
    UF1.RabSredaRub.Value = UF1.RaschSredaRub.Value
End Sub

Private Sub RazreshaemoeP_Change()
    UF1.RazreshaemoeP.Value = Trim(UF1.RazreshaemoeP.Value)
    UF1.RazreshaemoeP = Replace(UF1.RazreshaemoeP, ".", ",")
    BaseUstr(BaseUstrIndex, 26) = UF1.RazreshaemoeP.Value
End Sub

Private Sub RazreshaemoeP_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If UF1.ComboBoxTipUstroistva.Value = "Автоцистерна для СУГ" And Val(UF1.RazreshaemoeP.Value) < Val(UF1.RaschetnP.Value) Then
        MsgBox ("Для сосудов с СУГ не допускается снижение давления. П.402 ФНП ОРПД")
    End If
    If IsNumeric(UF1.RazreshaemoeP.Value) Then
        UF1.IspitatP.Value = Format((UF1.RazreshaemoeP.Value * 1.25), "###0.0#")
        ActiveDocument.Variables("RazreshaemoeP").Value = Format(UF1.RazreshaemoeP.Value, "###0.0#####") & " кгс/см" & Strings.ChrW(178)
        ActiveDocument.Variables("RazreshaemoePKrt").Value = Format(UF1.RazreshaemoeP.Value, "###0.0#####")
        ActiveDocument.Variables("RazreshaemoePMP").Value = Format(CDbl(UF1.RazreshaemoeP.Value) / 10, "###0.0#####") & " МПа"
        ActiveDocument.Variables("RazreshaemoePMPKrt").Value = Format(CDbl(UF1.RazreshaemoeP.Value) / 10, "###0.0#####")
    Else
        ActiveDocument.Variables("RazreshaemoeP").Value = UF1.RazreshaemoeP.Value
        ActiveDocument.Variables("RazreshaemoePKrt").Value = UF1.RazreshaemoeP.Value
        ActiveDocument.Variables("RazreshaemoePMP").Value = Strings.ChrW(31)
        ActiveDocument.Variables("RazreshaemoePMPKrt").Value = Strings.ChrW(31)
    End If
    If UF1.CBVakuum.Value = True Then
        If UF1.ComboBoxTipUstroistva.Value = "технологический трубопровод" Then
            UF1.IspitatP.Value = "2,0"
        Else
            UF1.IspitatP.Value = "1,25"
        End If
    End If
End Sub

Private Sub RegN_Change()
    ActiveDocument.Variables("RegN").Value = Trim(UF1.RegN.Value)
    If UF1.RegN.ListIndex <> -1 Then
        UF1.ZavN.ListIndex = UF1.RegN.ListIndex
        BaseUstrIndex = UstrIndxs.Item(UF1.RegN.ListIndex + 1)
        Call FillOutFormUstr(BaseUstrIndex, BaseUstr)
        UF1.SvedORemonte.Value = BaseRemont(BaseUstrIndex, 3)
        UF1.SvedOEPB.Value = BaseEPB(BaseUstrIndex, 3)
        Call FillOutElements(BaseUstrIndex, BaseElements)
    Else
    End If
End Sub

Private Sub RegN_Exit(ByVal Cancel As MSForms.ReturnBoolean)
If UF1.RegN.ListIndex <> -1 Then
    BaseUstrIndex = UstrIndxs.Item(UF1.RegN.ListIndex + 1)
    Call FillOutElements(BaseUstrIndex, BaseElements)
Else
    BaseUstrIndex = UBound(BaseUstr, 1)
    UF1.ZavN.Clear
    BaseUstr(BaseUstrIndex, 3) = UF1.poleRegNum.Value
    BaseUstr(BaseUstrIndex, 4) = UF1.RegN.Value
    UF1.RegN.Clear
    BaseUstr(BaseUstrIndex, 15) = Trim$(UF1.CBPodNaliv.Value)
    BaseUstr(BaseUstrIndex, 16) = Trim$(UF1.CBVakuum.Value)
    BaseUstr(BaseUstrIndex, 21) = Trim$(UF1.CBRubashka.Value)
    BaseElements(BaseUstrIndex, 163) = Trim$(UF1.CBZikl.Value)
    For i = 11 To 155 Step 8
        BaseElements(BaseUstrIndex, i) = "False"
    Next i
    Call FillOutFormUstr(BaseUstrIndex, BaseUstr)
    Call FillOutElements(BaseUstrIndex, BaseElements)
End If
    UF1.SpButRemont.Value = 1
    UF1.SvedORemonte.Value = BaseRemont(BaseUstrIndex, 3)
    UF1.SpButEPB.Value = 1
    UF1.SvedOEPB.Value = BaseEPB(BaseUstrIndex, 3)
End Sub

Private Sub RegNOPO_Change()
    ActiveDocument.Variables("RegNOPO").Value = Trim(UF1.RegNOPO.Value)
    If Trim(UF1.RegNOPO.Value) = "" Then ActiveDocument.Variables("RegNOPO").Value = Strings.ChrW(31)
    If UF1.RegNOPO.ListIndex <> -1 Then
        For i = 1 To UBound(BaseOPO, 1) - 1
            If BaseOPO(i, 1) = Trim(UF1.RegNOPO.Value) Then BaseOPOIndex = i
        Next i
    End If
    If UF1.RegNOPO.ListIndex <> -1 And UF1.ClassOpasOPO.ListCount = UF1.RegNOPO.ListCount Then UF1.ClassOpasOPO.ListIndex = UF1.RegNOPO.ListIndex
    If UF1.RegNOPO.ListIndex <> -1 And UF1.NazvOPO.ListCount = UF1.RegNOPO.ListCount Then UF1.NazvOPO.ListIndex = UF1.RegNOPO.ListIndex
End Sub

Private Sub RegNOPO_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If UF1.RegNOPO.ListIndex = -1 Then
        BaseOPOIndex = UBound(BaseOPO, 1)
        BaseOPO(BaseOPOIndex, 1) = UF1.RegNOPO.Value
        UF1.ClassOpasOPO.Clear
        UF1.NazvOPO.Clear
    End If
End Sub

Private Sub SavePribor_Click()

For i = 1 To UF1.SavePribor.TabIndex / 2
'    MsgBox (UF1.MultiPage1.Pages("Page2").Controls.Item("pribor" & i).Value)
'    MsgBox (UF1.MultiPage1.Pages("Page2").Controls.Item("dp" & i).Value)
    DateBase.Workbooks("data_base.xls").Worksheets("tablprib").Range("A" & i).Value = UF1.MultiPage1.Pages("Page5").Controls.Item("pribor" & i).Value
    DateBase.Workbooks("data_base.xls").Worksheets("tablprib").Range("B" & i).Value = UF1.MultiPage1.Pages("Page5").Controls.Item("dp" & i).Value
Next i
'Call Save_File("Page2", UF1.SavePribor.TabIndex - 1, MyFilePribor)

End Sub

Private Sub PnIs_Change()
If (UF1.PnIs.Value = True) Then
    UF1.AktAE.Visible = True
    UF1.AktAED.Visible = True
    UF1.Label54.Visible = True
    UF1.Label427.Visible = True
    For Each mark In AllCBp
        If mark > 174 And mark < 191 Then UF1.Controls.Item("CBp" & mark).Value = False
    Next
    UF1.Controls.Item("CBp175").Value = True
    UF1.Controls.Item("CBp190").Value = True
Else
    UF1.AktAE.Visible = False
    UF1.AktAED.Visible = False
    UF1.Label54.Visible = False
    UF1.Label427.Visible = False
    Call ComboBoxTipUstroistva_Change
End If
End Sub

Private Sub SpButElm_Change()
    Call FillOutElements(BaseUstrIndex, BaseElements)
End Sub

Private Sub SpButEPB_Change()
    UF1.NZapisEPB.Caption = "Запись №" & UF1.SpButEPB.Value
    If IsEmpty(BaseUstrIndex) Then
    Else
        UF1.SvedOEPB.Value = BaseEPB(BaseUstrIndex, UF1.SpButEPB.Value + 2)
    End If
End Sub

Private Sub SpButRemont_Change()
    UF1.NZapisRem.Caption = "Запись №" & UF1.SpButRemont.Value
    If IsEmpty(BaseUstrIndex) Then
        Z = 1 + 1
    Else
        UF1.SvedORemonte.Value = BaseRemont(BaseUstrIndex, UF1.SpButRemont.Value + 2)
    End If
End Sub

Private Sub SvedOEPB_Change()
    BaseEPB(BaseUstrIndex, UF1.SpButEPB.Value + 2) = UF1.SvedOEPB.Value
End Sub

Private Sub SvedORemonte_Change()
    BaseRemont(BaseUstrIndex, UF1.SpButRemont.Value + 2) = UF1.SvedORemonte.Value
End Sub

Private Sub TBCopyUstr_Click()
    X = UBound(BaseUstr, 1)
    For i = 5 To UBound(BaseUstr, 2)
        BaseUstr(X, i) = BaseUstr(BaseUstrIndex, i)
    Next i
    For i = 3 To UBound(BaseElements, 2)
        BaseElements(X, i) = BaseElements(BaseUstrIndex, i)
    Next i
    For i = 3 To UBound(BaseRemont, 2)
        BaseRemont(X, i) = BaseRemont(BaseUstrIndex, i)
    Next i
    For i = 3 To UBound(BaseEPB, 2)
        BaseEPB(X, i) = BaseEPB(BaseUstrIndex, i)
    Next i
    UF1.RegN.Clear
    UF1.ZavN.Clear
    BaseUstrIndex = X
    UF1.TBCopyUstr.Value = False
End Sub

Private Sub TBUstrComment_Change()
    BaseUstr(BaseUstrIndex, 38) = UF1.TBUstrComment.Value
End Sub

Private Sub Tolshin_Change()
If UF1.Tolshin.Value = True Then
    UF1.ProtokTolch.Enabled = True
    UF1.ProtokTolchD.Enabled = True
    UF1.Tolshin.Caption = "Есть"
Else
    UF1.ProtokTolch.Enabled = False
    UF1.ProtokTolchD.Enabled = False
    UF1.Tolshin.Caption = "Нет"
End If
End Sub

Private Sub Tverdost_Change()
If UF1.Tverdost.Value = True Then
    UF1.ProtkTV.Enabled = True
    UF1.ProtkTVD.Enabled = True
    UF1.Tverdost.Caption = "Есть"
Else
    UF1.ProtkTV.Enabled = False
    UF1.ProtkTVD.Enabled = False
    UF1.Tverdost.Caption = "Нет"
End If
End Sub

Private Sub UserForm_Initialize()
'Загружаем из файла приборы
tmp1 = DateBase.Workbooks("data_base.xls").Worksheets("tablprib").Range("A1:B17").Value
Dim Flag As Boolean
Flag = False
For i = 1 To 17
    UF1.Controls.Item("pribor" & i).Value = tmp1(i, 1)
    UF1.Controls.Item("dp" & i).Value = tmp1(i, 2)
    If CDate(tmp1(i, 2)) < Date Then
        UF1.Controls.Item("dp" & i).BackColor = wdColorRed '1000
        Flag = True
    End If
Next i
If (Flag) Then
    MsgBox "Есть просроченные приборы"
End If
index = DateBase.Workbooks("data_base.xls").Worksheets("Predpriyatiya").Range("A1").Value 'Indx(tmp2, "")
BasePredp = DateBase.Workbooks("data_base.xls").Worksheets("Predpriyatiya").Range("B2:I" & index + 2).Value
BasePredpIndex = UBound(BasePredp, 1)
index = DateBase.Workbooks("data_base.xls").Worksheets("OPO").Range("A1").Value 'Indx(tmpOPO1, "")
BaseOPO = DateBase.Workbooks("data_base.xls").Worksheets("OPO").Range("B2:D" & index + 2).Value
BaseOPOIndex = UBound(BaseOPO, 1)
index = DateBase.Workbooks("data_base.xls").Worksheets("Ustroystva").Range("A1").Value
BaseUstr = DateBase.Workbooks("data_base.xls").Worksheets("Ustroystva").Range("B2:AM" & index + 2).Value
BaseUstrIndex = UBound(BaseUstr, 1)
BaseElements = DateBase.Workbooks("data_base.xls").Worksheets("Elements").Range("B2:FI" & index + 2).Value
BaseEPB = DateBase.Workbooks("data_base.xls").Worksheets("EPB").Range("B2:R" & index + 2).Value
BaseRemont = DateBase.Workbooks("data_base.xls").Worksheets("Remont").Range("B2:R" & index + 2).Value
UF1.Predpriyatie.Clear
For i = 1 To UBound(BasePredp, 1) - 1
    Me.Predpriyatie.AddItem BasePredp(i, 1)
Next i

If Date > CDate("24.06.2025") Then MsgBox ("Проверь срок действия удостоверений по НК")


UF1.DataAktPoRez.Value = Format(Date, "dd.mm.yyyy")
UF1.AktVIKD.Value = Format(Date - 1, "dd.mm.yyyy")
'ActiveDocument.Variables("DataAktPoRez").Value = Format(Date, "dd.mm.yyyy")
'ActiveDocument.Variables("DateAktPoRez").Value = FormDat(Format(Date, "dd.mm.yyyy"))
 

For Each ctl In UF1.Controls
    If TypeName(ctl) = "CheckBox" And Left(ctl.Name, 3) = "CBp" Then AllCBp.Add (Val(ctl.Caption))
    If TypeName(ctl) = "CheckBox" And Left(ctl.Name, 4) = "CBvb" Then AllCBv.Add (Val(ctl.Caption))
    If TypeName(ctl) = "CheckBox" And Left(ctl.Name, 4) = "CBho" Then AllCBh.Add (Val(ctl.Caption))
    If TypeName(ctl) = "CheckBox" And Left(ctl.Name, 4) = "CBtt" Then AllCBt.Add (Val(ctl.Caption))
    If TypeName(ctl) = "CheckBox" And Left(ctl.Name, 4) = "CBsn" Then AllCBs.Add (Val(ctl.Caption))
Next

Call QuickSort(AllCBp, 1, AllCBp.Count)
Call QuickSort(AllCBv, 1, AllCBv.Count)
Call QuickSort(AllCBh, 1, AllCBh.Count)
Call QuickSort(AllCBt, 1, AllCBt.Count)
Call QuickSort(AllCBs, 1, AllCBs.Count)

With UF1.ComboBoxTechUsrtvo
    .Clear
    .AddItem "котел"
    .AddItem "экономайзер"
    .AddItem "сосуд"
    .AddItem "трубопровод"
    .AddItem "воздухосборник"
    .AddItem "газификатор"
    .AddItem "подогреватель"
    .AddItem "автоклав"
    .AddItem "емкость"
    .AddItem "баллон"
    .AddItem "баллоны гр.уст."
    .AddItem "бак"
    .AddItem "резервуар"
    .AddItem "техн. трубопровод"
    .AddItem "теплообменник"
    .AddItem "насос"
    .AddItem "компрессор"
End With
With UF1.ComboBoxRaschet
    .Clear
    .AddItem "Пассат"
    .AddItem "РД 10-249-98"
    .AddItem "РД 153-34.1-37.525-96 кислота и щелочь"
    .AddItem "ГОСТ 32388-2013 тех. трубопровод"
    .AddItem "ГОСТ 25215-82 баллоны"
    .AddItem "Справочник Лащинский"
    .AddItem "ГОСТ 32106-2013 Вибрация"
End With

'Для классов присваиваем объекты с которыми они будут работать
ReDim CBpclass(1 To AllCBp.Count)
    'Присваиваем последовательно значениям массива значения объектов
    For i = 1 To AllCBp.Count
        Set CBpclass(i).CBp = UF1.MultiPage1.Pages("Page4").Controls("CBp" & AllCBp(i))
    Next i
ReDim CBvclass(1 To AllCBv.Count)
    For i = 1 To AllCBv.Count
        Set CBvclass(i).CBv = UF1.MultiPage1.Pages("Page4").Controls("CBvb" & AllCBv(i))
    Next i
ReDim CBhclass(1 To AllCBh.Count)
    For i = 1 To AllCBh.Count
        Set CBhclass(i).CBh = UF1.MultiPage1.Pages("Page4").Controls("CBho" & AllCBh(i))
    Next i
ReDim CBtclass(1 To AllCBt.Count)
    For i = 1 To AllCBt.Count
        Set CBtclass(i).CBt = UF1.MultiPage1.Pages("Page4").Controls("CBtt" & AllCBt(i))
    Next i
ReDim CBsclass(1 To AllCBs.Count)
    For i = 1 To AllCBs.Count
        Set CBsclass(i).CBs = UF1.MultiPage1.Pages("Page4").Controls("CBsn" & AllCBs(i))
    Next i


End Sub

Sub SetComboBox(SelectComboBox, tipFNP)

If tipFNP = "CBp" Then
    For Each mark In AllCBp
        UF1.Controls.Item("CBp" & mark).Value = False
    Next
End If
If tipFNP = "CBvb" Then
    For Each mark In AllCBv
        UF1.Controls.Item("CBvb" & mark).Value = False
    Next
End If
If tipFNP = "CBho" Then
    For Each mark In AllCBh
        UF1.Controls.Item("CBho" & mark).Value = False
    Next
End If
If tipFNP = "CBtt" Then
    For Each mark In AllCBt
        UF1.Controls.Item("CBtt" & mark).Value = False
    Next
End If
If tipFNP = "CBsn" Then
    For Each mark In AllCBs
        UF1.Controls.Item("CBsn" & mark).Value = False
    Next
End If

For Each mark In SelectComboBox
    UF1.Controls.Item(tipFNP & mark).Value = True
Next mark

End Sub

Private Sub UZK_Change()
If UF1.UZK.Value = True Then
    UF1.ZakUZK.Enabled = True
    UF1.ZakUZKD.Enabled = True
    UF1.UZK.Caption = "Есть"
Else
    UF1.ZakUZK.Enabled = False
    UF1.ZakUZKD.Enabled = False
    UF1.UZK.Caption = "Нет"
End If
End Sub

Private Sub VikMK_Change()
If UF1.VikMK.Value = True Then
    UF1.AktVIKMK.Enabled = True
    UF1.AktVIKMKD.Enabled = True
    UF1.VikMK.Caption = "Есть"
Else
    UF1.AktVIKMK.Enabled = False
    UF1.AktVIKMKD.Enabled = False
    UF1.VikMK.Caption = "Нет"
End If
End Sub

Private Sub VKorp_Change()
    BaseUstr(BaseUstrIndex, 19) = UF1.VKorp.Value
    ActiveDocument.Variables("VKorp").Value = ", V=" & UF1.VKorp.Value & " м" & Strings.ChrW(179)
    If UF1.VKorp.Value = "" Then ActiveDocument.Variables("VKorp").Value = Strings.ChrW(31)
End Sub

Private Sub Vladelez_Change()
    BaseUstr(BaseUstrIndex, 6) = UF1.Vladelez.Value
    ActiveDocument.Variables("Vladelez").Value = Trim(UF1.Vladelez.Value)
End Sub

Private Sub VRub_Change()
    BaseUstr(BaseUstrIndex, 24) = UF1.VRub.Value
    ActiveDocument.Variables("VRub").Value = ", V=" & UF1.VRub.Value & " м" & Strings.ChrW(179)
    If UF1.VRub.Value = "" Then ActiveDocument.Variables("VRub").Value = Strings.ChrW(31)
End Sub

Private Sub ZakMPDZD_Change()
ActiveDocument.Variables("ZakMPDZD").Value = UF1.ZakMPDZD.Value
If UF1.ZakMPDZD.Enabled = True Then
    UF1.ProtkTV.Value = Val(UF1.ZakMPDZD.Value) + 1
Else
    UF1.ProtkTV.Value = Val(UF1.ZakMPDZD.Value)
End If
End Sub

Private Sub ZakMPDZDD_Change()
If IsDate(UF1.ZakMPDZDD.Value) Then
    ActiveDocument.Variables("ZakMPDZDD").Value = UF1.ZakMPDZDD.Value
    ActiveDocument.Variables("ZakMPDZDData").Value = FormDat(UF1.ZakMPDZDD.Value)
End If
End Sub

Private Sub ZakUZK_Change()
ActiveDocument.Variables("ZakUZK").Value = UF1.ZakUZK.Value
If UF1.ZakUZK.Enabled = True Then
    UF1.ZakMPDZD.Value = Val(UF1.ZakUZK.Value) + 1
Else
    UF1.ZakMPDZD.Value = Val(UF1.ZakUZK.Value)
End If
End Sub

Private Sub ZakUZKD_Change()
If IsDate(UF1.ZakUZKD.Value) Then
    ActiveDocument.Variables("ZakUZKD").Value = UF1.ZakUZKD.Value
    ActiveDocument.Variables("ZakUZKData").Value = FormDat(UF1.ZakUZKD.Value)
End If
End Sub

Private Sub ZavKontr_Change()
    BaseUstr(BaseUstrIndex, 36) = UF1.ZavKontr.Value
    ActiveDocument.Variables("ZavKontr").Value = Trim(UF1.ZavKontr.Value)
End Sub

Private Sub ZavN_Change()
    ActiveDocument.Variables("ZavN").Value = Trim(UF1.ZavN.Value)
    If UF1.ZavN.ListIndex <> -1 Then UF1.RegN.ListIndex = UF1.ZavN.ListIndex
'    UF1.NazvTehUstr.ListIndex = UF1.ZavN.ListIndex
End Sub

Private Sub ZavN_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    BaseUstr(BaseUstrIndex, 7) = UF1.ZavN.Value
End Sub

Private Sub ZavodIzg_Change()
    BaseUstr(BaseUstrIndex, 10) = UF1.ZavodIzg.Value
    ActiveDocument.Variables("ZavodIzg").Value = Trim(UF1.ZavodIzg.Value)
End Sub

Sub UstSosAktNK(VikMK, Tolshin, UZK, MPDZD, Tverdost, Ovalnost, ProgibCb, KontrGibCh, PnIs)
        UF1.VikMK.Value = VikMK
        UF1.Tolshin.Value = Tolshin
        UF1.UZK.Value = UZK
        UF1.MPDZD.Value = MPDZD
        UF1.Tverdost.Value = Tverdost
        UF1.Ovalnost.Value = Ovalnost
        UF1.ProgibCb.Value = ProgibCb
        UF1.KontrGibCh.Value = KontrGibCh
        UF1.PnIs.Value = PnIs

End Sub
Sub FillOutFormUstr(index, ustroystvo)
    UF1.poleRegNum.Value = ustroystvo(index, 3)
    UF1.DataRegistracii.Value = ustroystvo(index, 5)
    UF1.Vladelez.Value = ustroystvo(index, 6)
    UF1.DataIzg.Value = ustroystvo(index, 8)
    UF1.DataVvoda.Value = ustroystvo(index, 9)
    UF1.ZavodIzg.Value = ustroystvo(index, 10)
    UF1.NazvTehUstr.Value = ustroystvo(index, 11)
    UF1.ComboBoxTipUstroistva.Value = ustroystvo(index, 12)
    UF1.ComboBoxTechUsrtvo.Value = ustroystvo(index, 13)
    UF1.NaznTehUstr.Value = ustroystvo(index, 14)
    UF1.CBPodNaliv.Value = ustroystvo(index, 15)
    UF1.CBVakuum.Value = ustroystvo(index, 16)
    UF1.RaschetnP.Value = ustroystvo(index, 17)
    UF1.Raschetnt.Value = ustroystvo(index, 18)
    UF1.VKorp.Value = ustroystvo(index, 19)
    UF1.RaschSreda.Value = ustroystvo(index, 20)
    UF1.CBRubashka.Value = ustroystvo(index, 21)
    UF1.RaschetnPRub.Value = ustroystvo(index, 22)
    UF1.RaschetntRub.Value = ustroystvo(index, 23)
    UF1.VRub.Value = ustroystvo(index, 24)
    UF1.RaschSredaRub.Value = ustroystvo(index, 25)
    UF1.RazreshaemoeP.Value = ustroystvo(index, 26)
    UF1.RabTemp.Value = ustroystvo(index, 27)
    UF1.RabSreda.Value = ustroystvo(index, 28)
    UF1.RabocheePRub.Value = ustroystvo(index, 29)
    UF1.RabTempRub.Value = ustroystvo(index, 30)
    UF1.RabSredaRub.Value = ustroystvo(index, 31)
    UF1.IspitatP.Value = ustroystvo(index, 32)
    UF1.IspitatPRub.Value = ustroystvo(index, 33)
    UF1.FlanzSoed.Value = ustroystvo(index, 34)
    UF1.PrimSvMat.Value = ustroystvo(index, 35)
    UF1.ZavKontr.Value = ustroystvo(index, 36)
    UF1.PasportKolStr.Value = ustroystvo(index, 37)
    UF1.TBUstrComment.Value = ustroystvo(index, 38)
    
End Sub

Sub FillOutElements(index, ustroystvo)
Z = 3
For i = 1 To 10
    For n = 1 To 8
        If n > 1 Then
            UF1.Controls.Item("R" & i & "C" & n - 1).Value = ustroystvo(index, Z + (8 * (UF1.SpButElm.Value)))
        Else
            If i > 1 Then UF1.Controls.Item("CBR" & i).Value = ustroystvo(index, Z + (8 * (UF1.SpButElm.Value)))
        End If
    Z = Z + 1
    Next n
Next i
UF1.CBZikl.Value = ustroystvo(index, 163)
If IsNull(UF1.CBZikl.Value) Then UF1.CBZikl.Value = False
UF1.KolZicl.Value = ustroystvo(index, 164)
End Sub

