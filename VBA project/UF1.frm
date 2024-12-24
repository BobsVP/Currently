VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF1 
   Caption         =   "Сосуд"
   ClientHeight    =   7440
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12240
   OleObjectBlob   =   "UF1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UF1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub AktAED_Change()
ActiveDocument.Variables("AktAED").Value = UF1.AktAED.Value
ActiveDocument.Variables("AktAEData").Value = FormDat(UF1.AktAED.Value)
End Sub

Private Sub AktVIKD_Exit(ByVal Cancel As MSForms.ReturnBoolean)
'Выставляем дату НК
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
Else
    MsgBox ("Неправильный формат даты")
End If

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
ActiveDocument.Variables("AktVIKMKD").Value = UF1.AktVIKMKD.Value
ActiveDocument.Variables("AktVIKMKData").Value = FormDat(UF1.AktVIKMKD.Value)
End Sub

Private Sub CBFNPOPVB_Change()
If UF1.CBFNPOPVB.Value = True Then
    ActiveDocument.Variables("UdostExpHim").Value = ActiveDocument.Variables("TckZpt").Value & " АЭ.21.01562.001"
    ActiveDocument.Variables("FNPOPVB").Value = UF1.FNPOPVBR.Value
    ActiveDocument.Variables("punkt1-1OPVB").Value = ActiveDocument.Variables("Enter").Value & "п.п." & ActiveDocument.Variables("p9OPVB").Value & ActiveDocument.Variables("p161OPVB").Value & ActiveDocument.Variables("p164OPVB").Value & ActiveDocument.Variables("p169OPVB").Value
    ActiveDocument.Variables("punkt1-1OPVB").Value = ActiveDocument.Variables("punkt1-1OPVB").Value & ActiveDocument.Variables("p177OPVB").Value & ActiveDocument.Variables("p178OPVB").Value & ActiveDocument.Variables("p179OPVB").Value & UF1.FNPOPVBR.Value
    If ActiveDocument.Variables("p9OPVB").Value <> Strings.ChrW(31) Then ActiveDocument.Variables("p7-1OPVBTechRegl").Value = " Эксплуатация " & ActiveDocument.Variables("TechUsrtva").Value & " осуществляется в соответствии с технологическим регламентом, что соответствует требованиям п. 9" & UF1.FNPOPVBR.Value & "."
    ActiveDocument.Variables("p7-1OPVB").Value = "; п." & ActiveDocument.Variables("p164OPVB").Value & UF1.FNPOPVBR.Value
    ActiveDocument.Variables("p7-1OPVBProdl").Value = ActiveDocument.Variables("TckZpt").Value & " п." & ActiveDocument.Variables("p161OPVB").Value & UF1.FNPOPVBR.Value
    ActiveDocument.Variables("p7-3OPVB").Value = ActiveDocument.Variables("TckZpt").Value & " п.п." & ActiveDocument.Variables("p177OPVB").Value & ActiveDocument.Variables("p178OPVB").Value & ActiveDocument.Variables("p179OPVB").Value & UF1.FNPOPVBR.Value
    ActiveDocument.Variables("GIFNPOPVB").Value = ActiveDocument.Variables("TckZpt").Value & " п." & ActiveDocument.Variables("p169OPVB").Value & UF1.FNPOPVBR.Value
    If ActiveDocument.Variables("NTDAktNKFNPORPD").Value <> Strings.ChrW(31) Then
        ActiveDocument.Variables("NTDAktNKFNPORPD").Value = ActiveDocument.Variables("NTDAktNKFNPORPD").Value & Strings.Chr(13) & "Федеральные нормы и правила в области промышленной безопасности " & Mid(UF1.FNPOPVBR.Value, 64, 122) & "."
    Else
        ActiveDocument.Variables("NTDAktNKFNPORPD").Value = "Федеральные нормы и правила в области промышленной безопасности " & Mid(UF1.FNPOPVBR.Value, 64, 122) & "."
    End If
Else
    ActiveDocument.Variables("UdostExpHim").Value = Strings.ChrW(31)
    ActiveDocument.Variables("FNPOPVB").Value = Strings.ChrW(31)
    ActiveDocument.Variables("punkt1-1OPVB").Value = Strings.ChrW(31)
    ActiveDocument.Variables("p7-1OPVBTechRegl").Value = Strings.ChrW(31)
    ActiveDocument.Variables("p7-1OPVB").Value = Strings.ChrW(31)
    ActiveDocument.Variables("p7-1OPVBProdl").Value = Strings.ChrW(31)
    ActiveDocument.Variables("p7-3OPVB").Value = Strings.ChrW(31)
    ActiveDocument.Variables("GIFNPOPVB").Value = Strings.ChrW(31)
End If
'    If UF1.CBFNPORPD.Value = False And UF1.CBFNPOPVB.Value = False Then MsgBox ("Нужно выбрать ФНП")
End Sub

Private Sub CBFNPORPD_Click()
If UF1.CBFNPORPD.Value = True Then
    ActiveDocument.Variables("TckZpt").Value = ";"
    ActiveDocument.Variables("Enter").Value = ";" & Strings.Chr(13)
    ActiveDocument.Variables("punkt1-1").Value = "п.п." 'Сборка пункта 1-1
    For Each mark In AllCBp
        If UF1.Controls.Item("CBp" & mark).Value = True Then ActiveDocument.Variables("punkt1-1").Value = ActiveDocument.Variables("punkt1-1").Value & ActiveDocument.Variables("CBp" & mark).Value
    Next
    ActiveDocument.Variables("punkt1-1").Value = Left(ActiveDocument.Variables("punkt1-1").Value, Len(ActiveDocument.Variables("punkt1-1").Value) - 1) & UF1.FNPORPDR.Value
    ActiveDocument.Variables("UdostExp").Value = "АЭ.20.01562.003" 'Удостоверение эксперта
    ActiveDocument.Variables("FNPORPDR").Value = UF1.FNPORPDR.Value
    ActiveDocument.Variables("p7-1ORPDProdl").Value = " п.п." & ActiveDocument.Variables("CBp2").Value & ActiveDocument.Variables("CBp3").Value & ActiveDocument.Variables("CBp394").Value & ActiveDocument.Variables("CBp465").Value
    ActiveDocument.Variables("p7-1ORPDProdl").Value = ActiveDocument.Variables("p7-1ORPDProdl").Value & ActiveDocument.Variables("CBp468").Value & ActiveDocument.Variables("CBp471").Value & UF1.FNPORPDR.Value
    ActiveDocument.Variables("p7-5ORPD").Value = "; п.п." & ActiveDocument.Variables("CBp394").Value & ActiveDocument.Variables("CBp465").Value & ActiveDocument.Variables("CBp468").Value & ActiveDocument.Variables("CBp471").Value & UF1.FNPORPDR.Value
    ActiveDocument.Variables("p12-1pril2").Value = " п. 12.1. Приложения №2" & UF1.FNPORPDR.Value
    ActiveDocument.Variables("p12-2pril2").Value = " п. 12.2. Приложения №2" & UF1.FNPORPDR.Value
    ActiveDocument.Variables("p12-3pril2").Value = " п. 12.3. Приложения №2" & UF1.FNPORPDR.Value
    ActiveDocument.Variables("p12-5pril2").Value = " п. 12.5. Приложения №2" & UF1.FNPORPDR.Value 'УЗК
    For Each mark In AllCBp 'Установка размещение и обвязка
        If mark > 9 And mark < 92 Then ActiveDocument.Variables("punkt7-3NTD").Value = ActiveDocument.Variables("punkt7-3NTD").Value & ActiveDocument.Variables("CBp" & mark).Value
    Next
    ActiveDocument.Variables("punkt7-3NTD").Value = ActiveDocument.Variables("punkt7-3NTD").Value & ActiveDocument.Variables("CBp538").Value & ActiveDocument.Variables("CBp539").Value & ActiveDocument.Variables("CBp577").Value & ActiveDocument.Variables("CBp589").Value
    ActiveDocument.Variables("punkt7-3NTD").Value = " п.п." & ActiveDocument.Variables("punkt7-3NTD").Value & UF1.FNPORPDR.Value
    For Each mark In AllCBp 'ГИ
        If mark > 174 And mark < 191 Then ActiveDocument.Variables("GIFNPORPD").Value = ActiveDocument.Variables("GIFNPORPD").Value & ActiveDocument.Variables("CBp" & mark).Value
    Next
    ActiveDocument.Variables("GIFNPORPD").Value = " п.п." & ActiveDocument.Variables("GIFNPORPD").Value & UF1.FNPORPDR.Value
    ActiveDocument.Variables("PIFNPORPD").Value = " п.п. 175, 190" & UF1.FNPORPDR.Value
    ActiveDocument.Variables("PIFNPORPD1").Value = " п. 190" & UF1.FNPORPDR.Value
    ActiveDocument.Variables("NTDAktNKFNPORPD").Value = UF1.FNPORPDRNK.Value
Else
    ActiveDocument.Variables("TckZpt").Value = Strings.ChrW(31)
    ActiveDocument.Variables("Enter").Value = Strings.ChrW(31)
    ActiveDocument.Variables("punkt1-1").Value = Strings.ChrW(31)
    ActiveDocument.Variables("UdostExp").Value = Strings.ChrW(31)
    ActiveDocument.Variables("FNPORPDR").Value = Strings.ChrW(31)
    ActiveDocument.Variables("p7-1ORPDProdl").Value = Strings.ChrW(31)
    ActiveDocument.Variables("p7-5ORPD").Value = Strings.ChrW(31)
    ActiveDocument.Variables("p12-1pril2").Value = Strings.ChrW(31)
    ActiveDocument.Variables("p12-2pril2").Value = Strings.ChrW(31)
    ActiveDocument.Variables("p12-3pril2").Value = Strings.ChrW(31)
    ActiveDocument.Variables("p12-5pril2").Value = Strings.ChrW(31)
    ActiveDocument.Variables("GIFNPORPD").Value = Strings.ChrW(31)
    ActiveDocument.Variables("PIFNPORPD").Value = Strings.ChrW(31)
    ActiveDocument.Variables("PIFNPORPD1").Value = Strings.ChrW(31)
    ActiveDocument.Variables("NTDAktNKFNPORPD").Value = Strings.ChrW(31)
End If
    If UF1.CBFNPOPVB.Value = True Then Call CBFNPOPVB_Change
'    If UF1.CBFNPORPD.Value = False And UF1.CBFNPOPVB.Value = False Then MsgBox ("Нужно выбрать ФНП")
End Sub

Private Sub CBGOST34347_Change()
If UF1.CBGOST34347.Value = True Then
    ActiveDocument.Variables.Item("GOST34347").Value = UF1.GOST34347.Value
    ActiveDocument.Variables("GOST34347PMat").Value = " п.п. 4.1.4, 5.9.1" & UF1.GOST34347.Value
    ActiveDocument.Variables("GOST34347PSosEl").Value = ActiveDocument.Variables("TckZpt").Value & " п. 5.10.2" & UF1.GOST34347.Value
    ActiveDocument.Variables("GOST34347PGiI").Value = "; п.п. 7.11.3, 7.11.5, 7.11.10" & UF1.GOST34347.Value
    ActiveDocument.Variables.Item("GOST34347PPiI").Value = "; п. 7.11.9" & UF1.GOST34347.Value
    If ActiveDocument.Variables("NTDAktNKFNPORPD").Value <> Strings.ChrW(31) Then
        ActiveDocument.Variables("NTDAktNKFNPORPD").Value = ActiveDocument.Variables("NTDAktNKFNPORPD").Value & Strings.Chr(13) & Mid(UF1.GOST34347.Value, 2, 79) & "."
    Else
        ActiveDocument.Variables("NTDAktNKFNPORPD").Value = Mid(UF1.GOST34347.Value, 2, 79) & "."
    End If
Else
    ActiveDocument.Variables.Item("GOST34347").Value = Strings.ChrW(31)
    ActiveDocument.Variables("GOST34347PSosEl").Value = Strings.ChrW(31)
    ActiveDocument.Variables("GOST34347PMat").Value = Strings.ChrW(31)
    ActiveDocument.Variables("GOST34347PGiI").Value = Strings.ChrW(31)
    ActiveDocument.Variables.Item("GOST34347PPiI").Value = Strings.ChrW(31)
End If
End Sub

Private Sub CBp10_Click()
If UF1.CBp10.Value = True Then ActiveDocument.Variables("CBp10").Value = " 10,"
If UF1.CBp10.Value = False Then ActiveDocument.Variables("CBp10").Value = Strings.ChrW(31)
End Sub

Private Sub CBp22_Click()
If UF1.CBp22.Value = True Then ActiveDocument.Variables("CBp22").Value = " 22,"
If UF1.CBp22.Value = False Then ActiveDocument.Variables("CBp22").Value = Strings.ChrW(31)
End Sub

Private Sub CBp100_Click()
If UF1.CBp100.Value = True Then
ActiveDocument.Variables("CBp100").Value = " 100,"
ActiveDocument.Variables("p100FNPORPD").Value = " п. 100" & UF1.FNPORPDR.Value
End If
If UF1.CBp100.Value = False Then
ActiveDocument.Variables("CBp100").Value = Strings.ChrW(31)
ActiveDocument.Variables("p100FNPORPD").Value = Strings.ChrW(31)
End If
End Sub

Private Sub CBp154_Click()
If UF1.CBp154.Value = True Then ActiveDocument.Variables("CBp154").Value = " 154,"
If UF1.CBp154.Value = False Then ActiveDocument.Variables("CBp154").Value = Strings.ChrW(31)
End Sub

Private Sub CBp156_Click()
If UF1.CBp156.Value = True Then ActiveDocument.Variables("CBp156").Value = " 156,"
If UF1.CBp156.Value = False Then ActiveDocument.Variables("CBp156").Value = Strings.ChrW(31)
End Sub

Private Sub CBp175_Click()
If UF1.CBp175.Value = True Then ActiveDocument.Variables("CBp175").Value = " 175,"
If UF1.CBp175.Value = False Then ActiveDocument.Variables("CBp175").Value = Strings.ChrW(31)
End Sub

Private Sub CBp177_Click()
If UF1.CBp177.Value = True Then
    ActiveDocument.Variables("CBp177").Value = " 177,"
    If UF1.PnIs.Value = True Then UF1.PnIs.Value = False
Else
    ActiveDocument.Variables("CBp177").Value = Strings.ChrW(31)
End If
End Sub

Private Sub CBp178_Click()
If UF1.CBp178.Value = True Then
    ActiveDocument.Variables("CBp178").Value = " 178,"
    If UF1.PnIs.Value = True Then UF1.PnIs.Value = False
Else
    ActiveDocument.Variables("CBp178").Value = Strings.ChrW(31)
End If
End Sub

Private Sub CBp179_Click()
If UF1.CBp179.Value = True Then
    ActiveDocument.Variables("CBp179").Value = " 179,"
    If UF1.PnIs.Value = True Then UF1.PnIs.Value = False
Else
    ActiveDocument.Variables("CBp179").Value = Strings.ChrW(31)
End If
End Sub

Private Sub CBp184_Click()
If UF1.CBp184.Value = True Then ActiveDocument.Variables("CBp184").Value = " 184,"
If UF1.CBp184.Value = False Then ActiveDocument.Variables("CBp184").Value = Strings.ChrW(31)
End Sub

Private Sub CBp185_Click()
If UF1.CBp185.Value = True Then
    ActiveDocument.Variables("CBp185").Value = " 185,"
    If UF1.PnIs.Value = True Then UF1.PnIs.Value = False
Else
    ActiveDocument.Variables("CBp185").Value = Strings.ChrW(31)
End If
End Sub

Private Sub CBp186_Click()
If UF1.CBp186.Value = True Then
    ActiveDocument.Variables("CBp186").Value = " 186,"
    If UF1.PnIs.Value = True Then UF1.PnIs.Value = False
Else
    ActiveDocument.Variables("CBp186").Value = Strings.ChrW(31)
End If
End Sub

Private Sub CBp187_Click()
If UF1.CBp187.Value = True Then
    ActiveDocument.Variables("CBp187").Value = " 187,"
    If UF1.PnIs.Value = True Then UF1.PnIs.Value = False
Else
    ActiveDocument.Variables("CBp187").Value = Strings.ChrW(31)
End If
End Sub

Private Sub CBp188_Click()
If UF1.CBp188.Value = True Then
    ActiveDocument.Variables("CBp188").Value = " 188,"
    If UF1.PnIs.Value = True Then UF1.PnIs.Value = False
Else
    ActiveDocument.Variables("CBp188").Value = Strings.ChrW(31)
End If
End Sub

Private Sub CBp190_Click()
If UF1.CBp190.Value = True Then ActiveDocument.Variables("CBp190").Value = " 190,"
If UF1.CBp190.Value = False Then ActiveDocument.Variables("CBp190").Value = Strings.ChrW(31)
End Sub

Private Sub CBp2_Click()
If UF1.CBp2.Value = True Then ActiveDocument.Variables("CBp2").Value = " 2,"
If UF1.CBp2.Value = False Then ActiveDocument.Variables("CBp2").Value = Strings.ChrW(31)
End Sub

Private Sub CBp246_Click()
If UF1.CBp246.Value = True Then ActiveDocument.Variables("CBp246").Value = " 246,"
If UF1.CBp246.Value = False Then ActiveDocument.Variables("CBp246").Value = Strings.ChrW(31)
End Sub

Private Sub CBp257_Click()
If UF1.CBp257.Value = True Then ActiveDocument.Variables("CBp257").Value = " 257,"
If UF1.CBp257.Value = False Then ActiveDocument.Variables("CBp257").Value = Strings.ChrW(31)
End Sub

Private Sub CBp258_Click()
If UF1.CBp258.Value = True Then ActiveDocument.Variables("CBp258").Value = " 258,"
If UF1.CBp258.Value = False Then ActiveDocument.Variables("CBp258").Value = Strings.ChrW(31)
End Sub

Private Sub CBp260_Click()
If UF1.CBp260.Value = True Then ActiveDocument.Variables("CBp260").Value = " 260,"
If UF1.CBp260.Value = False Then ActiveDocument.Variables("CBp260").Value = Strings.ChrW(31)
End Sub

Private Sub CBp267_Click()
If UF1.CBp267.Value = True Then ActiveDocument.Variables("CBp267").Value = " 267,"
If UF1.CBp267.Value = False Then ActiveDocument.Variables("CBp267").Value = Strings.ChrW(31)
End Sub

Private Sub CBp268_Click()
If UF1.CBp268.Value = True Then ActiveDocument.Variables("CBp268").Value = " 268,"
If UF1.CBp268.Value = False Then ActiveDocument.Variables("CBp268").Value = Strings.ChrW(31)
End Sub

Private Sub CBp269_Click()
If UF1.CBp269.Value = True Then ActiveDocument.Variables("CBp269").Value = " 269,"
If UF1.CBp269.Value = False Then ActiveDocument.Variables("CBp269").Value = Strings.ChrW(31)
End Sub

Private Sub CBp270_Click()
If UF1.CBp270.Value = True Then ActiveDocument.Variables("CBp270").Value = " 270,"
If UF1.CBp270.Value = False Then ActiveDocument.Variables("CBp270").Value = Strings.ChrW(31)
End Sub

Private Sub CBp271_Click()
If UF1.CBp271.Value = True Then ActiveDocument.Variables("CBp271").Value = " 271,"
If UF1.CBp271.Value = False Then ActiveDocument.Variables("CBp271").Value = Strings.ChrW(31)
End Sub

Private Sub CBp3_Click()
If UF1.CBp3.Value = True Then ActiveDocument.Variables("CBp3").Value = " 3,"
If UF1.CBp3.Value = False Then ActiveDocument.Variables("CBp3").Value = Strings.ChrW(31)
End Sub

Private Sub CBp338_Click()
If UF1.CBp338.Value = True Then ActiveDocument.Variables("CBp338").Value = " 338,"
If UF1.CBp338.Value = False Then ActiveDocument.Variables("CBp338").Value = Strings.ChrW(31)
End Sub

Private Sub CBp339_Click()
If UF1.CBp339.Value = True Then ActiveDocument.Variables("CBp339").Value = " 339,"
If UF1.CBp339.Value = False Then ActiveDocument.Variables("CBp339").Value = Strings.ChrW(31)
End Sub

Private Sub CBp340_Click()
If UF1.CBp340.Value = True Then ActiveDocument.Variables("CBp340").Value = " 340,"
If UF1.CBp340.Value = False Then ActiveDocument.Variables("CBp340").Value = Strings.ChrW(31)
End Sub

Private Sub CBp341_Click()
If UF1.CBp341.Value = True Then ActiveDocument.Variables("CBp341").Value = " 341,"
If UF1.CBp341.Value = False Then ActiveDocument.Variables("CBp341").Value = Strings.ChrW(31)
End Sub

Private Sub CBp343_Click()
If UF1.CBp343.Value = True Then ActiveDocument.Variables("CBp343").Value = " 343,"
If UF1.CBp343.Value = False Then ActiveDocument.Variables("CBp343").Value = Strings.ChrW(31)
End Sub

Private Sub CBp348_Click()
If UF1.CBp348.Value = True Then
ActiveDocument.Variables("CBp348").Value = " 348,"
ActiveDocument.Variables("PredKlNet").Value = " Предохранительный клапан на сосуде не установлен, его установка не обязательна, так как рабочее давление в сосуде больше давления питающего источника, что соответствует требованиям п. 348" & UF1.FNPORPDR.Value & "."
UF1.CBp338.Value = False
UF1.CBp339.Value = False
UF1.CBp340.Value = False
UF1.CBp341.Value = False
UF1.CBp343.Value = False
UF1.CBp353.Value = False
End If
If UF1.CBp348.Value = False Then
ActiveDocument.Variables("CBp348").Value = Strings.ChrW(31)
ActiveDocument.Variables("PredKlNet").Value = Strings.ChrW(31)
End If
End Sub

Private Sub CBp353_Click()
If UF1.CBp353.Value = True Then ActiveDocument.Variables("CBp353").Value = " 353,"
If UF1.CBp353.Value = False Then ActiveDocument.Variables("CBp353").Value = Strings.ChrW(31)
End Sub

Private Sub CBp372_Click()
If UF1.CBp372.Value = True Then ActiveDocument.Variables("CBp372").Value = " 372,"
If UF1.CBp372.Value = False Then ActiveDocument.Variables("CBp372").Value = Strings.ChrW(31)
End Sub

Private Sub CBp373_Click()
If UF1.CBp373.Value = True Then ActiveDocument.Variables("CBp373").Value = " 373,"
If UF1.CBp373.Value = False Then ActiveDocument.Variables("CBp373").Value = Strings.ChrW(31)
End Sub

Private Sub CBp374_Click()
If UF1.CBp374.Value = True Then ActiveDocument.Variables("CBp374").Value = " 374,"
If UF1.CBp374.Value = False Then ActiveDocument.Variables("CBp374").Value = Strings.ChrW(31)
End Sub

Private Sub CBp378_Click()
If UF1.CBp378.Value = True Then ActiveDocument.Variables("CBp378").Value = " 378,"
If UF1.CBp378.Value = False Then ActiveDocument.Variables("CBp378").Value = Strings.ChrW(31)
End Sub

Private Sub CBp379_Click()
If UF1.CBp379.Value = True Then ActiveDocument.Variables("CBp379").Value = " 379,"
If UF1.CBp379.Value = False Then ActiveDocument.Variables("CBp379").Value = Strings.ChrW(31)
End Sub

Private Sub CBp38_Click()
If UF1.CBp38.Value = True Then ActiveDocument.Variables("CBp38").Value = " 38,"
If UF1.CBp38.Value = False Then ActiveDocument.Variables("CBp38").Value = Strings.ChrW(31)
End Sub

Private Sub CBp39_Click()
If UF1.CBp39.Value = True Then ActiveDocument.Variables("CBp39").Value = " 39,"
If UF1.CBp39.Value = False Then ActiveDocument.Variables("CBp39").Value = Strings.ChrW(31)
End Sub

Private Sub CBp394_Click()
If UF1.CBp394.Value = True Then ActiveDocument.Variables("CBp394").Value = " 394,"
If UF1.CBp394.Value = False Then ActiveDocument.Variables("CBp394").Value = Strings.ChrW(31)
End Sub

Private Sub CBp43_Click()
If UF1.CBp43.Value = True Then ActiveDocument.Variables("CBp43").Value = " 43,"
If UF1.CBp43.Value = False Then ActiveDocument.Variables("CBp43").Value = Strings.ChrW(31)
End Sub

Private Sub CBp45_Click()
If UF1.CBp45.Value = True Then ActiveDocument.Variables("CBp45").Value = " 45,"
If UF1.CBp45.Value = False Then ActiveDocument.Variables("CBp45").Value = Strings.ChrW(31)
End Sub

Private Sub CBp46_Click()
If UF1.CBp46.Value = True Then ActiveDocument.Variables("CBp46").Value = " 46,"
If UF1.CBp46.Value = False Then ActiveDocument.Variables("CBp46").Value = Strings.ChrW(31)
End Sub

Private Sub CBp465_Click()
If UF1.CBp465.Value = True Then ActiveDocument.Variables("CBp465").Value = " 465,"
If UF1.CBp465.Value = False Then ActiveDocument.Variables("CBp465").Value = Strings.ChrW(31)
End Sub

Private Sub CBp466_Click()
If UF1.CBp466.Value = True Then
    ActiveDocument.Variables("CBp466").Value = " 466,"
    ActiveDocument.Variables("CBp466-2").Value = " в соответствии с п. 466" & UF1.FNPORPDR.Value
    ActiveDocument.Variables("CBp466-1").Value = " и п. 466" & UF1.FNPORPDR.Value
    ActiveDocument.Variables("CBp466data").Value = "30.09." & Year(Date)
End If
If UF1.CBp466.Value = False Then ActiveDocument.Variables("CBp466").Value = Strings.ChrW(31)
If UF1.CBp466.Value = False Then ActiveDocument.Variables("CBp466-1").Value = Strings.ChrW(31)
End Sub

Private Sub CBp468_Click()
If UF1.CBp468.Value = True Then ActiveDocument.Variables("CBp468").Value = " 468,"
If UF1.CBp468.Value = False Then ActiveDocument.Variables("CBp468").Value = Strings.ChrW(31)
End Sub

Private Sub CBp469_Click()
If UF1.CBp469.Value = True Then ActiveDocument.Variables("CBp469").Value = " 469,"
If UF1.CBp469.Value = False Then ActiveDocument.Variables("CBp469").Value = Strings.ChrW(31)
End Sub

Private Sub CBp47_Click()
If UF1.CBp47.Value = True Then ActiveDocument.Variables("CBp47").Value = " 47,"
If UF1.CBp47.Value = False Then ActiveDocument.Variables("CBp47").Value = Strings.ChrW(31)
End Sub

Private Sub CBp471_Click()
If UF1.CBp471.Value = True Then ActiveDocument.Variables("CBp471").Value = " 471,"
If UF1.CBp471.Value = False Then ActiveDocument.Variables("CBp471").Value = Strings.ChrW(31)
End Sub

Private Sub CBp49_Click()
If UF1.CBp49.Value = True Then ActiveDocument.Variables("CBp49").Value = " 49,"
If UF1.CBp49.Value = False Then ActiveDocument.Variables("CBp49").Value = Strings.ChrW(31)
End Sub

Private Sub CBp50_Click()
If UF1.CBp50.Value = True Then ActiveDocument.Variables("CBp50").Value = " 50,"
If UF1.CBp50.Value = False Then ActiveDocument.Variables("CBp50").Value = Strings.ChrW(31)
End Sub

Private Sub CBp500_Click()
If UF1.CBp500.Value = True Then ActiveDocument.Variables("CBp500").Value = " 500,"
If UF1.CBp500.Value = False Then ActiveDocument.Variables("CBp500").Value = Strings.ChrW(31)
End Sub

Private Sub CBp502_Click()
If UF1.CBp502.Value = True Then ActiveDocument.Variables("CBp502").Value = " 502,"
If UF1.CBp502.Value = False Then ActiveDocument.Variables("CBp502").Value = Strings.ChrW(31)
End Sub

Private Sub CBp503_Click()
If UF1.CBp503.Value = True Then ActiveDocument.Variables("CBp503").Value = " 503,"
If UF1.CBp503.Value = False Then ActiveDocument.Variables("CBp503").Value = Strings.ChrW(31)
End Sub

Private Sub CBp505_Click()
If UF1.CBp505.Value = True Then ActiveDocument.Variables("CBp505").Value = " 505,"
If UF1.CBp505.Value = False Then ActiveDocument.Variables("CBp505").Value = Strings.ChrW(31)
End Sub

Private Sub CBp506_Click()
If UF1.CBp506.Value = True Then ActiveDocument.Variables("CBp506").Value = " 506,"
If UF1.CBp506.Value = False Then ActiveDocument.Variables("CBp506").Value = Strings.ChrW(31)
End Sub

Private Sub CBp538_Click()
If UF1.CBp538.Value = True Then ActiveDocument.Variables("CBp538").Value = " 538,"
If UF1.CBp538.Value = False Then ActiveDocument.Variables("CBp538").Value = Strings.ChrW(31)
End Sub

Private Sub CBp539_Click()
If UF1.CBp539.Value = True Then ActiveDocument.Variables("CBp539").Value = " 539,"
If UF1.CBp539.Value = False Then ActiveDocument.Variables("CBp539").Value = Strings.ChrW(31)
End Sub

Private Sub CBp540_Click()
If UF1.CBp540.Value = True Then ActiveDocument.Variables("CBp540").Value = " 540,"
If UF1.CBp540.Value = False Then ActiveDocument.Variables("CBp540").Value = Strings.ChrW(31)
End Sub

Private Sub CBp577_Click()
If UF1.CBp577.Value = True Then ActiveDocument.Variables("CBp577").Value = " 577,"
If UF1.CBp577.Value = False Then ActiveDocument.Variables("CBp577").Value = Strings.ChrW(31)
End Sub

Private Sub CBp589_Click()
If UF1.CBp589.Value = True Then ActiveDocument.Variables("CBp589").Value = " 589,"
If UF1.CBp589.Value = False Then ActiveDocument.Variables("CBp589").Value = Strings.ChrW(31)
End Sub

Private Sub CBp61_Click()
If UF1.CBp61.Value = True Then ActiveDocument.Variables("CBp61").Value = " 61,"
If UF1.CBp61.Value = False Then ActiveDocument.Variables("CBp61").Value = Strings.ChrW(31)
End Sub

Private Sub CBp64_Click()
If UF1.CBp64.Value = True Then ActiveDocument.Variables("CBp64").Value = " 64,"
If UF1.CBp64.Value = False Then ActiveDocument.Variables("CBp64").Value = Strings.ChrW(31)
End Sub

Private Sub CBp65_Click()
If UF1.CBp65.Value = True Then ActiveDocument.Variables("CBp65").Value = " 65,"
If UF1.CBp65.Value = False Then ActiveDocument.Variables("CBp65").Value = Strings.ChrW(31)
End Sub

Private Sub CBp68_Click()
If UF1.CBp68.Value = True Then ActiveDocument.Variables("CBp68").Value = " 68,"
If UF1.CBp68.Value = False Then ActiveDocument.Variables("CBp68").Value = Strings.ChrW(31)
End Sub

Private Sub CBp69_Click()
If UF1.CBp69.Value = True Then ActiveDocument.Variables("CBp69").Value = " 69,"
If UF1.CBp69.Value = False Then ActiveDocument.Variables("CBp69").Value = Strings.ChrW(31)
End Sub

Private Sub CBp71_Click()
If UF1.CBp71.Value = True Then ActiveDocument.Variables("CBp71").Value = " 71,"
If UF1.CBp71.Value = False Then ActiveDocument.Variables("CBp71").Value = Strings.ChrW(31)
End Sub

Private Sub CBp80_Click()
If UF1.CBp80.Value = True Then ActiveDocument.Variables("CBp80").Value = " 80,"
If UF1.CBp80.Value = False Then ActiveDocument.Variables("CBp80").Value = Strings.ChrW(31)
End Sub

Private Sub CBp81_Click()
If UF1.CBp81.Value = True Then ActiveDocument.Variables("CBp81").Value = " 81,"
If UF1.CBp81.Value = False Then ActiveDocument.Variables("CBp81").Value = Strings.ChrW(31)
End Sub

Private Sub CBp85_Click()
If UF1.CBp85.Value = True Then ActiveDocument.Variables("CBp85").Value = " 85,"
If UF1.CBp85.Value = False Then ActiveDocument.Variables("CBp85").Value = Strings.ChrW(31)
End Sub

Private Sub CBp86_Click()
If UF1.CBp86.Value = True Then ActiveDocument.Variables("CBp86").Value = " 86,"
If UF1.CBp86.Value = False Then ActiveDocument.Variables("CBp86").Value = Strings.ChrW(31)
End Sub

Private Sub CBp90_Click()
If UF1.CBp90.Value = True Then ActiveDocument.Variables("CBp90").Value = " 90,"
If UF1.CBp90.Value = False Then ActiveDocument.Variables("CBp90").Value = Strings.ChrW(31)
End Sub

Private Sub CBp91_Click()
If UF1.CBp91.Value = True Then ActiveDocument.Variables("CBp91").Value = " 91,"
If UF1.CBp91.Value = False Then ActiveDocument.Variables("CBp91").Value = Strings.ChrW(31)
End Sub

Private Sub CBRD2626012_Change()
If UF1.CBRD2626012.Value = True Then
    ActiveDocument.Variables("ovalnrd").Value = ActiveDocument.Variables("ovalnrd").Value & "; п. 5.4.3.2" & UF1.RD26_260.Value
    If ActiveDocument.Variables("NTDAktNKFNPORPD").Value <> Strings.ChrW(31) Then
         ActiveDocument.Variables("NTDAktNKFNPORPD").Value = ActiveDocument.Variables("NTDAktNKFNPORPD").Value & Strings.Chr(13) & Mid(UF1.RD26_260.Value, 2, 104) & "."
    Else
        ActiveDocument.Variables("NTDAktNKFNPORPD").Value = Mid(UF1.RD26_260.Value, 2, 104) & "."
    End If
Else
'    ActiveDocument.Variables("ovalnrd").Value = Strings.ChrW(31)
End If
End Sub

Private Sub CBRUA93_Click()
If UF1.CBRUA93.Value = True Then
    ActiveDocument.Variables("p4-5RUA").Value = "; п.п. 2.54, 2.55" & UF1.RUA93.Value
    ActiveDocument.Variables("GIRUA93").Value = "; п.п." & ActiveDocument.Variables("p2-102RUA").Value & ActiveDocument.Variables("p2-111RUA").Value & " 2.113" & UF1.RUA93.Value
    If ActiveDocument.Variables("NTDAktNKFNPORPD").Value <> Strings.ChrW(31) Then
        ActiveDocument.Variables("NTDAktNKFNPORPD").Value = ActiveDocument.Variables("NTDAktNKFNPORPD").Value & Strings.Chr(13) & Mid(UF1.RUA93.Value, 2, 140) & "."
    Else
        ActiveDocument.Variables("NTDAktNKFNPORPD").Value = "Федеральные нормы и правила в области промышленной безопасности " & Mid(UF1.RUA93.Value, 2, 140) & "."
    End If
Else
    ActiveDocument.Variables("p4-5RUA").Value = Strings.ChrW(31)
    ActiveDocument.Variables("GIRUA93").Value = Strings.ChrW(31)
End If
End Sub

Private Sub CBSO439_Change()
If UF1.CBSO439.Value = True Then
    ActiveDocument.Variables("VIKrdSO439").Value = "; п.п. 5.4, 5.7, 5.8, 5.10" & UF1.SO439.Value
    ActiveDocument.Variables("tverdSO439").Value = " п. 5.9" & UF1.SO439.Value
    ActiveDocument.Variables("ovalnrd").Value = "При измерениях овальности обечайки " & ActiveDocument.Variables("TechUsrtva").Value & " установлено, что овальность соответствует требованиям п. 5.6." & UF1.SO439.Value
    If ActiveDocument.Variables("NTDAktNKFNPORPD").Value <> Strings.ChrW(31) Then
        ActiveDocument.Variables("NTDAktNKFNPORPD").Value = ActiveDocument.Variables("NTDAktNKFNPORPD").Value & Strings.Chr(13) & Mid(UF1.SO439.Value, 2, 94) & "."
    Else
        ActiveDocument.Variables("NTDAktNKFNPORPD").Value = Mid(UF1.SO439.Value, 2, 94) & "."
    End If
    ActiveDocument.Variables("NTDAktNKPD").Value = Mid(UF1.SO439.Value, 2, 94) & "."
Else
    ActiveDocument.Variables("VIKrdSO439").Value = Strings.ChrW(31)
    ActiveDocument.Variables("tverdSO439").Value = Strings.ChrW(31)
    ActiveDocument.Variables("ovalnrd").Value = Strings.ChrW(31)
End If
End Sub

Private Sub CBSO464_Change()
If UF1.CBSO464.Value = True Then
    ActiveDocument.Variables("VIKrdSO464").Value = "; п.п. 5.1, 5.2, 5.3, 5.6, 5.10, 5.12, 5.13, 5.15, 5.18, 5.20" & UF1.SO464.Value
    ActiveDocument.Variables("TolSte464").Value = "; п. 5.7" & UF1.SO464.Value
    ActiveDocument.Variables("UZKSO469").Value = "; п. 5.19" & UF1.SO464.Value
    ActiveDocument.Variables("tverdSO464").Value = " п. 5.14" & UF1.SO464.Value
    ActiveDocument.Variables("GISO464").Value = "; раздела 4.8" & UF1.SO464.Value
    If ActiveDocument.Variables("NTDAktNKFNPORPD").Value <> Strings.ChrW(31) Then
        ActiveDocument.Variables("NTDAktNKFNPORPD").Value = ActiveDocument.Variables("NTDAktNKFNPORPD").Value & Strings.Chr(13) & Mid(UF1.SO464.Value, 2, 97) & "."
    Else
        ActiveDocument.Variables("NTDAktNKFNPORPD").Value = Mid(UF1.SO464.Value, 2, 97) & "."
    End If
    ActiveDocument.Variables("NTDAktNKPD").Value = Mid(UF1.SO464.Value, 2, 97) & "."
Else
    ActiveDocument.Variables("VIKrdSO464").Value = Strings.ChrW(31)
    ActiveDocument.Variables("TolSte464").Value = Strings.ChrW(31)
    ActiveDocument.Variables("UZKSO469").Value = Strings.ChrW(31)
    ActiveDocument.Variables("tverdSO464").Value = Strings.ChrW(31)
    ActiveDocument.Variables("GISO464").Value = Strings.ChrW(31)
End If
End Sub

Private Sub CBSO469_Click()
If UF1.CBSO469.Value = True Then
    ActiveDocument.Variables("VIKrdSO469").Value = "; п.п. 5.4, 5.5, 5.15" & UF1.SO469.Value
    ActiveDocument.Variables("tverdSO469").Value = " п. 5.29" & UF1.SO469.Value
    ActiveDocument.Variables("ovalnKG").Value = " п. 5.11" & UF1.SO469.Value
    ActiveDocument.Variables("SOrd").Value = " п. 5.13" & UF1.SO469.Value
    ActiveDocument.Variables("ovalnrd").Value = "При измерениях овальности барабанов котла установлено, что овальность соответствует требованиям п. 5.10." & UF1.SO469.Value
    If ActiveDocument.Variables("NTDAktNKFNPORPD").Value <> Strings.ChrW(31) Then
        ActiveDocument.Variables("NTDAktNKFNPORPD").Value = ActiveDocument.Variables("NTDAktNKFNPORPD").Value & Strings.Chr(13) & Mid(UF1.SO469.Value, 2, 188) & "."
    Else
        ActiveDocument.Variables("NTDAktNKFNPORPD").Value = Mid(UF1.SO469.Value, 2, 188) & "."
    End If
    ActiveDocument.Variables("NTDAktNKPD").Value = Mid(UF1.SO469.Value, 2, 188) & "."
Else
    ActiveDocument.Variables("VIKrdSO469").Value = Strings.ChrW(31)
    ActiveDocument.Variables("tverdSO469").Value = Strings.ChrW(31)
    ActiveDocument.Variables("ovalnKG").Value = Strings.ChrW(31)
    ActiveDocument.Variables("SOrd").Value = Strings.ChrW(31)
    ActiveDocument.Variables("ovalnrd").Value = Strings.ChrW(31)
End If
End Sub

Private Sub CBVM030104_Change()
If UF1.CBVM030104.Value = True Then
    ActiveDocument.Variables("p4-5VM").Value = "; п. 4.5" & UF1.RDVM03.Value
    ActiveDocument.Variables("p4-7VM").Value = "; п. 4.7" & UF1.RDVM03.Value
    If ActiveDocument.Variables("NTDAktNKFNPORPD").Value <> Strings.ChrW(31) Then
        ActiveDocument.Variables("NTDAktNKFNPORPD").Value = ActiveDocument.Variables("NTDAktNKFNPORPD").Value & Strings.Chr(13) & Mid(UF1.RDVM03.Value, 2, 229) & "."
    Else
        ActiveDocument.Variables("NTDAktNKFNPORPD").Value = Mid(UF1.RDVM03.Value, 2, 229) & "."
    End If
Else
    ActiveDocument.Variables("p4-5VM").Value = Strings.ChrW(31)
    ActiveDocument.Variables("p4-7VM").Value = Strings.ChrW(31)
End If
End Sub

Private Sub ddiam_Change()
UF1.ddiam.Value = Trim(UF1.ddiam.Value)
ActiveDocument.Variables("ddiam").Value = UF1.ddiam.Value
End Sub

Private Sub dp15_Change()
ActiveDocument.Variables("pribor15").Value = UF1.pribor15.Value & UF1.dp15.Value
End Sub

Private Sub dp16_Change()
ActiveDocument.Variables("pribor16").Value = UF1.pribor16.Value & UF1.dp16.Value
End Sub

Private Sub dtolsh_Change()
UF1.dtolsh.Value = Trim(UF1.dtolsh.Value)
ActiveDocument.Variables("dtolsh").Value = UF1.dtolsh.Value
End Sub

Private Sub KartOvalnD_Change()
ActiveDocument.Variables("KartOvalnD").Value = UF1.KartOvalnD.Value
ActiveDocument.Variables("KartOvalnData").Value = FormDat(UF1.KartOvalnD.Value)
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
ActiveDocument.Variables("KontrGibD").Value = UF1.KontrGibD.Value
ActiveDocument.Variables("KontrGibData").Value = FormDat(UF1.KontrGibD.Value)
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

Private Sub poleRegNum_Change()
ActiveDocument.Variables("No").Value = UF1.poleRegNum.Value & "№"
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
ActiveDocument.Variables("ProgibD").Value = UF1.ProgibD.Value
ActiveDocument.Variables("ProgibData").Value = FormDat(UF1.ProgibD.Value)
End Sub

Private Sub ProtkTVD_Change()
ActiveDocument.Variables("ProtkTVD").Value = UF1.ProtkTVD.Value
ActiveDocument.Variables("ProtkTVData").Value = FormDat(UF1.ProtkTVD.Value)
End Sub

Private Sub ProtokTolchD_Change()
ActiveDocument.Variables("ProtokTolchD").Value = UF1.ProtokTolchD.Value
ActiveDocument.Variables("ProtokTolchData").Value = FormDat(UF1.ProtokTolchD.Value)
End Sub

Private Sub AktGID_Change()
If IsDate(UF1.AktGID.Value) Then
    ActiveDocument.Variables("DoNgoda").Value = Format(DateAdd("yyyy", Val(UF1.NaNLet.Value), UF1.AktGID.Value), "dd.mm.yyyy")
    ActiveDocument.Variables("AktGID").Value = UF1.AktGID.Value
    ActiveDocument.Variables("AktGIData").Value = FormDat(UF1.AktGID.Value)
    UF1.AktAED.Value = UF1.AktGID.Value
End If
End Sub

Private Sub AktAE_Change()
ActiveDocument.Variables("AktAE").Value = UF1.AktAE.Value
End Sub

Private Sub AktGI_Change()
ActiveDocument.Variables("AktGI").Value = UF1.AktGI.Value
If UF1.AktGI.Enabled = True Then
    UF1.AktAE.Value = Val(UF1.AktGI.Value) + 1
Else
    UF1.AktAE.Value = Val(UF1.AktGI.Value)
End If
End Sub

Private Sub CommandButton1_Click()
Dim objWrd As Object
Set objWrd = CreateObject("Word.Application")
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

Private Sub IspitatP_Exit(ByVal Cancel As MSForms.ReturnBoolean)
UF1.IspitatP.Value = Trim(UF1.IspitatP.Value)
UF1.IspitatP.Value = Replace(UF1.IspitatP.Value, ".", ",")
ActiveDocument.Variables("IspitatP").Value = Format(UF1.IspitatP.Value, "###0.0#####")
ActiveDocument.Variables("IspitatPMP").Value = Format(UF1.IspitatP.Value / 10, "###0.0#####")
End Sub

Private Sub KartOvaln_Change()
ActiveDocument.Variables("KartOvaln").Value = UF1.KartOvaln.Value
If UF1.KartOvaln.Enabled = True Then
    UF1.Progib.Value = Val(UF1.KartOvaln.Value) + 1
Else
    UF1.Progib.Value = Val(UF1.KartOvaln.Value)
End If
End Sub

Private Sub NaNLet_Exit(ByVal Cancel As MSForms.ReturnBoolean)
UF1.NaNLet.Value = Trim(UF1.NaNLet.Value)
If IsNumeric(UF1.NaNLet.Value) Then 'Если в поле число то добавляем год, года или лет
    If (Val(UF1.NaNLet.Value) = 1) Then ActiveDocument.Variables("NaNLet").Value = UF1.NaNLet.Value & " год"
    If (Val(UF1.NaNLet.Value) > 1 And Val(UF1.NaNLet.Value) < 5) Then ActiveDocument.Variables("NaNLet").Value = UF1.NaNLet & " года"
    If (Val(UF1.NaNLet.Value) > 4) Then ActiveDocument.Variables("NaNLet").Value = UF1.NaNLet.Value & " лет"
    ActiveDocument.Variables("DoNgoda").Value = Format(DateAdd("yyyy", Val(UF1.NaNLet.Value), UF1.AktGID.Value), "dd.mm.yyyy")
Else
    MsgBox ("Нужно ввести число")
End If
End Sub

Private Sub DataAktPoRez_Exit(ByVal Cancel As MSForms.ReturnBoolean)
UF1.DataAktPoRez.Value = Trim(UF1.DataAktPoRez.Value)
ActiveDocument.Variables("DataAktPoRez").Value = UF1.DataAktPoRez.Value
If IsDate(UF1.DataAktPoRez.Value) Then
ActiveDocument.Variables("DateAktPoRez").Value = FormDat(UF1.DataAktPoRez.Value)
End If
End Sub

Private Sub DataIzg_Exit(ByVal Cancel As MSForms.ReturnBoolean)
UF1.DataIzg.Value = Trim(UF1.DataIzg.Value)
ActiveDocument.Variables("DataIzg").Value = UF1.DataIzg.Value
End Sub

Private Sub DataRegistracii_Exit(ByVal Cancel As MSForms.ReturnBoolean)
UF1.DataRegistracii.Value = Trim(UF1.DataRegistracii.Value)
ActiveDocument.Variables("DataRegistracii").Value = UF1.DataRegistracii.Value
End Sub

Private Sub DataVvoda_Exit(ByVal Cancel As MSForms.ReturnBoolean)
UF1.DataVvoda.Value = Trim(UF1.DataVvoda.Value)
ActiveDocument.Variables("DataVvoda").Value = UF1.DataVvoda.Value
ActiveDocument.Variables("SrokSlugb").Value = Year(Date) - Val(Right(UF1.DataVvoda.Value, 4))
End Sub

Private Sub ddlina_Change()
UF1.ddlina.Value = Trim(UF1.ddlina.Value)
ActiveDocument.Variables("ddlina").Value = UF1.ddlina.Value
End Sub

Private Sub dgost_Change()
UF1.dgost.Value = Trim(UF1.dgost.Value)
ActiveDocument.Variables("dgost").Value = UF1.dgost.Value
End Sub

Private Sub Dogovor_Exit(ByVal Cancel As MSForms.ReturnBoolean)
UF1.Dogovor.Value = Trim(UF1.Dogovor.Value)
ActiveDocument.Variables("Dogovor").Value = UF1.Dogovor.Value
End Sub

Private Sub A13_Exit(ByVal Cancel As MSForms.ReturnBoolean)
UF1.A13 = Replace(UF1.A13, ".", ",")
End Sub

Private Sub AktVIK_Change()
ActiveDocument.Variables("AktVIK").Value = UF1.AktVIK.Value
UF1.AktVIKMK.Value = Val(UF1.AktVIK.Value) + 1
End Sub

Private Sub DogovorData_Exit(ByVal Cancel As MSForms.ReturnBoolean)
UF1.DogovorData.Value = Trim(UF1.DogovorData.Value)
ActiveDocument.Variables("DogovorData").Value = UF1.DogovorData.Value
End Sub

Private Sub Button1_Click()
Application.ScreenUpdating = False

If UF1.OptionKotel.Value = True Then ' Если составляем заключение для котлов
    Call DeleteBookmarks("zikl") ' пункт 7.1. циклы
    ActiveDocument.Bookmarks("P7p1mat").Range.Delete ' пункт 7.1. материалы
    ActiveDocument.Bookmarks("VikRezTruboprov").Range.Delete ' ВИК результаты осмотра трубопровода
    ActiveDocument.Bookmarks("TipSvS").Range.Delete ' пункт в УЗК - тип сварного соединения
    ActiveDocument.Variables.Item("PassatT").Value = Strings.ChrW(31)
    ActiveDocument.Variables("OzOsR").Value = UF1.RD1024998.Value
    ActiveDocument.Variables("VIKrd").Value = " п.п. 5.4, 5.5, 5.15"
    ActiveDocument.Variables("punkt7-3NTD").Value = " п." & ActiveDocument.Variables("CBp10").Value & ActiveDocument.Variables("FNPORPDR").Value
'    ActiveDocument.Variables.Item("GOST34347").Value = Strings.ChrW(31)
'    ActiveDocument.Variables.Item("GOST34347PSosEl").Value = Strings.ChrW(31)
'    ActiveDocument.Variables.Item("GOST34347PMat").Value = Strings.ChrW(31)
'    ActiveDocument.Variables.Item("GOST34347PGiI").Value = Strings.ChrW(31)
'    ActiveDocument.Variables("tverdprd").Value = " п. 5.29." & UF1.SO469.Value
'    ActiveDocument.Variables("ovalnrd").Value = "При измерениях овальности барабанов котла установлено, что овальность соответствует требованиям п. 5.10." & UF1.SO469.Value
    ActiveDocument.Variables("RabSredaTopl").Value = "Вид топлива"
End If
If UF1.OptionSosud.Value = True Then ' Если составляем заключение для сосудов
    ActiveDocument.Bookmarks("KotlObor").Range.Delete ' пункт 7.3. оборудование котла
    ActiveDocument.Bookmarks("TipSvS").Range.Delete ' пункт в УЗК - тип сварного соединения
    ActiveDocument.Bookmarks("VikRezTruboprov").Range.Delete ' ВИК результаты осмотра трубопровода
'    ActiveDocument.Variables("punkt7-3NTD").Value = " п.п." & ActiveDocument.Variables("CBp65").Value & ActiveDocument.Variables("CBp68").Value & ActiveDocument.Variables("CBp69").Value
    ActiveDocument.Variables.Item("OzOsR").Value = "ГОСТ 34233.1-2017, ГОСТ 34233.2-2017, ГОСТ 34233.6-2017"
    Call DeleteBookmarks("MetKonstrKotla") ' Металлоконструкции котла
End If
If UF1.OptionTruboprovod.Value = True Then ' Если составляем заключение для трубопровода
    Call DeleteBookmarks("zikl") ' пункт 7.1. циклы
    ActiveDocument.Bookmarks("P7p1mat").Range.Delete ' пункт 7.1. материалы
    ActiveDocument.Bookmarks("KotlObor").Range.Delete ' пункт 7.3. оборудование котла
    ActiveDocument.Bookmarks("VikRezKontr").Range.Delete ' ВИК результаты осмотра
'    ActiveDocument.Variables("punkt7-3NTD").Value = " п.п." & ActiveDocument.Variables("CBp71").Value & ActiveDocument.Variables("CBp80").Value & ActiveDocument.Variables("CBp81").Value & ActiveDocument.Variables("CBp85").Value & ActiveDocument.Variables("CBp86").Value & ActiveDocument.Variables("CBp90").Value & ActiveDocument.Variables("CBp91").Value & ActiveDocument.Variables("FNPORPDR").Value
    ActiveDocument.Variables("punkt7-3-1").Value = "Прокладка и оснащение "
    ActiveDocument.Variables("punkt7-3").Value = Strings.ChrW(31)
    ActiveDocument.Variables.Item("PassatT").Value = Strings.ChrW(31)
    ActiveDocument.Variables("korpusa").Value = Strings.ChrW(31)
    ActiveDocument.Variables("OzOsR").Value = UF1.RD1024998.Value
End If
Call OsnovnPunkt
'Индивидуальные изменения для каждого техустройства
'If UF1.ComboBoxTipUstroistva.Value = "Воздухосборник" Then Call Vozduhosbornik
If UF1.ComboBoxTipUstroistva.Value = "Баллоны групповой установки" Then Call BallGroUst
If UF1.ComboBoxTipUstroistva.Value = "Автоцистерна для СУГ" Then Call Avtozisterna
If UF1.ComboBoxTipUstroistva.Value = "НЖУ, ЦЖУ, УДХ" Then Call NGUCGUUDH
If UF1.ComboBoxTipUstroistva.Value = "Газификатор ГХК" Then Call Gasifikator
If UF1.ComboBoxTipUstroistva.Value = "Сосуд с вакуумом" Then Call VakuumSosud
If UF1.ComboBoxTipUstroistva.Value = "Сосуд под налив" Then Call SosudPodNaliv
If UF1.ComboBoxTipUstroistva.Value = "водогрейный котел" Then Call VodgKotl
If UF1.ComboBoxTipUstroistva.Value = "экономайзер" Then Call Ekonomayzer
If UF1.ComboBoxTipUstroistva.Value = "электрический котел" Then Call ElektroKotel
If UF1.ComboBoxTipUstroistva.Value = "паровой котел" Then Call ParKotl
If UF1.ComboBoxTipUstroistva.Value = "трубопровод пара" Then Call TruboprovPara

'Эксперт по газу
If UF1.ExpertGas.Value = False Then
Call DeleteBookmarks("ExprtPG")
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
'Подготавливаем вариант гидравлического или пневматического испытания
If (UF1.PnIs.Value = False) Then
    Call DeleteBookmarks("PnevmatIsp")
Else
    Call DeleteBookmarks("GidroIsp")
End If
If Val(UF1.KolZicl.Value) > 1000 Then
ActiveDocument.Bookmarks("ziklmen1000").Range.Delete
End If

'Собираем заголовок для сохранения файла
ActiveDocument.BuiltInDocumentProperties("Title").Value = Replace(UF1.NazvTehUstr.Value & " рег.№" & UF1.RegN.Value & "(" & Year(Date) & ")", "/", "-")
UF1.hide
End Sub

Private Sub dstal_Change()
UF1.dstal.Value = Trim(UF1.dstal.Value)
ActiveDocument.Variables("dstal").Value = UF1.dstal.Value
End Sub

Private Sub KolZicl_Exit(ByVal Cancel As MSForms.ReturnBoolean)
UF1.KolZicl.Value = Val(Trim(UF1.KolZicl.Value))
ActiveDocument.Variables("KolZicl").Value = UF1.KolZicl.Value
End Sub

Private Sub ComboBoxTipUstroistva_Change()
Dim SelectComboBox()
If UF1.OptionSosud.Value = True Then
    If UF1.ComboBoxTipUstroistva.ListIndex = 0 Then
        Call ClearAllF
        SelectComboBox = Array(2, 3, 10, 65, 68, 69, 154, 175, 178, 185, 187, 188, 338, 339, 340, 341, 343, 353, 465, 468, 471)
        Call SetComboBox(SelectComboBox)
        ActiveDocument.Variables("TipTehUstr").Value = "сосуд, работающий под давлением"
        ActiveDocument.Variables("TechUsrtva").Value = "воздухосборника"
        ActiveDocument.Variables("NaznTehUstr").Value = "воздухосборник"
        ActiveDocument.Variables("NaznTehUstr").Value = "сосуд для сжатого воздуха"
        UF1.CBFNPORPD.Value = True
        UF1.CBSO439.Value = True
        UF1.CBGOST34347.Value = True
    End If
    If UF1.ComboBoxTipUstroistva.ListIndex = 1 Then
        Call ClearAllF
        SelectComboBox = Array(2, 3, 10, 65, 68, 69, 154, 175, 178, 185, 187, 188, 338, 339, 340, 341, 343, 353, 465, 468, 471)
        Call SetComboBox(SelectComboBox)
        ActiveDocument.Variables("TipTehUstr").Value = "сосуд, работающий под давлением"
        ActiveDocument.Variables("TechUsrtva").Value = "сосуда"
        ActiveDocument.Variables("TechUsrtvo").Value = "сосуд"
        ActiveDocument.Variables("TechUsrtvoB").Value = UCase(Left(ActiveDocument.Variables("TechUsrtvo").Value, 1)) & LCase(Mid(ActiveDocument.Variables("TechUsrtvo").Value, 2)) 'Делаем первую букву большой
        ActiveDocument.Variables("NaznTehUstr").Value = "сосуд для хранения СО2"
        UF1.CBFNPORPD.Value = True
        UF1.CBSO439.Value = True
        UF1.CBGOST34347.Value = True
        UF1.CBRD2626012.Value = True
    End If
    If UF1.ComboBoxTipUstroistva.ListIndex = 2 Then
        Call ClearAllF
        SelectComboBox = Array(2, 3, 10, 175, 179, 185, 186, 187, 188, 465, 468, 469, 471, 538, 539, 540, 577, 589)
        Call SetComboBox(SelectComboBox)
        ActiveDocument.Variables("TipTehUstr").Value = "баллоны групповой установки"
        ActiveDocument.Variables("TechUsrtva").Value = "баллонов"
        ActiveDocument.Variables("TechUsrtvo").Value = "баллоны групповой установки"
        ActiveDocument.Variables("TechUsrtvoB").Value = UCase(Left(ActiveDocument.Variables("TechUsrtvo").Value, 1)) & LCase(Mid(ActiveDocument.Variables("TechUsrtvo").Value, 2)) 'Делаем первую букву большой
        ActiveDocument.Variables("NaznTehUstr").Value = "для подъема давления"
        Call MnogCHislo
        UF1.CBFNPORPD.Value = True
        UF1.CBSO439.Value = True
        UF1.CBGOST34347.Value = True
    End If
    If UF1.ComboBoxTipUstroistva.ListIndex = 3 Then
        Call ClearAllF
        SelectComboBox = Array(2, 3, 10, 65, 68, 69, 100, 154, 175, 178, 185, 187, 188, 338, 339, 340, 343, 353, 465, 468)
        Call SetComboBox(SelectComboBox)
        ActiveDocument.Variables("TipTehUstr").Value = "сосуд, работающий под давлением"
        ActiveDocument.Variables("NaznTehUstr").Value = "для хранения и транспортировки СУГ"
        UF1.CBFNPORPD.Value = True
        UF1.CBSO439.Value = True
        UF1.CBGOST34347.Value = True
        UF1.CBRD2626012.Value = True
    End If
    If UF1.ComboBoxTipUstroistva.ListIndex = 4 Then
        Call ClearAllF
        SelectComboBox = Array(2, 3, 10, 68, 69, 100, 175, 178, 185, 186, 187, 188, 338, 339, 340, 341, 343, 353, 465, 468)
        Call SetComboBox(SelectComboBox)
        ActiveDocument.Variables("TechUsrtva").Value = "подогревателя"
        ActiveDocument.Variables("TechUsrtvo").Value = "подогреватель"
        ActiveDocument.Variables("TechUsrtvoB").Value = UCase(Left(ActiveDocument.Variables("TechUsrtvo").Value, 1)) & LCase(Mid(ActiveDocument.Variables("TechUsrtvo").Value, 2)) 'Делаем первую букву большой
        ActiveDocument.Variables("NaznTehUstr").Value = "подогрев сетевой воды"
        UF1.CBFNPORPD.Value = True
        UF1.CBSO439.Value = True
        UF1.CBGOST34347.Value = True
    End If
    If UF1.ComboBoxTipUstroistva.ListIndex = 5 Then
        Call ClearAllF
        SelectComboBox = Array(2, 3, 10, 65, 68, 69, 100, 154, 175, 190, 338, 339, 340, 341, 343, 353, 465, 468, 471)
        Call SetComboBox(SelectComboBox)
        ActiveDocument.Variables("TechUsrtva").Value = "газификатора"
        ActiveDocument.Variables("TechUsrtvo").Value = "газификатор"
        ActiveDocument.Variables("TechUsrtvoB").Value = UCase(Left(ActiveDocument.Variables("TechUsrtvo").Value, 1)) & LCase(Mid(ActiveDocument.Variables("TechUsrtvo").Value, 2)) 'Делаем первую букву большой
        ActiveDocument.Variables("NaznTehUstr").Value = "хранение сжиженных криопродуктов"
        UF1.CBFNPORPD.Value = True
        UF1.CBVM030104.Value = True
    End If
    If UF1.ComboBoxTipUstroistva.ListIndex = 6 Then
        Call ClearAllF
        SelectComboBox = Array()
        Call SetComboBox(SelectComboBox)
        ActiveDocument.Variables("TipTehUstr").Value = "сосуд, работающий без давления (под вакуумом)"
        ActiveDocument.Variables("p9OPVB").Value = Strings.ChrW(31)
        ActiveDocument.Variables("p161OPVB").Value = " 161,"
        ActiveDocument.Variables("p164OPVB").Value = " 164,"
        ActiveDocument.Variables("p169OPVB").Value = " 169,"
        ActiveDocument.Variables("p177OPVB").Value = " 177,"
        ActiveDocument.Variables("p178OPVB").Value = " 178,"
        ActiveDocument.Variables("p179OPVB").Value = " 179,"
        ActiveDocument.Variables("p2-102RUA").Value = " 2.102,"
        ActiveDocument.Variables("p2-111RUA").Value = Strings.ChrW(31)
        ActiveDocument.Variables("NaznTehUstr").Value = "вакуумная емкость"
        UF1.CBFNPOPVB.Value = True
        UF1.CBGOST34347.Value = True
        UF1.CBRUA93.Value = True
    End If
    If UF1.ComboBoxTipUstroistva.ListIndex = 7 Then
        Call ClearAllF
        SelectComboBox = Array()
        Call SetComboBox(SelectComboBox)
        ActiveDocument.Variables("TipTehUstr").Value = "сосуд, работающий без давления (под налив)"
        ActiveDocument.Variables("p9OPVB").Value = " 9,"
        ActiveDocument.Variables("p161OPVB").Value = " 161,"
        ActiveDocument.Variables("p164OPVB").Value = " 164,"
        ActiveDocument.Variables("p169OPVB").Value = " 169,"
        ActiveDocument.Variables("p177OPVB").Value = " 177,"
        ActiveDocument.Variables("p178OPVB").Value = " 178,"
        ActiveDocument.Variables("p179OPVB").Value = " 179,"
        ActiveDocument.Variables("p2-102RUA").Value = Strings.ChrW(31)
        ActiveDocument.Variables("p2-111RUA").Value = " 2.111,"
        ActiveDocument.Variables("NaznTehUstr").Value = "емкость под налив"
        UF1.CBFNPOPVB.Value = True
        UF1.CBGOST34347.Value = True
        UF1.CBRUA93.Value = True
    End If
End If
If UF1.OptionKotel.Value = True Then
    If UF1.ComboBoxTipUstroistva.ListIndex = 0 Then
        Call ClearAllF
        SelectComboBox = Array(2, 3, 10, 38, 39, 43, 45, 47, 49, 50, 100, 154, 175, 177, 185, 186, 187, 188, 257, 258, 260, 267, 271, 465, 468, 471)
        Call SetComboBox(SelectComboBox)
        ActiveDocument.Variables("TipTehUstr").Value = "паровой котел"
        ActiveDocument.Variables("NaznTehUstr").Value = "выработка пара"
        UF1.CBFNPORPD.Value = True
        UF1.CBSO469.Value = True
    End If
    If UF1.ComboBoxTipUstroistva.ListIndex = 1 Then
        Call ClearAllF
        SelectComboBox = Array(2, 3, 10, 38, 39, 45, 46, 47, 50, 154, 175, 177, 185, 186, 187, 188, 267, 271, 465, 468, 471)
        Call SetComboBox(SelectComboBox)
        ActiveDocument.Variables("TipTehUstr").Value = "водогрейный котел"
        ActiveDocument.Variables("NaznTehUstr").Value = "нагрев сетевой воды"
        UF1.CBFNPORPD.Value = True
        UF1.CBSO469.Value = True
    End If
    If UF1.ComboBoxTipUstroistva.ListIndex = 2 Then
        Call ClearAllF
        SelectComboBox = Array(2, 3, 10, 22, 39, 43, 50, 100, 154, 175, 178, 185, 186, 187, 188, 257, 258, 260, 267, 271, 465, 468, 471, 500, 502, 503, 505, 506)
        Call SetComboBox(SelectComboBox)
        ActiveDocument.Variables("TipTehUstr").Value = "электрокотел"
        ActiveDocument.Variables("NaznTehUstr").Value = "выработка пара"
        UF1.CBFNPORPD.Value = True
        UF1.CBSO469.Value = True
    End If
    If UF1.ComboBoxTipUstroistva.ListIndex = 3 Then
        Call ClearAllF
        SelectComboBox = Array(2, 3, 10, 45, 61, 64, 175, 177, 185, 186, 187, 188, 267, 268, 269, 270, 271, 465, 468)
        Call SetComboBox(SelectComboBox)
        ActiveDocument.Variables("TipTehUstr").Value = "экономайзер"
        ActiveDocument.Variables("TechUsrtva").Value = "экономайзера"
        ActiveDocument.Variables("TechUsrtvo").Value = "экономайзер"
        ActiveDocument.Variables("TechUsrtvoB").Value = UCase(Left(ActiveDocument.Variables("TechUsrtvo").Value, 1)) & LCase(Mid(ActiveDocument.Variables("TechUsrtvo").Value, 2)) 'Делаем первую букву большой
        ActiveDocument.Variables("NaznTehUstr").Value = "подогрев питательной воды"
        UF1.CBFNPORPD.Value = True
        UF1.CBSO469.Value = True
    End If
End If
If UF1.OptionTruboprovod.Value = True Then
    If UF1.ComboBoxTipUstroistva.ListIndex = 0 Then
        Call ClearAllF
        SelectComboBox = Array(2, 3, 10, 71, 80, 81, 85, 86, 90, 91, 154, 156, 175, 184, 185, 187, 188, 372, 373, 374, 394, 465, 468, 469, 471)
        Call SetComboBox(SelectComboBox)
        ActiveDocument.Variables("TipTehUstr").Value = "трубопровод пара"
        ActiveDocument.Variables("NaznTehUstr").Value = "транспортировка пара"
        UF1.CBFNPORPD.Value = True
        UF1.CBSO464.Value = True
    End If
    If UF1.ComboBoxTipUstroistva.ListIndex = 1 Then
        Call ClearAllF
        SelectComboBox = Array(2, 3, 10, 71, 80, 81, 85, 86, 90, 91, 154, 156, 175, 184, 185, 187, 188, 372, 373, 374, 394, 465, 468, 469, 471)
        Call SetComboBox(SelectComboBox)
        ActiveDocument.Variables("TipTehUstr").Value = "трубопровод горячего водоснабжения"
        ActiveDocument.Variables("NaznTehUstr").Value = "транспортировка сетевой воды"
        UF1.CBFNPORPD.Value = True
        UF1.CBSO464.Value = True
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

Private Sub NazvOPO_Exit(ByVal Cancel As MSForms.ReturnBoolean)
UF1.NazvOPO.Value = Trim(UF1.NazvOPO.Value)
ActiveDocument.Variables("NazvOPO").Value = UF1.NazvOPO.Value
End Sub

Private Sub NazvTehUstr_Exit(ByVal Cancel As MSForms.ReturnBoolean)
UF1.NazvTehUstr.Value = Trim(UF1.NazvTehUstr.Value)
ActiveDocument.Variables("NazvTehUstr").Value = UF1.NazvTehUstr.Value
End Sub

Private Sub ogost_Change()
UF1.ogost.Value = Trim(UF1.ogost.Value)
ActiveDocument.Variables("ogost").Value = UF1.ogost.Value
End Sub

Private Sub OptionKotel_Change()
If UF1.OptionKotel.Value = True Then
    With UF1.ComboBoxTipUstroistva
        .Clear
        .AddItem "паровой котел"
        .AddItem "водогрейный котел"
        .AddItem "электрический котел"
        .AddItem "экономайзер"
    End With
    UF1.AktVIKMK.Enabled = True
    UF1.AktVIKMKD.Enabled = True
    UF1.VikMK.Value = True
    UF1.KontrGib.Enabled = True
    UF1.KontrGibD.Enabled = True
    UF1.KontrGibCh.Value = True
    ActiveDocument.Variables("TechUsrtva").Value = "котла"
    ActiveDocument.Variables("TechUsrtvo").Value = "котел"
    ActiveDocument.Variables("TechUsrtvoB").Value = UCase(Left(ActiveDocument.Variables("TechUsrtvo").Value, 1)) & LCase(Mid(ActiveDocument.Variables("TechUsrtvo").Value, 2)) 'Делаем первую букву большой
    UF1.Label418.Caption = "Вид топлива"
    UF1.KolZicl.Visible = False
    UF1.Label18.Visible = False
    UF1.Label5.Caption = "Завод изг."
'    ActiveDocument.Variables("SOrd").Value = UF1.SO469.Value
    ActiveDocument.Variables("TimeGI").Value = "20 минут"
End If
End Sub

Private Sub OptionSosud_Change()
If UF1.OptionSosud.Value = True Then
    With UF1.ComboBoxTipUstroistva
        .Clear
        .AddItem "Воздухосборник"
        .AddItem "НЖУ, ЦЖУ, УДХ"
        .AddItem "Баллоны групповой установки"
        .AddItem "Автоцистерна для СУГ"
        .AddItem "Подогреватель"
        .AddItem "Газификатор ГХК"
        .AddItem "Сосуд с вакуумом"
        .AddItem "Сосуд под налив"
    End With
    ActiveDocument.Variables("TechUsrtva").Value = "сосуда"
    ActiveDocument.Variables("TechUsrtvo").Value = "сосуд"
    ActiveDocument.Variables("TechUsrtvoB").Value = UCase(Left(ActiveDocument.Variables("TechUsrtvo").Value, 1)) & LCase(Mid(ActiveDocument.Variables("TechUsrtvo").Value, 2)) 'Делаем первую букву большой
'    ActiveDocument.Variables("SOrd").Value = UF1.SO439.Value
    UF1.KolZicl.Visible = True
    UF1.Label18.Visible = True
    UF1.Label5.Caption = "Завод изг."
    ActiveDocument.Variables("TimeGI").Value = "10 минут"
End If
End Sub

Private Sub OptionTruboprovod_Change()
If UF1.OptionTruboprovod.Value = True Then
    With UF1.ComboBoxTipUstroistva
        .Clear
        .AddItem "трубопровод пара"
        .AddItem "трубопровод горячей воды"
    End With
    ActiveDocument.Variables("TechUsrtva").Value = "трубопровода"
    ActiveDocument.Variables("TechUsrtvo").Value = "трубопровод"
    ActiveDocument.Variables("TechUsrtvoB").Value = UCase(Left(ActiveDocument.Variables("TechUsrtvo").Value, 1)) & LCase(Mid(ActiveDocument.Variables("TechUsrtvo").Value, 2)) 'Делаем первую букву большой
    UF1.KolZicl.Visible = False
    UF1.Label18.Visible = False
    UF1.Label5.Caption = "Монт. орг."
    ActiveDocument.Variables("TimeGI").Value = "10 минут"
End If
End Sub

Private Sub ostal_Change()
UF1.ostal.Value = Trim(UF1.ostal.Value)
ActiveDocument.Variables("ostal").Value = UF1.ostal.Value
End Sub

Private Sub otolsh_Change()
UF1.otolsh.Value = Trim(UF1.otolsh.Value)
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

Private Sub Predpriyatie_Exit(ByVal Cancel As MSForms.ReturnBoolean)
UF1.Predpriyatie.Value = Trim(UF1.Predpriyatie.Value)
ActiveDocument.Variables("Predpriyatie").Value = UF1.Predpriyatie.Value
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

Private Sub RabocheeP_Exit(ByVal Cancel As MSForms.ReturnBoolean)
UF1.RabocheeP.Value = Trim(UF1.RabocheeP.Value)
UF1.RabocheeP.Value = Replace(UF1.RabocheeP.Value, ".", ",")
UF1.RazreshaemoeP.Value = UF1.RabocheeP.Value
ActiveDocument.Variables("RabocheeP").Value = Format(UF1.RabocheeP.Value, "###0.0#####") & " кгс/см" & Strings.ChrW(178)
End Sub

Private Sub RabSreda_Exit(ByVal Cancel As MSForms.ReturnBoolean)
UF1.RabSreda.Value = Trim(UF1.RabSreda.Value)
ActiveDocument.Variables("RabSreda").Value = UF1.RabSreda.Value
End Sub

Private Sub RabTemp_Change()
ActiveDocument.Variables("RabTemp").Value = UF1.RabTemp.Value
End Sub

Private Sub RaschetnP_Exit(ByVal Cancel As MSForms.ReturnBoolean)
UF1.RaschetnP.Value = Trim(UF1.RaschetnP.Value)
UF1.RaschetnP.Value = Replace(UF1.RaschetnP.Value, ".", ",")
ActiveDocument.Variables("RaschetnP").Value = Format(UF1.RaschetnP.Value, "###0.0#####") & " кгс/см" & Strings.ChrW(178)
End Sub

Private Sub RazreshaemoeP_Exit(ByVal Cancel As MSForms.ReturnBoolean)
UF1.RazreshaemoeP.Value = Trim(UF1.RazreshaemoeP.Value)
UF1.RazreshaemoeP = Replace(UF1.RazreshaemoeP, ".", ",")
If IsNumeric(UF1.RazreshaemoeP.Value) Then
    UF1.IspitatP.Value = Format((UF1.RazreshaemoeP.Value * 1.25), "###0.0#")
    ActiveDocument.Variables("RazreshaemoeP").Value = Format(UF1.RazreshaemoeP.Value, "###0.0#####") & " кгс/см" & Strings.ChrW(178)
    ActiveDocument.Variables("RazreshaemoePKrt").Value = Format(UF1.RazreshaemoeP.Value, "###0.0#####")
    ActiveDocument.Variables("RazreshaemoePMP").Value = Format(Val(UF1.RazreshaemoeP.Value) / 10, "###0.0#####") & " МПа"
Else
    ActiveDocument.Variables("RazreshaemoeP").Value = UF1.RazreshaemoeP.Value
    ActiveDocument.Variables("RazreshaemoePKrt").Value = UF1.RazreshaemoeP.Value
    ActiveDocument.Variables("RazreshaemoePMP").Value = Strings.ChrW(31)
End If
If UF1.ComboBoxTipUstroistva.Value = "Автоцистерна для СУГ" And Val(UF1.RazreshaemoeP.Value) < Val(UF1.RaschetnP.Value) Then
    MsgBox ("Для сосудов с СУГ не допускается снижение давления. П.402 ФНП ОРПД")
End If
End Sub

Private Sub RegN_Exit(ByVal Cancel As MSForms.ReturnBoolean)
UF1.RegN.Value = Trim(UF1.RegN.Value)
ActiveDocument.Variables("RegN").Value = UF1.RegN.Value
End Sub

Private Sub RegNOPO_Exit(ByVal Cancel As MSForms.ReturnBoolean)
UF1.RegNOPO.Value = Trim(UF1.RegNOPO.Value)
ActiveDocument.Variables("RegNOPO").Value = UF1.RegNOPO.Value
If UF1.RegNOPO.Value = "" Then ActiveDocument.Variables("RegNOPO").Value = Strings.ChrW(31)
End Sub

Private Sub SavePribor_Click()
Dim objObject As String
Dim objObject1 As String

Open MyFilePribor For Output As 1
For i = 1 To 16
    objObject = "pribor" & i
    objObject1 = "dp" & i
    Print #1, UF1.Controls.Item(objObject).Value
    Print #1, UF1.Controls.Item(objObject1).Value
    Next i

Close 1

End Sub

'Private Sub M3_Change()
'If UF1.M3.Value = True Then
'UF1.A6 = UF1.A6 & " м" & ChrW(179)
'End If
'End Sub


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
'Задаем значения комбобох
    With UF1.ComboBoxTipUstroistva
        .Clear
        .AddItem "Воздухосборник"
        .AddItem "НЖУ, ЦЖУ, УДХ"
        .AddItem "Баллоны групповой установки"
        .AddItem "Автоцистерна для СУГ"
        .AddItem "Подогреватель"
        .AddItem "Газификатор ГХК"
        .AddItem "Сосуд с вакуумом"
        .AddItem "Сосуд под налив"
    End With
'Загружаем из файла приборы
Dim objObject As String
Dim objObject1 As String
Dim Flag As Boolean
Flag = False
Dim a
Dim x1 As Date
Open MyFilePribor For Input As 1
For i = 1 To 16
    objObject = "pribor" & i
    objObject1 = "dp" & i
    Line Input #1, a
    UF1.Controls.Item(objObject).Value = a
    Line Input #1, a
    UF1.Controls.Item(objObject1).Value = a
    x1 = a
    If (x1 < Date) Then
    Flag = True
    UF1.Controls.Item(objObject1).BackColor = 1000
    End If
    Next i
Close 1

UF1.DataAktPoRez = Format(Date, "dd.mm.yyyy")
ActiveDocument.Variables("DataAktPoRez").Value = Format(Date, "dd.mm.yyyy")

If (Flag) Then
MsgBox "Имеются просроченные приборы"
End If

For Each ctl In UF1.Controls
    If TypeName(ctl) = "CheckBox" And Left(ctl.Name, 3) = "CBp" Then
    AllCBp.Add (Val(ctl.Caption))
    End If
Next
Call QuickSort(AllCBp, 1, AllCBp.Count)

End Sub

Sub SetComboBox(SelectComboBox)
For Each mark In AllCBp
    UF1.Controls.Item("CBp" & mark).Value = False
Next

For Each mark In SelectComboBox
    UF1.Controls.Item("CBp" & mark).Value = True
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

Private Sub ZakMPDZD_Change()
ActiveDocument.Variables("ZakMPDZD").Value = UF1.ZakMPDZD.Value
If UF1.ZakMPDZD.Enabled = True Then
    UF1.ProtkTV.Value = Val(UF1.ZakMPDZD.Value) + 1
Else
    UF1.ProtkTV.Value = Val(UF1.ZakMPDZD.Value)
End If
End Sub

Private Sub ZakMPDZDD_Change()
ActiveDocument.Variables("ZakMPDZDD").Value = UF1.ZakMPDZDD.Value
ActiveDocument.Variables("ZakMPDZDData").Value = FormDat(UF1.ZakMPDZDD.Value)
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
ActiveDocument.Variables("ZakUZKD").Value = UF1.ZakUZKD.Value
ActiveDocument.Variables("ZakUZKData").Value = FormDat(UF1.ZakUZKD.Value)
End Sub

Private Sub ZavN_Exit(ByVal Cancel As MSForms.ReturnBoolean)
UF1.ZavN.Value = Trim(UF1.ZavN.Value)
ActiveDocument.Variables("ZavN").Value = UF1.ZavN.Value
End Sub

Private Sub ZavNIzd_Exit(ByVal Cancel As MSForms.ReturnBoolean)
UF1.ZavNIzd.Value = Trim(UF1.ZavNIzd.Value)
End Sub

Private Sub ZavodIzg_Exit(ByVal Cancel As MSForms.ReturnBoolean)
UF1.ZavodIzg.Value = Trim(UF1.ZavodIzg.Value)
ActiveDocument.Variables("ZavodIzg").Value = UF1.ZavodIzg.Value
End Sub
