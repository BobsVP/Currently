Attribute VB_Name = "Start"
'Public MyFilePribor As String 'Файл в котором храним приборы
Public AllCBp As New Collection  'Переменная для хранения  всех пунктов ФНП ОРПД
Public AllCBv As New Collection  'Переменная для хранения  всех пунктов ФНП ОПВБ
Public AllCBh As New Collection  'Переменная для хранения  всех пунктов ФНП ХОПО
Public AllCBt As New Collection  'Переменная для хранения  всех пунктов ФНП ТТ
Public AllCBs As New Collection  'Переменная для хранения  всех пунктов ФНП СНН
Public CBpclass() As New CBcont     'Регистрируем класс для объектов ФНП ОРПД
Public CBvclass() As New CBcont     'Регистрируем класс для объектов ФНП ОПВБ
Public CBhclass() As New CBcont     'Регистрируем класс для объектов ФНП ХОПО
Public CBtclass() As New CBcont     'Регистрируем класс для объектов ФНП ТТ
Public CBsclass() As New CBcont     'Регистрируем класс для объектов ФНП СНН
Public DateBase As Object           'Файл в котором храним все данные

Sub Main()
'
' А1 Макрос
' Макрос создан 15.04.2011
'
'Set AllCBp = New Scripting.Dictionary
'MyFilePribor = ActiveDocument.AttachedTemplate.Path
'MyFilePribor = MyFilePribor & "\tablprib.txt"
    Dim sFilePath
    sFilePath = ActiveDocument.AttachedTemplate.Path & "\data_base.xls"    'ActiveDocument.Path &
On Error Resume Next
    Set DateBase = GetObject(, "Excel.Application") 'http://www.excelworld.ru/stuff/vba_function/object/getobject/28-1-0-132
If Err.Number <> 0 Then Err.Clear
    If DateBase Is Nothing Then
        Set DateBase = CreateObject("Excel.Application")       'получаем указатель на Application
        DateBase.Workbooks.Open sFilePath
    End If
    DateBase.Visible = False                               'скрываем окно Excel
    DateBase.ScreenUpdating = False
'DateBase.Workbooks("data_base.xls").Worksheets("tablprib").Range("A20").Value = "Проверка связи"
'Dim UF1 As New Form1 'при объявлении как переменная начинаются глюки
UF1.Show
    
'    DateBase.Close False
    DateBase.Workbooks("data_base.xls").Save
    DateBase.Quit          'убираем указатель на excel
    Set DateBase = Nothing

    
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
Call FoAndRe("требованиям;", "требованиям")
Call FoAndRe("требованиями;", "требованиями")
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
Dim Flag As Boolean
Flag = False

    If ActiveDocument.Bookmarks.Exists("KotlObor") = True Then ActiveDocument.Bookmarks("KotlObor").Range.Delete ' пункт 7.3. оборудование котла по новому оформлению все в одном пункте
    ActiveDocument.Variables("punkt7-3").Value = Strings.ChrW(31)
    ActiveDocument.Bookmarks("R7p4").Range.Delete 'Пункт про КИПиА

If UF1.CBFNPORPD.Value = True Then
    ActiveDocument.Variables("punkt1-1").Value = "п.п." 'Сборка пункта 1-1 для ОРПД
    For Each mark In AllCBp
        If UF1.Controls.Item("CBp" & mark).Value = True Then ActiveDocument.Variables("punkt1-1").Value = ActiveDocument.Variables("punkt1-1").Value & ActiveDocument.Variables("CBp" & mark).Value
    Next
    ActiveDocument.Variables("punkt1-1").Value = Left(ActiveDocument.Variables("punkt1-1").Value, Len(ActiveDocument.Variables("punkt1-1").Value) - 1) & UF1.FNPORPDR.Value & ActiveDocument.Variables("TckZpt1").Value
'    ActiveDocument.Variables("FNPORPDR").Value = UF1.FNPORPDR.Value
    ActiveDocument.Variables("p7-1ORPDProdl").Value = " п.п." & ActiveDocument.Variables("CBp2").Value & ActiveDocument.Variables("CBp3").Value & ActiveDocument.Variables("CBp394").Value & ActiveDocument.Variables("CBp465").Value
    ActiveDocument.Variables("p7-1ORPDProdl").Value = ActiveDocument.Variables("p7-1ORPDProdl").Value & ActiveDocument.Variables("CBp468").Value & ActiveDocument.Variables("CBp471").Value & UF1.FNPORPDR.Value
    SetVar = Array(38, 39, 43, 45, 46, 47, 49, 50, 61, 64, 71, 80, 81, 85, 86, 90, 91)
    ActiveDocument.Variables("p7-3ORPD").Value = " п.п."
    Call SetValue(SetVar, "CBp", "p7-3ORPD")
    ActiveDocument.Variables("p7-3ORPD").Value = ActiveDocument.Variables("p7-3ORPD").Value & UF1.FNPORPDR.Value
'    SetVar = Array(257, 258, 260, 267, 268, 269, 270, 271, 338, 339, 340, 341, 342, 343, 351, 353, 354, 500, 502, 503, 505, 506, 519, 521, 523, 540)
'    ActiveDocument.Variables("p7-4ORPD").Value = " п.п."
'    Call SetValue(SetVar, "CBp", "p7-4ORPD")
'    ActiveDocument.Variables("p7-4ORPD").Value = ActiveDocument.Variables("p7-4ORPD").Value & UF1.FNPORPDR.Value
    SetVar = Array(10, 22, 38, 39, 43, 45, 46, 47, 49, 50, 61, 64, 65, 68, 69, 71, 80, 81, 85, 86, 90, 91, 257, 258, 260, 267, 268, 269, 270, 271, _
    338, 339, 340, 341, 342, 343, 351, 353, 354, 500, 502, 503, 505, 506, 519, 521, 523, 538, 539, 540, 577, 589)
    Call SetValue(SetVar, "CBp", "punkt7-3NTD")
    ActiveDocument.Variables("punkt7-3NTD").Value = " п.п." & ActiveDocument.Variables("punkt7-3NTD").Value & UF1.FNPORPDR.Value
    ActiveDocument.Variables("p7-5ORPD").Value = "; п.п." & ActiveDocument.Variables("CBp394").Value & ActiveDocument.Variables("CBp465").Value & ActiveDocument.Variables("CBp468").Value & UF1.FNPORPDR.Value
    ActiveDocument.Variables("p-8FNPORPDR").Value = UF1.FNPORPDR.Value
    ActiveDocument.Variables("p12-1pril2").Value = " п. 12.1. Приложения №2" & ActiveDocument.Variables("PunktPril8ORPD").Value & UF1.FNPORPDR.Value
    ActiveDocument.Variables("p12-2pril2").Value = " п. 12.2. Приложения №2" & UF1.FNPORPDR.Value
    ActiveDocument.Variables("p12-3pril2").Value = " п. 12.3. Приложения №2" & UF1.FNPORPDR.Value
    ActiveDocument.Variables("p12-5pril2").Value = " п. 12.5. Приложения №2, п. 3 Приложения №8" & UF1.FNPORPDR.Value 'УЗК
    ActiveDocument.Variables("p1-pril8").Value = ", что соответствует требованиям п. 1. Приложения №8" & UF1.FNPORPDR.Value
'    For Each mark In AllCBp 'Установка размещение и обвязка
'        If mark > 9 And mark < 92 Then ActiveDocument.Variables("punkt7-3NTD").Value = ActiveDocument.Variables("punkt7-3NTD").Value & ActiveDocument.Variables("CBp" & mark).Value
'    Next
'    ActiveDocument.Variables("punkt7-3NTD").Value = ActiveDocument.Variables("punkt7-3NTD").Value & ActiveDocument.Variables("CBp538").Value & ActiveDocument.Variables("CBp539").Value & ActiveDocument.Variables("CBp577").Value & ActiveDocument.Variables("CBp589").Value
    SetVar = Array(175, 177, 178, 179, 184, 185, 186, 187, 188, 190, 469) 'ГИ
    Call SetValue(SetVar, "CBp", "GIFNPORPD")
    ActiveDocument.Variables("GIFNPORPD").Value = " п.п." & ActiveDocument.Variables("GIFNPORPD").Value & UF1.FNPORPDR.Value
    ActiveDocument.Variables("PIFNPORPD").Value = " п.п. 175, 190" & UF1.FNPORPDR.Value
    ActiveDocument.Variables("PIFNPORPD1").Value = " п. 190" & UF1.FNPORPDR.Value
    ActiveDocument.Variables("NTDAktNKFNPORPD").Value = "Федеральные нормы и правила в области промышленной безопасности " & Mid(UF1.FNPORPDR.Value, 64, 104) & "."
    ActiveDocument.Variables("TckZpt").Value = ";"
    ActiveDocument.Variables("TckZpt1").Value = ";" & Strings.Chr(13)
    Flag = True
End If
If UF1.CBFNPOPVB.Value = True Then
    ActiveDocument.Variables("punkt1-1OPVB").Value = "п.п." 'Сборка пункта 1-1 для ОПВБ
    For Each mark In AllCBv
        If UF1.Controls.Item("CBvb" & mark).Value = True Then ActiveDocument.Variables("punkt1-1OPVB").Value = ActiveDocument.Variables("punkt1-1OPVB").Value & ActiveDocument.Variables("CBvb" & mark).Value
    Next
    ActiveDocument.Variables("punkt1-1OPVB").Value = ActiveDocument.Variables("TckZpt1").Value & Left(ActiveDocument.Variables("punkt1-1OPVB").Value, Len(ActiveDocument.Variables("punkt1-1OPVB").Value) - 1) & UF1.FNPOPVBR.Value
    If ActiveDocument.Variables("CBvb9").Value <> Strings.ChrW(31) Then ActiveDocument.Variables("p7-1OPVBTechRegl").Value = " Эксплуатация " & ActiveDocument.Variables("TechUsrtva").Value & " осуществляется в соответствии с технологическим регламентом, что соответствует требованиям п. 9" & UF1.FNPOPVBR.Value & "."
    ActiveDocument.Variables("p7-1OPVB").Value = ActiveDocument.Variables("CBvb164").Value & ActiveDocument.Variables("CBvb182").Value & ActiveDocument.Variables("CBvb193").Value & ActiveDocument.Variables("CBvb203").Value
    If Len(ActiveDocument.Variables("p7-1OPVB").Value) < 9 Then
        ActiveDocument.Variables("p7-1OPVB").Value = " п." & ActiveDocument.Variables("p7-1OPVB").Value & UF1.FNPOPVBR.Value
    Else
        ActiveDocument.Variables("p7-1OPVB").Value = " п.п." & ActiveDocument.Variables("p7-1OPVB").Value & UF1.FNPOPVBR.Value
    End If
    ActiveDocument.Variables("p7-1OPVBProdl").Value = ActiveDocument.Variables("TckZpt").Value & " п." & ActiveDocument.Variables("CBvb161").Value & UF1.FNPOPVBR.Value
    ActiveDocument.Variables("TTrDop1").Value = " Техническое состояние " & ActiveDocument.Variables("TechUsrtva").Value & " соответствует требованиям п.п. 30, 169" & UF1.FNPOPVBR.Value & "."
    SetVar = Array(43, 47, 48, 53, 177, 178, 179, 184, 185, 186, 189, 190, 196, 197, 198, 199, 203)
    Call SetValue(SetVar, "CBvb", "p7-3OPVB")
    ActiveDocument.Variables("p7-3OPVB").Value = ActiveDocument.Variables("TckZpt").Value & " п.п." & ActiveDocument.Variables("p7-3OPVB").Value & UF1.FNPORPDR.Value
'    ActiveDocument.Variables("p7-3OPVB").Value = ActiveDocument.Variables("TckZpt").Value & " п.п." & ActiveDocument.Variables("CBvb177").Value & ActiveDocument.Variables("CBvb178").Value & ActiveDocument.Variables("CBvb179").Value
'    ActiveDocument.Variables("p7-3OPVB").Value = ActiveDocument.Variables("p7-3OPVB").Value & ActiveDocument.Variables("CBvb196").Value & ActiveDocument.Variables("CBvb197").Value & ActiveDocument.Variables("CBvb198").Value
'    ActiveDocument.Variables("p7-3OPVB").Value = ActiveDocument.Variables("p7-3OPVB").Value & ActiveDocument.Variables("CBvb199").Value & ActiveDocument.Variables("CBvb203").Value & UF1.FNPOPVBR.Value
    ActiveDocument.Variables("p-8FNPOPVB").Value = ActiveDocument.Variables("TckZpt").Value & UF1.FNPOPVBR.Value
'    ActiveDocument.Variables("GIFNPOPVB").Value = ActiveDocument.Variables("TckZpt").Value & " п." & ActiveDocument.Variables("CBvb169").Value & UF1.FNPOPVBR.Value
    ActiveDocument.Variables("OsnRez-OPVB").Value = " п.п." & ActiveDocument.Variables("CBvb120").Value & ActiveDocument.Variables("CBvb121").Value & UF1.FNPOPVBR.Value & ActiveDocument.Variables("TckZpt1").Value
    ActiveDocument.Variables("NTDAktNKFNPOPVB").Value = "Федеральные нормы и правила в области промышленной безопасности " & Mid(UF1.FNPOPVBR.Value, 64, 122) & "."
    If ActiveDocument.Variables("TckZpt").Value <> Strings.ChrW(31) Then ActiveDocument.Variables("NTDAktNKFNPOPVB").Value = Strings.Chr(13) & ActiveDocument.Variables("NTDAktNKFNPOPVB").Value
    ActiveDocument.Variables("TckZpt1").Value = ";" & Strings.Chr(13)
    ActiveDocument.Variables("TckZpt").Value = ";"
    ActiveDocument.Variables("TckZpt7-1").Value = ";"
    Flag = True
End If
If UF1.CBFNPHOPO.Value = True Then
    ActiveDocument.Variables("punkt1-1HOPO").Value = "п.п." 'Сборка пункта 1-1 для ХОПО
    For Each mark In AllCBh
        If UF1.Controls.Item("CBho" & mark).Value = True Then ActiveDocument.Variables("punkt1-1HOPO").Value = ActiveDocument.Variables("punkt1-1HOPO").Value & ActiveDocument.Variables("CBho" & mark).Value
    Next
    ActiveDocument.Variables("punkt1-1HOPO").Value = ActiveDocument.Variables("TckZpt1").Value & Left(ActiveDocument.Variables("punkt1-1HOPO").Value, Len(ActiveDocument.Variables("punkt1-1HOPO").Value) - 1) & UF1.FNPHOPO.Value
    SetVar = Array(140, 142, 234)
    Call SetValue(SetVar, "CBho", "p7-1HOPO")
    If Len(ActiveDocument.Variables("p7-1HOPO").Value) > UBound(SetVar) + 2 Then
        ActiveDocument.Variables("p7-1HOPO").Value = Replace(ActiveDocument.Variables("p7-1HOPO").Value, Strings.ChrW(31), "")
        ActiveDocument.Variables("p7-1HOPO").Value = Left(ActiveDocument.Variables("p7-1HOPO").Value, Len(ActiveDocument.Variables("p7-1HOPO").Value) - 1)
        If (ActiveDocument.Variables("p7-1HOPO").Value Like "*,*") Then
           ActiveDocument.Variables("p7-1HOPO").Value = ActiveDocument.Variables("TckZpt7-1").Value & " п.п." & ActiveDocument.Variables("p7-1HOPO").Value & UF1.FNPHOPO.Value
        Else
           ActiveDocument.Variables("p7-1HOPO").Value = ActiveDocument.Variables("TckZpt7-1").Value & " п." & ActiveDocument.Variables("p7-1HOPO").Value & UF1.FNPHOPO.Value
        End If
    End If
    If UF1.CBho126.Value = True Then
        ActiveDocument.Variables("p7-1HOPOProdl").Value = ActiveDocument.Variables("TckZpt").Value & " п." & ActiveDocument.Variables("CBho126").Value & UF1.FNPHOPO.Value
    Else
        ActiveDocument.Variables("p7-1HOPOProdl").Value = ActiveDocument.Variables("TckZpt").Value & " п. 4" & UF1.FNPPP.Value
    End If
    SetVar = Array(135, 136, 137, 144, 145, 149, 150, 151, 152, 238, 240, 241, 242, 244, 247)
    Call SetValue(SetVar, "CBho", "p7-3HOPOProdl")
    ActiveDocument.Variables("p7-3HOPOProdl").Value = ActiveDocument.Variables("TckZpt").Value & " п.п." & ActiveDocument.Variables("p7-3HOPOProdl").Value & UF1.FNPHOPO.Value
    ActiveDocument.Variables("p7-4HOPO").Value = " п.п. " & ActiveDocument.Variables("CBho11").Value & ActiveDocument.Variables("CBho12").Value & ActiveDocument.Variables("CBho267").Value & UF1.FNPHOPO.Value
    ActiveDocument.Variables("p-8FNPHOPO").Value = ActiveDocument.Variables("TckZpt").Value & UF1.FNPHOPO.Value
    ActiveDocument.Variables("OsnRez-HOPO").Value = " п.п." & ActiveDocument.Variables("CBho11").Value & ActiveDocument.Variables("CBho12").Value & ActiveDocument.Variables("CBho255").Value & ActiveDocument.Variables("CBho267").Value & UF1.FNPHOPO.Value & ActiveDocument.Variables("TckZpt").Value
    ActiveDocument.Variables("NTDAktNKFNPHOPO").Value = "Федеральные нормы и правила в области промышленной безопасности " & Mid(UF1.FNPHOPO.Value, 64, 66) & "."
    If ActiveDocument.Variables("TckZpt").Value <> Strings.ChrW(31) Then ActiveDocument.Variables("NTDAktNKFNPHOPO").Value = Strings.Chr(13) & ActiveDocument.Variables("NTDAktNKFNPHOPO").Value
    ActiveDocument.Variables("TckZpt1").Value = ";" & Strings.Chr(13)
    ActiveDocument.Variables("TckZpt").Value = ";"
    ActiveDocument.Variables("TckZpt7-1").Value = ";"
    Flag = True
End If
If UF1.CBFNPPBETT.Value = True Then
    ActiveDocument.Variables("punkt1-1PBETT").Value = "п.п." 'Сборка пункта 1-1 для ТТ
    For Each mark In AllCBt
        If UF1.Controls.Item("CBtt" & mark).Value = True Then ActiveDocument.Variables("punkt1-1PBETT").Value = ActiveDocument.Variables("punkt1-1PBETT").Value & ActiveDocument.Variables("CBtt" & mark).Value
    Next
    ActiveDocument.Variables("punkt1-1PBETT").Value = ActiveDocument.Variables("TckZpt1").Value & Left(ActiveDocument.Variables("punkt1-1PBETT").Value, Len(ActiveDocument.Variables("punkt1-1PBETT").Value) - 1) & UF1.FNPTehnTrub.Value
    ActiveDocument.Variables("p-8FNPPBETT").Value = ActiveDocument.Variables("TckZpt").Value & UF1.FNPTehnTrub.Value
    SetVar = Array(27, 29, 35, 36, 59, 65, 85, 94, 100)
    ActiveDocument.Variables("p7-3PBETT").Value = ActiveDocument.Variables("TckZpt").Value & " п.п."
    Call SetValue(SetVar, "CBtt", "p7-3PBETT")
    ActiveDocument.Variables("p7-3PBETT").Value = ActiveDocument.Variables("p7-3PBETT").Value & UF1.FNPTehnTrub.Value
    SetVar = Array(141, 144, 145, 148, 164, 165, 166, 167, 168)
    ActiveDocument.Variables("GIFNPPBETT").Value = " п.п."
    Call SetValue(SetVar, "CBtt", "GIFNPPBETT")
    ActiveDocument.Variables("GIFNPPBETT").Value = ActiveDocument.Variables("GIFNPPBETT").Value & UF1.FNPTehnTrub.Value
    ActiveDocument.Variables("punkt7TTMPD").Value = " п. 191 подпункт " & Strings.Chr(171) & "а" & Strings.Chr(187) & UF1.FNPTehnTrub.Value
    ActiveDocument.Variables("punkt7TTRasch").Value = " соответствует требованиям п.п." & ActiveDocument.Variables("CBtt190").Value & ActiveDocument.Variables("CBtt191").Value & UF1.FNPTehnTrub.Value & " и"
    If ActiveDocument.Variables("CBtt25").Value <> Strings.ChrW(31) Then ActiveDocument.Variables("p7-1TT").Value = ActiveDocument.Variables("TckZpt7-1").Value & " п." & ActiveDocument.Variables("CBtt25").Value & UF1.FNPTehnTrub.Value
    If ActiveDocument.Variables("CBtt191").Value <> Strings.ChrW(31) Then ActiveDocument.Variables("p191PBETT").Value = " п." & ActiveDocument.Variables("CBtt191").Value & UF1.FNPTehnTrub.Value
    ActiveDocument.Variables("NTDAktNKFNPPBETT").Value = "Федеральные нормы и правила в области промышленной безопасности " & Mid(UF1.FNPTehnTrub.Value, 64, 63) & "."
    ActiveDocument.Variables("ZDPBTT").Value = " п. 191 подпункт " & Strings.Chr(171) & "а" & Strings.Chr(187) & UF1.FNPTehnTrub.Value
    If ActiveDocument.Variables("TckZpt").Value <> Strings.ChrW(31) Then ActiveDocument.Variables("NTDAktNKFNPPBETT").Value = Strings.Chr(13) & ActiveDocument.Variables("NTDAktNKFNPPBETT").Value
    ActiveDocument.Variables("TckZpt1").Value = ";" & Strings.Chr(13)
    ActiveDocument.Variables("TckZpt").Value = ";"
    Flag = True
End If
If UF1.CBFNPPBSNN.Value = True Then
    ActiveDocument.Variables("punkt1-1PBSNN").Value = "п.п." 'Сборка пункта 1-1 для Складов нефти и нефтепродуктов
    For Each mark In AllCBs
        If UF1.Controls.Item("CBsn" & mark).Value = True Then ActiveDocument.Variables("punkt1-1PBSNN").Value = ActiveDocument.Variables("punkt1-1PBSNN").Value & ActiveDocument.Variables("CBsn" & mark).Value
    Next
    ActiveDocument.Variables("punkt1-1PBSNN").Value = ActiveDocument.Variables("TckZpt1").Value & Left(ActiveDocument.Variables("punkt1-1PBSNN").Value, Len(ActiveDocument.Variables("punkt1-1PBSNN").Value) - 1) & UF1.FNPPESNN.Value
    ActiveDocument.Variables("OsnRez-PBSNN").Value = " п.п." & ActiveDocument.Variables("CBsn77").Value & ActiveDocument.Variables("CBsn81").Value & ActiveDocument.Variables("CBsn87").Value & ActiveDocument.Variables("CBsn94").Value & ActiveDocument.Variables("CBsn98").Value
    ActiveDocument.Variables("OsnRez-PBSNN").Value = ActiveDocument.Variables("OsnRez-PBSNN").Value & ActiveDocument.Variables("CBsn102").Value & ActiveDocument.Variables("CBsn104").Value & ActiveDocument.Variables("CBsn105").Value & UF1.FNPPESNN.Value
    SetVar = Array(137, 141, 148)
    ActiveDocument.Variables("p7-1SNN").Value = ActiveDocument.Variables("TckZpt7-1").Value & " п.п."
    Call SetValue(SetVar, "CBsn", "p7-1SNN")
    ActiveDocument.Variables("p7-1SNN").Value = ActiveDocument.Variables("p7-1SNN").Value & UF1.FNPPESNN.Value
    SetVar = Array(140, 142, 144, 146, 147, 149, 150, 151, 156, 157, 159, 160, 167, 168)
    ActiveDocument.Variables("p7-3PSNN").Value = ActiveDocument.Variables("TckZpt").Value & " п.п."
    Call SetValue(SetVar, "CBsn", "p7-3PSNN")
    ActiveDocument.Variables("p7-3PSNN").Value = ActiveDocument.Variables("p7-3PSNN").Value & UF1.FNPPESNN.Value
'    ActiveDocument.Variables("p7-3PSNN").Value = " п.п." & ActiveDocument.Variables("CBsn146").Value & ActiveDocument.Variables("CBsn147").Value & ActiveDocument.Variables("CBsn149").Value & ActiveDocument.Variables("CBsn150").Value & UF1.FNPPESNN.Value & ActiveDocument.Variables("TckZpt1").Value
    ActiveDocument.Variables("p-8FNPPSNN").Value = ActiveDocument.Variables("TckZpt").Value & UF1.FNPPESNN.Value
    ActiveDocument.Variables("NTDAktNKFNPPBSNN").Value = "Федеральные нормы и правила в области промышленной безопасности " & Mid(UF1.FNPPESNN.Value, 64, 66) & "."
    If ActiveDocument.Variables("TckZpt").Value <> Strings.ChrW(31) Then ActiveDocument.Variables("NTDAktNKFNPPBSNN").Value = Strings.Chr(13) & ActiveDocument.Variables("NTDAktNKFNPPBSNN").Value
    ActiveDocument.Variables("TckZpt1").Value = ";" & Strings.Chr(13)
    ActiveDocument.Variables("TckZpt").Value = ";"
    Flag = True
End If
If Flag = False Then 'Если не отмечен ни один пункт(устройство не подпадает под действие ФНП) оформление по общим пунктам
    ActiveDocument.Variables("p7-1ORPDProdl").Value = " п. 2 статьи 7, п. 1 статьи 13 Федерального закона " & Strings.Chr(171) & "О промышленной безопасности опасных производственных объектов" & Strings.Chr(187) & " от 21.07.1997 г. №116-ФЗ; "
    ActiveDocument.Variables("p7-1OPVBProdl").Value = "п. 4" & UF1.FNPPP.Value
    ActiveDocument.Variables("punkt1-1obsh").Value = ActiveDocument.Variables("p7-1ORPDProdl").Value & Strings.Chr(13) & "п.п. 4, 23, 25" & UF1.FNPPP.Value
    ActiveDocument.Variables("punkt7-3NTD").Value = " НТД"
    ActiveDocument.Variables("p-8FNPORPDR").Value = " НТД"
'    ActiveDocument.Bookmarks("R7p4").Range.Delete
End If

If UF1.CBPodNaliv.Value = True Then
    ActiveDocument.Variables("TimeGI").Value = "4 часов"
    ActiveDocument.Variables("PadDavl").Value = Strings.ChrW(31)
    ActiveDocument.Variables("PodRabDav").Value = Strings.ChrW(31)
    ActiveDocument.Variables("DavlNeVishe").Value = " без избыточного давления (под налив)"
    ActiveDocument.Variables("p2pr8RD3417302").Value = Strings.ChrW(31)
End If
If UF1.CBVakuum.Value = True Then
    ActiveDocument.Variables("IndxP").Value = "абс"
    ActiveDocument.Variables("VacuumP").Value = ", вакуум до " & Format(1 - CDbl(UF1.RazreshaemoeP.Value), "###0.0#####") & " кгс/см" & Strings.ChrW(178)
    ActiveDocument.Variables("RazreshaemoeVKM").Value = "вакуум до " & Format((1 - CDbl(UF1.RazreshaemoeP.Value)) / 10, "###0.0#####") & " (" & Format(1 - CDbl(UF1.RazreshaemoeP.Value), "###0.0#####") & ")"
    ActiveDocument.Variables("DavlNeVishe").Value = " без избыточного давления (вакуум до " & Format(1 - CDbl(UF1.RazreshaemoeP.Value), "###0.0#####") & " кгс/см" & Strings.ChrW(178) & ")"
End If
If UF1.CBp466.Value = True Then  'если активен пункт 466 то разрешаем до 30.09
    If (DateDiff("d", UF1.AktGID.Value, ActiveDocument.Variables("CBp466data").Value)) < 0 Then
        ActiveDocument.Variables("DoNgoda") = Format(DateAdd("yyyy", Val(UF1.NaNLet.Value), ActiveDocument.Variables("CBp466data").Value), "dd.mm.yyyy")
    Else
        ActiveDocument.Variables("DoNgoda") = Format(DateAdd("yyyy", Val(UF1.NaNLet.Value) - 1, ActiveDocument.Variables("CBp466data").Value), "dd.mm.yyyy")
    End If
End If
'Индивидуальные изменения для пунктов
'Пункт 7.4 удаляем если пустой
If ActiveDocument.Variables("p7-4ORPD").Value = Strings.ChrW(31) And ActiveDocument.Variables("p7-4HOPO").Value = Strings.ChrW(31) Then
    ActiveDocument.Variables("KIPiA").Value = Strings.ChrW(31)
Else
    If ActiveDocument.Variables("p7-4HOPO").Value <> Strings.ChrW(31) Then
        ActiveDocument.Variables("p7-4HOPO").Value = ActiveDocument.Variables("p7-4HOPO").Value & "."
    Else
        ActiveDocument.Variables("p7-4ORPD").Value = ActiveDocument.Variables("p7-4ORPD").Value & "."
    End If
End If
'Пункт 7.5.4
If UF1.CBGOST34347.Value = True Then
    ActiveDocument.Variables("punkt7-5-4Mat").Value = ActiveDocument.Variables("GOST34347PMat").Value
    ActiveDocument.Variables("TckZpt").Value = ";"
Else
    ActiveDocument.Variables("TckZpt").Value = Strings.ChrW(31)
End If
If UF1.CBFNPORPD.Value = True Then
    ActiveDocument.Variables("punkt7-5-4Mat").Value = ActiveDocument.Variables("punkt7-5-4Mat").Value & ActiveDocument.Variables("TckZpt").Value & ActiveDocument.Variables("p100FNPORPD").Value
    ActiveDocument.Variables("TckZpt").Value = ";"
End If
If UF1.CBFNPOPVB.Value = True Then
    ActiveDocument.Variables("punkt7-5-4Mat").Value = ActiveDocument.Variables("punkt7-5-4Mat").Value & ActiveDocument.Variables("TckZpt").Value & ActiveDocument.Variables("p7-1OPVB").Value
    ActiveDocument.Variables("TckZpt").Value = ";"
End If
If UF1.CBFNPPBETT.Value = True Then
    ActiveDocument.Variables("punkt7-5-4Mat").Value = ActiveDocument.Variables("punkt7-5-4Mat").Value & ActiveDocument.Variables("p7-1TT").Value
    ActiveDocument.Variables("TckZpt").Value = ";"
End If
If UF1.CBFNPPBSNN.Value = True And UF1.CBFNPPBETT.Value = True Then
    ActiveDocument.Variables("punkt7-5-4Mat").Value = ActiveDocument.Variables("punkt7-5-4Mat").Value & ActiveDocument.Variables("p7-1SNN").Value
    ActiveDocument.Variables("TckZpt").Value = ";"
End If
If Strings.Len(ActiveDocument.Variables("punkt7-5-4Mat").Value) > 10 Then
    ActiveDocument.Variables("punkt7-5-4Mat").Value = " (материалы " & ActiveDocument.Variables("TechUsrtva").Value & " соответствуют требованиям" & ActiveDocument.Variables("punkt7-5-4Mat").Value & ")"
Else
    ActiveDocument.Variables("punkt7-5-4Mat").Value = Strings.ChrW(31)
End If

If UF1.PnIs.Value = True Then 'добавляем давления испытаний для акустики
    If IsNumeric(UF1.ddlina.Value) Then ActiveDocument.Variables("RazmContZon").Value = Format((3.14 * CDbl(UF1.odiam.Value) * (CDbl(UF1.odlina.Value) + CDbl(UF1.odiam.Value) / 2) / 1000000), "###0.0")
    ActiveDocument.Variables("RazreshaemoeP0_5").Value = Format(CDbl(UF1.RazreshaemoeP.Value) * 0.5, "###0.0###")
    ActiveDocument.Variables("RazreshaemoeP0_5MP").Value = Format(CDbl(UF1.RazreshaemoeP.Value) / 10 * 0.5, "###0.0###")
    ActiveDocument.Variables("RazreshaemoeP0_75").Value = Format(CDbl(UF1.RazreshaemoeP.Value) * 0.75, "###0.0##")
    ActiveDocument.Variables("RazreshaemoeP0_75MP").Value = Format(CDbl(UF1.RazreshaemoeP.Value) / 10 * 0.75, "###0.0###")
    ActiveDocument.Variables("RazreshaemoeP0_25").Value = Format(CDbl(UF1.RazreshaemoeP.Value) * 0.25, "###0.0##")
End If

If UF1.CBRubashka.Value = True Or IsNull(UF1.CBRubashka.Value) Then
    ActiveDocument.Variables("RabSredaToplRasch").Value = "в корпусе: рабочая среда"
    ActiveDocument.Variables("RabSredaTopl").Value = "в корпусе: рабочая среда"
    ActiveDocument.Tables(2).Columns.Add
    ActiveDocument.Tables(2).AutoFitBehavior (wdAutoFitWindow)
    ActiveDocument.Tables(2).Columns(1).PreferredWidthType = wdPreferredWidthPercent
    ActiveDocument.Tables(2).Columns(1).PreferredWidth = 50
    ActiveDocument.Tables(2).Columns(2).PreferredWidthType = wdPreferredWidthPercent
    ActiveDocument.Tables(2).Columns(2).PreferredWidth = 25
    ActiveDocument.Tables(2).Columns(3).PreferredWidthType = wdPreferredWidthPercent
    ActiveDocument.Tables(2).Columns(3).PreferredWidth = 25
    ActiveDocument.Tables(2).Cell(Row:=1, Column:=2).Range = "Корпус"
    ActiveDocument.Tables(2).Cell(Row:=2, Column:=3).Range = Format(CDbl(UF1.RabocheePRub.Value) / 10, "###0.0#####") & " (" & Format(CDbl(UF1.RabocheePRub.Value), "###0.0#####") & ")"
    ActiveDocument.Tables(2).Cell(Row:=3, Column:=3).Range = UF1.RabTempRub.Value
    ActiveDocument.Tables(2).Cell(Row:=4, Column:=3).Range = UF1.RabSredaRub.Value
    If UF1.CBRubashka.Value = True Then
        ActiveDocument.Tables(2).Cell(Row:=1, Column:=3).Range = "Рубашка"
        ActiveDocument.Variables("VKorp").Value = ActiveDocument.Variables("VKorp").Value & "." & Strings.Chr(13) & "в рубашке: рабочая среда - " & UF1.RaschSredaRub.Value & _
        ", P=" & Format(UF1.RaschetnPRub.Value, "###0.0#####") & " кгс/см" & Strings.ChrW(178) & ", t=" & UF1.RaschetntRub.Value & Strings.ChrW(176) & "С" & ActiveDocument.Variables("VRub").Value
        ActiveDocument.Variables("ParamVRub").Value = "." & Strings.Chr(13) & "в рубашке: рабочая среда - " & UF1.RabSredaRub.Value & _
        ", P=" & Format(UF1.RabocheePRub.Value, "###0.0#####") & " кгс/см" & Strings.ChrW(178) & ", t=" & UF1.RabTempRub.Value & Strings.ChrW(176) & "С"
    Else
        ActiveDocument.Tables(2).Cell(Row:=1, Column:=3).Range = "Трубная система"
        ActiveDocument.Variables("VKorp").Value = ActiveDocument.Variables("VKorp").Value & "." & Strings.Chr(13) & "в трубной системе: рабочая среда - " & UF1.RaschSredaRub.Value & _
        ", P=" & Format(UF1.RaschetnPRub.Value, "###0.0#####") & " кгс/см" & Strings.ChrW(178) & ", t=" & UF1.RaschetntRub.Value & Strings.ChrW(176) & "С" & ActiveDocument.Variables("VRub").Value
        ActiveDocument.Variables("ParamVRub").Value = "." & Strings.Chr(13) & "в трубной системе: рабочая среда - " & UF1.RabSredaRub.Value & _
        ", P=" & Format(UF1.RabocheePRub.Value, "###0.0#####") & " кгс/см" & Strings.ChrW(178) & ", t=" & UF1.RabTempRub.Value & Strings.ChrW(176) & "С"
    End If
    If UF1.OptionTruboprovod.Value = True Then
        ActiveDocument.Tables(2).Cell(Row:=1, Column:=2).Range = "до РОУ"
        ActiveDocument.Tables(2).Cell(Row:=1, Column:=3).Range = "после РОУ"
        ActiveDocument.Variables("RabSredaToplRasch").Value = "до РОУ: рабочая среда"
        ActiveDocument.Variables("RabSredaTopl").Value = "до РОУ: рабочая среда"
        ActiveDocument.Variables("VKorp").Value = "." & Strings.Chr(13) & "после РОУ: " _
        & "P=" & Format(UF1.RaschetnPRub.Value, "###0.0#####") & " кгс/см" & Strings.ChrW(178) & ", t=" & UF1.RaschetntRub.Value & Strings.ChrW(176) & "С" & ActiveDocument.Variables("VRub").Value
        ActiveDocument.Variables("ParamVRub").Value = "." & Strings.Chr(13) & "после РОУ: " & _
        "P=" & Format(UF1.RabocheePRub.Value, "###0.0#####") & " кгс/см" & Strings.ChrW(178) & ", t=" & UF1.RabTempRub.Value & Strings.ChrW(176) & "С"
    End If
End If

If UF1.OptionTruboprovod.Value <> True Then ActiveDocument.Tables(3).Rows(2).Delete

If (ActiveDocument.Variables("TechUsrtvo").Value Like "*трубопров*") Then ActiveDocument.Variables("PasportPar").Value = "Расчетные (проектные) параметры"
If UF1.CBFNPORPD.Value = True Or UF1.CBFNPPBETT.Value = True Then ActiveDocument.Variables("punkt7MPD").Value = ";" & ActiveDocument.Variables("punkt7MPD").Value

End Sub

Sub variable()
'ActiveDocument.Variables("punkt1-1OPVB").Value = Strings.Chr(13) & "п.п. 4, 23, 25 Федеральных норм и правил в области промышленной безопасности " & Strings.Chr(171) & "Правила проведения экспертизы промышленной безопасности" & Strings.Chr(187) & ", утвержденных приказом Федеральной службы по экологическому, технологическому и атомному надзору от 20.10.2020 г. №420, зарегистрированных в Минюсте России рег.№61391 от 11.12.2020 г"
ActiveDocument.Variables("VIKRezKontr").Value = "при визуальном и измерительном контроле дефектов, препятствующих дальнейшей эксплуатации, не обнаружено" ' & Strings.Chr(171) & "Методические рекомендации о порядке проведения магнитопорошкового контроля технических устройств и сооружений, применяемых и эксплуатируемых на опасных производственных объектах" & Strings.Chr(187) & ", утвержденные Приказом Ростехнадзора от 13.12.2006 г. №1072" 'Strings.ChrW(31)
'MsgBox (ActiveDocument.Variables("CBp178").Value)
'    Dim V As variable, S As String
'    For Each V In ActiveDocument.Variables
'        S = V.Name & vbTab & V.Value & vbNewLine
'    MsgBox S
'    Next
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
'Sub SetBookmark(NameBookmarks As String, ValueBookmarks As String)
'    Set TTMP = ActiveDocument.Bookmarks(NameBookmarks).Range
'    TTMP.Text = ValueBookmarks
'    ActiveDocument.Bookmarks.Add Name:=NameBookmarks, Range:=TTMP
'End Sub

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

Sub TckZpt()
'Пункт 1-1
'If UF1.CBFNPHOPO.Value = True And UF1.CBFNPPBETT.Value = True Then 'Тогда добавляем пунктуацию в тех.трубопроводы
'    Call TckZptPBETT
'End If
'If UF1.CBFNPOPVB.Value = True Then 'Тогда добавляем пунктуацию в ХОПО и тех. трубопроводы
'    If UF1.CBFNPHOPO.Value = True Then
'        Call TckZptHOPO
'    Else
'        If UF1.CBFNPPBETT.Value = True Then Call TckZptPBETT
'    End If
'End If
'If UF1.CBFNPORPD.Value = True Then 'Тогда добавляем пунктуацию в ОПВБ, ХОПО и тех. трубопроводы
'    If UF1.CBFNPOPVB.Value = True Then
'        Call TckZptOPVB
'    Else
'        If UF1.CBFNPHOPO.Value = True Then
'            Call TckZptHOPO
'        Else
'            If UF1.CBFNPPBETT.Value = True Then Call TckZptPBETT
'        End If
'    End If
'End If
'Пункт эксперты давление и химия
'If UF1.ExpertORPD.Value = True And UF1.ExpertHim.Value = True Then ActiveDocument.Variables("UdostExpHim").Value = ";" & Strings.Chr(13) & ActiveDocument.Variables("UdostExpHim").Value
''Пункт 7-1
'If ActiveDocument.Variables("p7-1OPVB").Value <> Strings.ChrW(31) And ActiveDocument.Variables("p7-1HOPO").Value <> Strings.ChrW(31) Then ActiveDocument.Variables("p7-1HOPO").Value = ";" & ActiveDocument.Variables("p7-1HOPO").Value
'If ActiveDocument.Variables("GOST34347PMat").Value <> Strings.ChrW(31) Then
'    If ActiveDocument.Variables("p7-1OPVB").Value <> Strings.ChrW(31) Then
'        ActiveDocument.Variables("p7-1OPVB").Value = ";" & ActiveDocument.Variables("p7-1OPVB").Value
'    Else
'        If ActiveDocument.Variables("p7-1HOPO").Value <> Strings.ChrW(31) Then ActiveDocument.Variables("p7-1HOPO").Value = ";" & ActiveDocument.Variables("p7-1HOPO").Value
'    End If
'End If
End Sub

Sub TckZptPBETT()
'    ActiveDocument.Variables("punkt1-1PBETT").Value = ";" & Strings.Chr(13) & ActiveDocument.Variables("punkt1-1PBETT").Value
'    ActiveDocument.Variables("p-8FNPPBETT").Value = ";" & ActiveDocument.Variables("p-8FNPPBETT").Value
'    ActiveDocument.Variables("p7-3PBETT").Value = ";" & ActiveDocument.Variables("p7-3PBETT").Value
'    ActiveDocument.Variables("NTDAktNKFNPPBETT").Value = Strings.Chr(13) & ActiveDocument.Variables("NTDAktNKFNPPBETT").Value

End Sub
Sub TckZptHOPO()
'    ActiveDocument.Variables("punkt1-1HOPO").Value = ";" & Strings.Chr(13) & ActiveDocument.Variables("punkt1-1HOPO").Value
'    ActiveDocument.Variables("p-8FNPHOPO").Value = ";" & ActiveDocument.Variables("p-8FNPHOPO").Value
'    ActiveDocument.Variables("p7-1HOPOProdl").Value = ";" & ActiveDocument.Variables("p7-1HOPOProdl").Value
'    ActiveDocument.Variables("p7-3HOPOProdl").Value = ";" & ActiveDocument.Variables("p7-3HOPOProdl").Value
'    ActiveDocument.Variables("NTDAktNKFNPHOPO").Value = Strings.Chr(13) & ActiveDocument.Variables("NTDAktNKFNPHOPO").Value

End Sub

Sub TckZptOPVB()
'    ActiveDocument.Variables("punkt1-1OPVB").Value = ";" & Strings.Chr(13) & ActiveDocument.Variables("punkt1-1OPVB").Value
'    ActiveDocument.Variables("p-8FNPOPVB").Value = ";" & ActiveDocument.Variables("p-8FNPOPVB").Value
'    ActiveDocument.Variables("p7-1OPVBProdl").Value = ";" & ActiveDocument.Variables("p7-1OPVBProdl").Value
'    ActiveDocument.Variables("p7-3OPVB").Value = ";" & ActiveDocument.Variables("p7-3OPVB").Value
'    ActiveDocument.Variables("NTDAktNKFNPOPVB").Value = Strings.Chr(13) & ActiveDocument.Variables("NTDAktNKFNPOPVB").Value

End Sub

Sub SetValue(SetVal, tipFNP, var)

For Each mark In SetVal
    Dim var1 As String
    var1 = tipFNP & mark
    ActiveDocument.Variables(var).Value = ActiveDocument.Variables(var).Value & ActiveDocument.Variables(var1).Value
Next mark

End Sub

