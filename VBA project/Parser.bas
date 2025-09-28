Attribute VB_Name = "Parser"
Option Explicit

Function SearchAndCopy(objDoc As Object, str1 As String, str2 As String)
    Dim MyRange
    Dim rStart&, rEnd&
    Set MyRange = objDoc.Content
    With MyRange
        With .Find
            .ClearFormatting
            .Text = str1
            .Forward = True
            .Wrap = wdFindStop
            .Format = False
            .MatchCase = False
            .MatchWholeWord = True
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
            If .Execute = True Then
                rStart = MyRange.End
                rEnd = rStart
'                MsgBox (.Parent.Paragraphs(1).Next.Range.Text)
            End If
'            If .found Then rStart = MyRange.End: rEnd = rStart
        End With
    End With
    If str1 <> "Ѕ≈«ќѕј—Ќќ—“»" Then
        Set MyRange = objDoc.Range(Start:=rStart + 500)
        With MyRange
            With .Find
                .Text = str1
'                .Start = rStart + Len(str1)
                If .Execute = True Then
                    rStart = MyRange.End
                    rEnd = rStart
'                    MsgBox (.Parent.Paragraphs(1).Next.Range.Text)
                End If
            End With
        End With
    End If
    Set MyRange = objDoc.Range(Start:=rStart + Len(str1))
    With MyRange
        With .Find
            .Text = str2
            .Execute
            If .found Then rEnd = MyRange.Start
        End With
    End With
    If rEnd > rStart Then
        SearchAndCopy = objDoc.Range(rStart, rEnd).Text
'        xc = objWrd.Selection.Copy
    End If

End Function

Function Titul(ByVal tit As String)
Dim pozn, pozc
Dim tmp, regexpResult As String

    tit = Replace(tit, Chr(13), " ")
    tit = Replace(tit, "  ", " ")
    tit = Trim(tit)
    regexpResult = RegExpExtract(tit, "н*а*\s*техничес[а-€ј-я]+\s*устройст[а-€ј-я]+\s*[-,:]*") '»щем "на техническое устройство" с вариаци€ми
    If regexpResult <> "" Then tit = Replace(tit, regexpResult, "")
    regexpResult = RegExpExtract(tit, "н*а*\s*сооруж[а-€ј-я]+\s*[-,:]*") '»щем "на сооружение" с вариаци€ми
    If regexpResult <> "" Then tit = Replace(tit, regexpResult, "")
    regexpResult = RegExpExtract(tit, "н*а*\s*технол[а-€ј-я]+\s*трубопр[а-€ј-я]+\s*[-,:]*") '»щем "на технологический трубопровод" с вариаци€ми
    If regexpResult <> "" Then
        tit = Replace(tit, regexpResult, "")
        UF1.OptionSoorugenie.Value = True
        UF1.ComboBoxTipUstroistva.Value = "технологический трубопровод"
    End If
    regexpResult = RegExpExtract(tit, "[—с]ос[ј-яа-€]+[,]?\s+работающ[ј-яа-€]+\s+под\s+давлен[ј-яа-€]+") '»щем "сосуд работающий под давлением" с вариаци€ми
    If regexpResult <> "" Then
        tit = Replace(tit, regexpResult, "")
        UF1.OptionSosud.Value = True
        UF1.ComboBoxTipUstroistva.Value = "—осуд под давлением"
    End If
    tit = FirstLastCh(tit)
    regexpResult = RegExpExtract(tit, "(примен€[а-€ј-я]+\s)|(устано[а-€ј-я]+\s)") '»щем "примен€емое" или "установленное"
    If regexpResult <> "" Then
        pozn = InStr(tit, regexpResult)
        regexpResult = RegExpExtract(tit, "клас[а-€ј-я]*\s+опасн[а-€ј-я]+") '»щем "класс опасности"
        If regexpResult <> "" Then
            pozc = InStr(tit, regexpResult) + Len(regexpResult)
        Else
            regexpResult = RegExpExtract(tit, "[Aј]?\d{2}.\d{5}.\d{4}") '»щем рег.єќѕќ
            If regexpResult <> "" Then
                pozc = InStr(tit, regexpResult) + Len(regexpResult)
            Else
                pozc = 0
            End If
        End If
    End If
    If pozc > pozn Then
        tit = Left(tit, pozn - 1) & Right(tit, Len(tit) - pozc)
    End If
    tit = FirstLastCh(tit)
    regexpResult = RegExpExtract(tit, "[з«пѕр–][аое][вгз][.]\s*є") '»щем "зав. поз. рег. є"
    If regexpResult <> "" Then
        pozn = InStr(tit, regexpResult)
        tit = Left(tit, pozn - 1)
    End If
    regexpResult = RegExpExtract(tit, "^[з«пѕр–с—у”][чтаое][еавгз][тнои.][нсдц][а-€ј-я.]+\s*є") '»щем "зав. уч. поз. рег. ст. є" полные
    If regexpResult <> "" Then
        pozn = InStr(tit, regexpResult)
        tit = Left(tit, pozn - 1)
    End If
    regexpResult = RegExpExtract(tit, "[с—у”][чт][.]+\s*є") '»щем " уч. ст. є"
    If regexpResult <> "" Then
        pozn = InStr(tit, regexpResult)
        tit = Left(tit, pozn - 1)
    End If
    MsgBox (tit)
'    pozc = InStrRev(tit, "є", -1, vbBinaryCompare)
'    If pozc > 0 Then tit = Left(tit, pozc - 5)
'    pozn = InStr(tit, "примен")
'    If pozn > 0 Then
'        If regexpResult <> "" Then
'        pozc = InStr(pozn, tit, regexpResult, vbBinaryCompare)
'            If pozc < InStr(tit, "класс ") Then
'                pozc = InStr(tit, "класс ")
'                tmp = Mid(tit, pozc + 15)
'            Else
'                tmp = Mid(tit, pozc + Len(regexpResult))
'            End If
'        End If
'    tit = Left(tit, pozn - 1) & tmp
'    End If
'    If pozc + Len(regexpResult) < Len(tit) Then
    tit = FirstLastCh(tit)
    Titul = Trim(tit)
End Function

Function FirstLastCh(ByVal str As String)
Dim strt, nd, charCode
    str = Trim(str)
    If str = "" Then Exit Function
    strt = Left(str, 1)
    charCode = Asc(strt)
    Debug.Print "—имвол '" & strt & "' имеет код: " & charCode
    nd = Right(str, 1)
    Debug.Print "—имвол '" & nd & "' имеет код: " & Asc(nd)
    If strt = "," Or strt = "-" Or strt = ":" Or strt = Chr(150) Then
        If str <> "" Then str = Right(str, Len(str) - 1)
    End If
    If nd = "," Or nd = "-" Or nd = ":" Or nd = "." Or nd = ";" Then
        If str <> "" Then str = Left(str, Len(str) - 1)
    End If
    FirstLastCh = Trim(str)
End Function

Function FirstLastAbzaz(str As String)
Dim poz1, poz2
    poz1 = InStr(str, Chr(13))
'    MsgBox (poz1)
    poz2 = InStrRev(str, Chr(13))
'    MsgBox (poz2)
    FirstLastAbzaz = Mid(str, poz1 + 1, poz2 - poz1)
    Debug.Print (FirstLastAbzaz)
End Function

Public Function RegExpExtract(Text As String, Pattern As String, Optional item As Integer = 1) As String
    Dim regex, matches
    On Error GoTo ErrHandl
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = Pattern
    regex.Global = True
    If regex.Test(Text) Then
        Set matches = regex.Execute(Text)
        RegExpExtract = matches.item(item - 1)
        Exit Function
    End If
    Exit Function
ErrHandl:
    RegExpExtract = ""
End Function

Function SplitDocuments(str As String, ByRef BaseDocuments As Variant)
    Dim a() As String, regexpResult As String
    Dim i, n, x
    
    str = Trim(str)
    n = 1
    x = UBound(BaseDocuments, 1)
    a = Split(str, Strings.Chr(13))
    For i = 0 To UBound(a)
        regexpResult = RegExpExtract(a(i), "проведе.+\s+эксперт.+\s+рас.+:")
        If regexpResult <> "" Then a(i) = ""
        regexpResult = RegExpExtract(a(i), "свидетельс.+\s+о\s+регистр.+\s+опасн[а-€ј-я]+\s+")
        If regexpResult <> "" Then a(i) = ""
        regexpResult = RegExpExtract(a(i), "справ[а-€ј-я]+\s+о\s+режим[а-€ј-я]+\s+рабо[а-€ј-я]+")
        If regexpResult <> "" Then a(i) = ""
        regexpResult = RegExpExtract(a(i), "технич[а-€ј-я]+\s+диагностир[а-€ј-я,]+\s+нераз[а-€ј-я]+")
        If regexpResult <> "" Then a(i) = ""
    Next i
    For i = 0 To UBound(a)
        a(i) = Replace(a(i), Chr(13), "")
        a(i) = FirstLastCh(a(i))
        If a(i) <> "" And regexpResult = "" Then
            BaseDocuments(x, n) = a(i)
'            MsgBox (a(i))
            n = n + 1
        End If
    Next i

End Function

Function SplitSvedeniya(ByVal str As String)
    Dim regexpResult As String
    Dim i, n, x, pozn, pozc
    Dim flag As Boolean
    Dim coll As New Collection
    
    flag = True
    str = Trim(str)
    pozn = 1
    pozc = 1
    
    Do
        regexpResult = RegExpExtract(str, "[ј-я][а-€][а-€()\-,\s]+:")
        If regexpResult <> "" Then
'            MsgBox (regexpResult)
            n = InStr(str, regexpResult)
            If n < 10 Then
                pozn = n
                regexpResult = RegExpExtract(str, "[ј-я][а-€][а-€()\-,\s]+:", 2)
                If regexpResult <> "" Then
                    pozc = InStr(str, regexpResult)
                Else
                    pozc = 1
                End If
            Else
                pozn = 1
                pozc = n
            End If
            If pozn < pozc Then
                regexpResult = Mid(str, pozn, pozc - pozn)
'                MsgBox (regexpResult)
                coll.Add regexpResult
                str = Replace(str, regexpResult, "")
            End If
        End If
        If regexpResult = "" Then flag = False
    Loop While (flag)
'    MsgBox (str)
    coll.Add str
    Set SplitSvedeniya = coll
    
'    n = 1
'    x = UBound(BaseDocuments, 1)
'    a = Split(str, Strings.Chr(13))
'    For i = 0 To UBound(a)
'    Next i


End Function

Function CleanResult(str As String, regexpResult As String)
    Dim i
    str = Replace(str, regexpResult, "")
    str = Replace(str, Chr(13), " ")
    str = Replace(str, Chr(9), " ")
    str = Replace(str, Chr(7), " ")
    For i = 1 To 5
        str = Replace(str, "  ", " ")
    Next i
    str = FirstLastCh(str)
    CleanResult = str
End Function

Function ExtractSvedeniya(ByVal coll As Collection)

Dim arr(1 To 12, 1 To 2) As Variant
Dim regexpResult As String
Dim i, n, str As String
Dim tmp As New Collection
Dim flag As Boolean

flag = True

    arr(1, 1) = "Ќазначение"
    arr(1, 2) = "[нЌ]азначени.+:"
    arr(2, 1) = "изготовитель"
    arr(2, 2) = ".*[и»]зготовитель.*:"
    arr(3, 1) = "ћонтажна€ организаци€"
    arr(3, 2) = ".*[мћ]онтажна€.*:"
    arr(4, 1) = "владельцы и дата перемещени€"
    arr(4, 2) = ".*перемещени€.*:"
    arr(5, 1) = "рабоча€ среда"
    arr(5, 2) = ".*[с—]реда.*:"
    arr(6, 1) = "¬ид топлива"
    arr(6, 2) = ".*топлив.*:"
    arr(7, 1) = "«аводской номер"
    arr(7, 2) = ".*[з«]аводск[ио][ей].*:"
    arr(8, 1) = "—пособ соединени€ элементов"
    arr(8, 2) = ".*соединени€.*:"
    arr(9, 1) = "ѕримененные сварочные материалы"
    arr(9, 2) = ".*[пѕ]римененные.*:"
    arr(10, 1) = "—ведени€  о неразрушающем контроле"
    arr(10, 2) = ".*заводс.*[нЌдƒ][е][рф][еа][кз].*:"
    arr(11, 1) = "—ведени€  о неразрушающем контроле"
    arr(11, 2) = ".*[нЌдƒ][е][рф][еа][кз].*изготов.*:"
    arr(12, 1) = " оличество циклов"
    arr(12, 2) = ".*циклов.*:"
    
    For i = 1 To coll.Count
        For n = 1 To UBound(arr, 1)
            str = Replace(Trim$(coll(i)), Chr(13), "")
            regexpResult = RegExpExtract(str, Trim$(arr(n, 2)))
            If regexpResult <> "" Then
                str = CleanResult(str, regexpResult)
                flag = False
                If n = 1 Then UF1.NaznTehUstr.Value = str
                If n = 2 Then UF1.ZavodIzg.Value = str
                If n = 3 Then UF1.MontagOrg.Value = str
                If n = 4 Then UF1.Vladelez.Value = str
                If n = 5 Or n = 6 Then UF1.RaschSreda.Value = str
                If n = 7 Then UF1.ZavN.Value = str
                If n = 8 Then UF1.FlanzSoed.Value = str
                If n = 9 Then UF1.PrimSvMat.Value = str
                If n = 10 Or n = 11 Then UF1.ZavKontr.Value = str
                If n = 12 Then UF1.KolZicl.Value = str
                MsgBox (str)
            End If
        Next n
        If flag Then
            tmp.Add coll(i)
        End If
        flag = True
    Next i
Set ExtractSvedeniya = tmp

End Function

Function ExtractSvedeniyaData(ByVal coll As Collection)

Dim arr(1 To 3, 1 To 2) As Variant
Dim regexpResult As String
Dim i, n, str As String
Dim tmp As New Collection
Dim flag As Boolean

flag = True

    arr(1, 1) = "дата изготовлени€"
    arr(1, 2) = ".*[и»]зготовлени€.*:"
    arr(2, 1) = "дата монтажа"
    arr(2, 2) = ".*[мћ]онтажа.*:"
    arr(3, 1) = "дата ввода"
    arr(3, 2) = ".*[в¬]вод.*:"
    
    For i = 1 To coll.Count
        For n = 1 To UBound(arr, 1)
            str = Replace(Trim$(coll(i)), Chr(13), "")
            regexpResult = RegExpExtract(str, Trim$(arr(n, 2)))
            If regexpResult <> "" Then
                str = CleanResult(str, regexpResult)
                flag = False
                regexpResult = RegExpExtract(str, ".*[12][09][0-9][0-9]")
                If regexpResult <> "" Then
                    If n = 1 Then UF1.DataIzg.Value = regexpResult
                    If n = 2 Then UF1.DataMontaga.Value = regexpResult
                    If n = 3 Then UF1.DataVvoda.Value = regexpResult
                Else
                    If n = 1 Then UF1.DataIzg.Value = str
                    If n = 2 Then UF1.DataMontaga.Value = str
                    If n = 3 Then UF1.DataVvoda.Value = str
                End If
                MsgBox (regexpResult)
            End If
        Next n
        If flag Then
            tmp.Add coll(i)
        End If
        flag = True
    Next i
Set ExtractSvedeniyaData = tmp

End Function


Function ExtractSvedeniyaRegN(ByVal coll As Collection)
Dim arr(1 To 4, 1 To 2) As Variant
Dim regexpResult As String
Dim i, n, str As String
Dim tmp As New Collection
Dim flag As Boolean
Dim poz

flag = True
poz = 0

    arr(1, 1) = "–егистрационный номер"
    arr(1, 2) = ".*[р–]егистрационный.*:"
    arr(2, 1) = "”четный номер"
    arr(2, 2) = ".*[у”]четный.*:"
    arr(3, 1) = "—танционный номер"
    arr(3, 2) = ".*[с—]танционный.*:"
    arr(4, 1) = "ѕозици€"
    arr(4, 2) = ".*[пѕ]озици€.*:"
    
    For i = 1 To coll.Count
        For n = 1 To UBound(arr, 1)
            str = Replace(Trim$(coll(i)), Chr(13), "")
            regexpResult = RegExpExtract(str, Trim$(arr(n, 2)))
            If regexpResult <> "" Then
                str = CleanResult(str, regexpResult)
                flag = False
                poz = InStr(str, ",")
'                If poz > 0 Then
'                    UF1.DataRegistracii.Value = Right(str, Len(str) - poz)
'                    MsgBox (Right(str, Len(str) - poz))
'                    str = Left(str, poz - 1)
'                End If
                If n = 1 Or n = 2 Then
                    If poz > 0 Then
                        UF1.DataRegistracii.Value = Right(str, Len(str) - poz)
                        MsgBox (Right(str, Len(str) - poz))
                        str = Left(str, poz - 1)
                    End If
                    UF1.RegN.Value = str
                End If
                If n = 1 Then UF1.poleRegNum.Value = "рег.є"
                If n = 2 Then UF1.poleRegNum.Value = "уч.є"
                If n = 3 Then
                    UF1.poleRegNum.Value = "ст.є"
                    UF1.Position.Value = str
                End If
                If n = 4 Then
                    UF1.poleRegNum.Value = "поз.є"
                    UF1.Position.Value = str
                End If
                MsgBox (str)
            End If
        Next n
        If flag Then
            tmp.Add coll(i)
        End If
        flag = True
    Next i
Set ExtractSvedeniyaRegN = tmp

End Function

Function ExtractSvedeniyaRaschParam(ByVal coll As Collection)
Dim arr(1 To 2, 1 To 2) As Variant
Dim regexpResult As String
Dim i, n, str As String
Dim tmp As New Collection
Dim flag As Boolean
Dim tmp1 As String
flag = True

    arr(1, 1) = "–асчетные параметры"
    arr(1, 2) = ".*[р–]асчетн.+:"
    arr(2, 1) = "–абочие параметры"
    arr(2, 2) = ".*[р–]а[зб][ро][еч].+:"
    
    For i = 1 To coll.Count
        For n = 1 To UBound(arr, 1)
            str = Replace(Trim$(coll(i)), Chr(13), "")
            regexpResult = RegExpExtract(str, Trim$(arr(n, 2)))
            If regexpResult <> "" Then
                str = CleanResult(str, regexpResult)
                flag = False
                regexpResult = RegExpExtract(str, "[0-9,]+\s*[к мћ][пѕг][ајс]")
                If regexpResult <> "" Then
                    regexpResult = RegExpExtract(regexpResult, "[0-9,]+")
                    If n = 1 Then UF1.RaschetnP.Value = regexpResult
                    If n = 2 Then UF1.RazreshaemoeP.Value = regexpResult
                    MsgBox (regexpResult)
                Else
                    regexpResult = RegExpExtract(regexpResult, "налив")
                    If regexpResult <> "" Then UF1.CBPodNaliv.Value = True
                    regexpResult = RegExpExtract(regexpResult, "вакуум")
                    If regexpResult <> "" Then UF1.CBVakuum.Value = True
                End If
                tmp1 = "[0-9,]+\s*[мћ][3" & Strings.ChrW(179) & "]"
                Debug.Print tmp1
                regexpResult = RegExpExtract(str, tmp1)
                If regexpResult <> "" Then
                    regexpResult = RegExpExtract(regexpResult, "[0-9,]+")
                    UF1.VKorp.Value = regexpResult
                    MsgBox (regexpResult)
                End If
                tmp1 = "[+-]*[0-9,]*[" & Strings.ChrW(247) & "]*[+-]*[0-9,∞]+[\s]*[0∞" & Strings.Chr(40) & "]*[\s]*[cCс—]"
                Debug.Print tmp1
                regexpResult = RegExpExtract(str, tmp1)
                If regexpResult <> "" Then
                    tmp1 = "[+-]*[0-9,]*[" & Strings.ChrW(247) & "]*[+-]*[0-9,]+"
                    regexpResult = RegExpExtract(regexpResult, tmp1)
                    If n = 1 Then UF1.Raschetnt.Value = regexpResult
                    If n = 2 Then UF1.RabTemp.Value = regexpResult
                    MsgBox (regexpResult)
                End If
            End If
        Next n
        If flag Then
            tmp.Add coll(i)
        End If
        flag = True
    Next i
Set ExtractSvedeniyaRaschParam = tmp

End Function

Function ExtractElements(objDoc As Object)
Dim table As table
Dim str As String
Dim i, n


For n = 1 To 2
If objDoc.Tables.Count > n Then Set table = objDoc.Tables(n)
    For i = 1 To table.Columns.Count
        str = str & table.Rows(1).Cells(i).Range
    Next i
    
    If str Like "*олщина*" Then
        Call ExtractElementsTable(table)
    Else
        Set table = Nothing
    End If
    str = ""
Next n

End Function

Function ExtractElementsTable(table As table)
Dim i, n, z
Dim str As String

UF1.SpButElm.Value = 0

For i = 2 To table.Rows.Count
    If i > 11 Then UF1.SpButElm.Value = UF1.SpButElm.Value + 1
    z = -1 - UF1.SpButElm.Value
    If i > 2 Then UF1.Controls.item("CBR" & i + z).Value = True
    For n = 1 To table.Columns.Count
        str = table.Rows(1).Cells(n).Range
        If str Like "*аименова*" Then UF1.Controls.item("R" & i + z & "C1").Value = CleanResult(table.Rows(i).Cells(n).Range, "")
        If str Like "*лина*" Or str Like "*ысота*" Then UF1.Controls.item("R" & i + z & "C2").Value = CleanResult(table.Rows(i).Cells(n).Range, "")
        If str Like "*иаметр*" Then UF1.Controls.item("R" & i + z & "C3").Value = CleanResult(table.Rows(i).Cells(n).Range, "")
        If str Like "*олщи*" Then UF1.Controls.item("R" & i + z & "C4").Value = CleanResult(table.Rows(i).Cells(n).Range, "")
        If str Like "*тали*" Then UF1.Controls.item("R" & i + z & "C6").Value = CleanResult(table.Rows(i).Cells(n).Range, "")
        If str Like "*√ќ—“*" Or str Like "*“”*" Then UF1.Controls.item("R" & i + z & "C7").Value = CleanResult(table.Rows(i).Cells(n).Range, "")
    Next n
Next i
'UF1.Controls.Item("R" & i & "C" & n - 1).Value
End Function

Function RecRem(coll As Collection)
Dim arr(1 To 3, 1 To 2) As Variant
Dim i, n
Dim tmp As New Collection, regexpResult As String, str As String
Dim flag As Boolean
flag = True

    arr(1, 1) = "—ведени€ о реконструкци€х и ремонтах"
    arr(1, 2) = ".*реконструкци.+:"
    arr(2, 1) = "—ведени€ об экспертизах"
    arr(2, 2) = ".*экспертизах[ а-€]*:"
    arr(3, 1) = "—ведени€ о дефектоскопическом контроле"
    arr(3, 2) = ".*[нЌдƒ][е][рф][еа][кз].*монтаж.*:"
    
    For i = 1 To coll.Count
        For n = 1 To UBound(arr, 1)
            str = coll(i)
            regexpResult = RegExpExtract(str, Trim$(arr(n, 2)))
            If regexpResult <> "" Then
                str = Replace(str, regexpResult, "")
'                a = Split(str, Strings.Chr(13))
                If n = 1 Then Call SvedRecRem(str, "SvedORemonte", "SpButRemont")
                If n = 2 Or n = 3 Then Call SvedRecRem(str, "SvedOEPB", "SpButEPB")
                flag = False
            End If
        Next n
        If flag Then
            tmp.Add coll(i)
        End If
        flag = True
    Next i
Set RecRem = tmp

End Function

Function SvedRecRem(str As String, item As String, item1 As String)
Dim i, n, a() As String, tmp As String
    
    a = Split(str, Strings.Chr(13))
    UF1.Controls.item(item1).Value = 1
    
    For i = 0 To UBound(a)
        tmp = CleanResult(a(i), "")
        If tmp <> "" Then
            UF1.Controls.item(item).Value = tmp
            UF1.Controls.item(item1).Value = UF1.Controls.item(item1).Value + 1
            MsgBox (a(i))
        End If
    Next i
    
End Function

Function DeviceDefinition(coll As Collection)
Dim regexpResult As String
    
    regexpResult = RegExpExtract(coll(1), "н*а*\s*сооруж[а-€ј-я]+\s*[-,:]*") '»щем "на сооружение" с вариаци€ми
    If regexpResult <> "" Then
        UF1.OptionSoorugenie.Value = True
        If coll(4) Like "*трубопровода*" Then UF1.ComboBoxTipUstroistva.Value = "технологический трубопровод"
        If coll(4) Like "*резервуара*" Or coll(4) Like "*бака*" Then UF1.ComboBoxTipUstroistva.Value = "резервуар"
    End If
    
    regexpResult = RegExpExtract(coll(1), "н*а*\s*техничес[а-€ј-я]+\s*устройст[а-€ј-я]+\s*[-,:]*") '»щем "на техническое устройство" с вариаци€ми
    If regexpResult <> "" Then
        If coll(4) Like "*котла*" Or coll(4) Like "*экономайзера*" Then
            UF1.OptionKotel.Value = True
            If UF1.NazvTehUstr.Value Like "*экономайзер*" Then UF1.ComboBoxTipUstroistva.Value = "экономайзер"
            If UF1.NazvTehUstr.Value Like "*котел*" Then
                If UF1.NaznTehUstr.Value Like "*пара*" Then UF1.ComboBoxTipUstroistva.Value = "паровой котел"
                If UF1.NaznTehUstr.Value Like "*воды*" Then UF1.ComboBoxTipUstroistva.Value = "водогрейный котел"
            End If
        End If
        If coll(4) Like "*сосуда*" Or coll(4) Like "*газификатора*" Or coll(4) Like "*теплообменника*" Or coll(4) Like "*подогревател€*" _
        Or coll(4) Like "*баллон*" Or coll(4) Like "*воздухосборник*" Or coll(4) Like "*автоклава*" Or coll(4) Like "*газификатора*" Or coll(4) Like "*бака*" Then
            UF1.OptionSosud.Value = True
            If UF1.NazvTehUstr.Value Like "*газификатор*" Then UF1.ComboBoxTipUstroistva.Value = "√азификатор √’ "
            If UF1.NazvTehUstr.Value Like "*теплообменник*" Or UF1.NaznTehUstr.Value Like "*подогревател*" Then UF1.ComboBoxTipUstroistva.Value = "ѕодогреватель"
            If UF1.NazvTehUstr.Value Like "*баллон*" Then UF1.ComboBoxTipUstroistva.Value = "Ѕаллон"
            If UF1.NazvTehUstr.Value Like "*баллоны*" Then UF1.ComboBoxTipUstroistva.Value = "Ѕаллоны групповой установки"
        End If
        If coll(4) Like "*трубопровода*" Then
            UF1.OptionTruboprovod.Value = True
            If UF1.NaznTehUstr.Value Like "*пара*" Then UF1.ComboBoxTipUstroistva.Value = "трубопровод пара"
            If UF1.NaznTehUstr.Value Like "*воды*" Then UF1.ComboBoxTipUstroistva.Value = "трубопровод гор€чей воды"
        End If
        If coll(4) Like "*насоса*" Or coll(4) Like "*компрессора*" Then
            UF1.OptionOstalnoe.Value = True
            If UF1.NazvTehUstr.Value Like "*насос*" Then UF1.ComboBoxTipUstroistva.Value = "Ќасос"
            If UF1.NazvTehUstr.Value Like "*компрессор*" Then UF1.ComboBoxTipUstroistva.Value = " омпрессор"
        End If
    End If
End Function
