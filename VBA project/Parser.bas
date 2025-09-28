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
    If str1 <> "������������" Then
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
    regexpResult = RegExpExtract(tit, "�*�*\s*��������[�-��-�]+\s*��������[�-��-�]+\s*[-,:]*") '���� "�� ����������� ����������" � ����������
    If regexpResult <> "" Then tit = Replace(tit, regexpResult, "")
    regexpResult = RegExpExtract(tit, "�*�*\s*������[�-��-�]+\s*[-,:]*") '���� "�� ����������" � ����������
    If regexpResult <> "" Then tit = Replace(tit, regexpResult, "")
    regexpResult = RegExpExtract(tit, "�*�*\s*������[�-��-�]+\s*�������[�-��-�]+\s*[-,:]*") '���� "�� ��������������� �����������" � ����������
    If regexpResult <> "" Then
        tit = Replace(tit, regexpResult, "")
        UF1.OptionSoorugenie.Value = True
        UF1.ComboBoxTipUstroistva.Value = "��������������� �����������"
    End If
    regexpResult = RegExpExtract(tit, "[��]��[�-��-�]+[,]?\s+��������[�-��-�]+\s+���\s+������[�-��-�]+") '���� "����� ���������� ��� ���������" � ����������
    If regexpResult <> "" Then
        tit = Replace(tit, regexpResult, "")
        UF1.OptionSosud.Value = True
        UF1.ComboBoxTipUstroistva.Value = "����� ��� ���������"
    End If
    tit = FirstLastCh(tit)
    regexpResult = RegExpExtract(tit, "(�������[�-��-�]+\s)|(������[�-��-�]+\s)") '���� "�����������" ��� "�������������"
    If regexpResult <> "" Then
        pozn = InStr(tit, regexpResult)
        regexpResult = RegExpExtract(tit, "����[�-��-�]*\s+�����[�-��-�]+") '���� "����� ���������"
        If regexpResult <> "" Then
            pozc = InStr(tit, regexpResult) + Len(regexpResult)
        Else
            regexpResult = RegExpExtract(tit, "[A�]?\d{2}.\d{5}.\d{4}") '���� ���.����
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
    regexpResult = RegExpExtract(tit, "[������][���][���][.]\s*�") '���� "���. ���. ���. �"
    If regexpResult <> "" Then
        pozn = InStr(tit, regexpResult)
        tit = Left(tit, pozn - 1)
    End If
    regexpResult = RegExpExtract(tit, "^[����������][�����][�����][����.][����][�-��-�.]+\s*�") '���� "���. ��. ���. ���. ��. �" ������
    If regexpResult <> "" Then
        pozn = InStr(tit, regexpResult)
        tit = Left(tit, pozn - 1)
    End If
    regexpResult = RegExpExtract(tit, "[����][��][.]+\s*�") '���� " ��. ��. �"
    If regexpResult <> "" Then
        pozn = InStr(tit, regexpResult)
        tit = Left(tit, pozn - 1)
    End If
    MsgBox (tit)
'    pozc = InStrRev(tit, "�", -1, vbBinaryCompare)
'    If pozc > 0 Then tit = Left(tit, pozc - 5)
'    pozn = InStr(tit, "������")
'    If pozn > 0 Then
'        If regexpResult <> "" Then
'        pozc = InStr(pozn, tit, regexpResult, vbBinaryCompare)
'            If pozc < InStr(tit, "����� ") Then
'                pozc = InStr(tit, "����� ")
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
    Debug.Print "������ '" & strt & "' ����� ���: " & charCode
    nd = Right(str, 1)
    Debug.Print "������ '" & nd & "' ����� ���: " & Asc(nd)
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
        regexpResult = RegExpExtract(a(i), "�������.+\s+�������.+\s+���.+:")
        If regexpResult <> "" Then a(i) = ""
        regexpResult = RegExpExtract(a(i), "����������.+\s+�\s+�������.+\s+�����[�-��-�]+\s+")
        If regexpResult <> "" Then a(i) = ""
        regexpResult = RegExpExtract(a(i), "�����[�-��-�]+\s+�\s+�����[�-��-�]+\s+����[�-��-�]+")
        If regexpResult <> "" Then a(i) = ""
        regexpResult = RegExpExtract(a(i), "������[�-��-�]+\s+����������[�-��-�,]+\s+�����[�-��-�]+")
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
        regexpResult = RegExpExtract(str, "[�-�][�-�][�-�()\-,\s]+:")
        If regexpResult <> "" Then
'            MsgBox (regexpResult)
            n = InStr(str, regexpResult)
            If n < 10 Then
                pozn = n
                regexpResult = RegExpExtract(str, "[�-�][�-�][�-�()\-,\s]+:", 2)
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

    arr(1, 1) = "����������"
    arr(1, 2) = "[��]��������.+:"
    arr(2, 1) = "������������"
    arr(2, 2) = ".*[��]�����������.*:"
    arr(3, 1) = "��������� �����������"
    arr(3, 2) = ".*[��]��������.*:"
    arr(4, 1) = "��������� � ���� �����������"
    arr(4, 2) = ".*�����������.*:"
    arr(5, 1) = "������� �����"
    arr(5, 2) = ".*[��]����.*:"
    arr(6, 1) = "��� �������"
    arr(6, 2) = ".*������.*:"
    arr(7, 1) = "��������� �����"
    arr(7, 2) = ".*[��]������[��][��].*:"
    arr(8, 1) = "������ ���������� ���������"
    arr(8, 2) = ".*����������.*:"
    arr(9, 1) = "����������� ��������� ���������"
    arr(9, 2) = ".*[��]����������.*:"
    arr(10, 1) = "��������  � ������������� ��������"
    arr(10, 2) = ".*������.*[����][�][��][��][��].*:"
    arr(11, 1) = "��������  � ������������� ��������"
    arr(11, 2) = ".*[����][�][��][��][��].*�������.*:"
    arr(12, 1) = "���������� ������"
    arr(12, 2) = ".*������.*:"
    
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

    arr(1, 1) = "���� ������������"
    arr(1, 2) = ".*[��]�����������.*:"
    arr(2, 1) = "���� �������"
    arr(2, 2) = ".*[��]������.*:"
    arr(3, 1) = "���� �����"
    arr(3, 2) = ".*[��]���.*:"
    
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

    arr(1, 1) = "��������������� �����"
    arr(1, 2) = ".*[��]��������������.*:"
    arr(2, 1) = "������� �����"
    arr(2, 2) = ".*[��]������.*:"
    arr(3, 1) = "����������� �����"
    arr(3, 2) = ".*[��]����������.*:"
    arr(4, 1) = "�������"
    arr(4, 2) = ".*[��]������.*:"
    
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
                If n = 1 Then UF1.poleRegNum.Value = "���.�"
                If n = 2 Then UF1.poleRegNum.Value = "��.�"
                If n = 3 Then
                    UF1.poleRegNum.Value = "��.�"
                    UF1.Position.Value = str
                End If
                If n = 4 Then
                    UF1.poleRegNum.Value = "���.�"
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

    arr(1, 1) = "��������� ���������"
    arr(1, 2) = ".*[��]������.+:"
    arr(2, 1) = "������� ���������"
    arr(2, 2) = ".*[��]�[��][��][��].+:"
    
    For i = 1 To coll.Count
        For n = 1 To UBound(arr, 1)
            str = Replace(Trim$(coll(i)), Chr(13), "")
            regexpResult = RegExpExtract(str, Trim$(arr(n, 2)))
            If regexpResult <> "" Then
                str = CleanResult(str, regexpResult)
                flag = False
                regexpResult = RegExpExtract(str, "[0-9,]+\s*[����][���][���]")
                If regexpResult <> "" Then
                    regexpResult = RegExpExtract(regexpResult, "[0-9,]+")
                    If n = 1 Then UF1.RaschetnP.Value = regexpResult
                    If n = 2 Then UF1.RazreshaemoeP.Value = regexpResult
                    MsgBox (regexpResult)
                Else
                    regexpResult = RegExpExtract(regexpResult, "�����")
                    If regexpResult <> "" Then UF1.CBPodNaliv.Value = True
                    regexpResult = RegExpExtract(regexpResult, "������")
                    If regexpResult <> "" Then UF1.CBVakuum.Value = True
                End If
                tmp1 = "[0-9,]+\s*[��][3" & Strings.ChrW(179) & "]"
                Debug.Print tmp1
                regexpResult = RegExpExtract(str, tmp1)
                If regexpResult <> "" Then
                    regexpResult = RegExpExtract(regexpResult, "[0-9,]+")
                    UF1.VKorp.Value = regexpResult
                    MsgBox (regexpResult)
                End If
                tmp1 = "[+-]*[0-9,]*[" & Strings.ChrW(247) & "]*[+-]*[0-9,�]+[\s]*[0�" & Strings.Chr(40) & "]*[\s]*[cC��]"
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
    
    If str Like "*������*" Then
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
        If str Like "*��������*" Then UF1.Controls.item("R" & i + z & "C1").Value = CleanResult(table.Rows(i).Cells(n).Range, "")
        If str Like "*����*" Or str Like "*�����*" Then UF1.Controls.item("R" & i + z & "C2").Value = CleanResult(table.Rows(i).Cells(n).Range, "")
        If str Like "*������*" Then UF1.Controls.item("R" & i + z & "C3").Value = CleanResult(table.Rows(i).Cells(n).Range, "")
        If str Like "*����*" Then UF1.Controls.item("R" & i + z & "C4").Value = CleanResult(table.Rows(i).Cells(n).Range, "")
        If str Like "*����*" Then UF1.Controls.item("R" & i + z & "C6").Value = CleanResult(table.Rows(i).Cells(n).Range, "")
        If str Like "*����*" Or str Like "*��*" Then UF1.Controls.item("R" & i + z & "C7").Value = CleanResult(table.Rows(i).Cells(n).Range, "")
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

    arr(1, 1) = "�������� � �������������� � ��������"
    arr(1, 2) = ".*������������.+:"
    arr(2, 1) = "�������� �� �����������"
    arr(2, 2) = ".*�����������[ �-�]*:"
    arr(3, 1) = "�������� � ������������������ ��������"
    arr(3, 2) = ".*[����][�][��][��][��].*������.*:"
    
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
    
    regexpResult = RegExpExtract(coll(1), "�*�*\s*������[�-��-�]+\s*[-,:]*") '���� "�� ����������" � ����������
    If regexpResult <> "" Then
        UF1.OptionSoorugenie.Value = True
        If coll(4) Like "*������������*" Then UF1.ComboBoxTipUstroistva.Value = "��������������� �����������"
        If coll(4) Like "*����������*" Or coll(4) Like "*����*" Then UF1.ComboBoxTipUstroistva.Value = "���������"
    End If
    
    regexpResult = RegExpExtract(coll(1), "�*�*\s*��������[�-��-�]+\s*��������[�-��-�]+\s*[-,:]*") '���� "�� ����������� ����������" � ����������
    If regexpResult <> "" Then
        If coll(4) Like "*�����*" Or coll(4) Like "*������������*" Then
            UF1.OptionKotel.Value = True
            If UF1.NazvTehUstr.Value Like "*�����������*" Then UF1.ComboBoxTipUstroistva.Value = "�����������"
            If UF1.NazvTehUstr.Value Like "*�����*" Then
                If UF1.NaznTehUstr.Value Like "*����*" Then UF1.ComboBoxTipUstroistva.Value = "������� �����"
                If UF1.NaznTehUstr.Value Like "*����*" Then UF1.ComboBoxTipUstroistva.Value = "����������� �����"
            End If
        End If
        If coll(4) Like "*������*" Or coll(4) Like "*������������*" Or coll(4) Like "*��������������*" Or coll(4) Like "*�������������*" _
        Or coll(4) Like "*������*" Or coll(4) Like "*��������������*" Or coll(4) Like "*���������*" Or coll(4) Like "*������������*" Or coll(4) Like "*����*" Then
            UF1.OptionSosud.Value = True
            If UF1.NazvTehUstr.Value Like "*�����������*" Then UF1.ComboBoxTipUstroistva.Value = "����������� ���"
            If UF1.NazvTehUstr.Value Like "*�������������*" Or UF1.NaznTehUstr.Value Like "*������������*" Then UF1.ComboBoxTipUstroistva.Value = "�������������"
            If UF1.NazvTehUstr.Value Like "*������*" Then UF1.ComboBoxTipUstroistva.Value = "������"
            If UF1.NazvTehUstr.Value Like "*�������*" Then UF1.ComboBoxTipUstroistva.Value = "������� ��������� ���������"
        End If
        If coll(4) Like "*������������*" Then
            UF1.OptionTruboprovod.Value = True
            If UF1.NaznTehUstr.Value Like "*����*" Then UF1.ComboBoxTipUstroistva.Value = "����������� ����"
            If UF1.NaznTehUstr.Value Like "*����*" Then UF1.ComboBoxTipUstroistva.Value = "����������� ������� ����"
        End If
        If coll(4) Like "*������*" Or coll(4) Like "*�����������*" Then
            UF1.OptionOstalnoe.Value = True
            If UF1.NazvTehUstr.Value Like "*�����*" Then UF1.ComboBoxTipUstroistva.Value = "�����"
            If UF1.NazvTehUstr.Value Like "*����������*" Then UF1.ComboBoxTipUstroistva.Value = "����������"
        End If
    End If
End Function
