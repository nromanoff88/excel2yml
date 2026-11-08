Attribute VB_Name = "modYamlCodec"
Option Explicit

' =============================================================================
' Модуль: modYamlCodec
' Назначение: Чтение/запись UTF-8 (без BOM), парсинг и сериализация YAML-файлов
'             конфигурации таблиц ODS (в соответствии со schema_ods_table.json).
' =============================================================================

Public Const ARRAY_ITEM_DELIMITER As String = "---"

' Чтение текстового файла в кодировке UTF-8 (корректно обрабатывает файлы с BOM и без BOM)
Public Function ReadUtf8File(ByVal filePath As String) As String
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2 ' adTypeText
    stm.Charset = "utf-8"
    stm.Open
    stm.LoadFromFile filePath
    ReadUtf8File = stm.ReadText(-1) ' adReadAll
    stm.Close
    Set stm = Nothing
End Function

' Запись текстового файла в кодировке UTF-8 БЕЗ BOM с переводами строк CRLF
Public Sub WriteUtf8NoBomFile(ByVal filePath As String, ByVal textContent As String)
    Dim normalized As String
    normalized = NormalizeLineEndingsLf(textContent)
    normalized = Replace(normalized, vbLf, vbCrLf)

    Dim stmText As Object
    Dim stmBin As Object
    Set stmText = CreateObject("ADODB.Stream")
    stmText.Type = 2 ' adTypeText
    stmText.Charset = "utf-8"
    stmText.Open
    stmText.WriteText normalized

    stmText.Position = 0
    stmText.Type = 1 ' adTypeBinary

    ' Пропускаем 3 байта UTF-8 BOM (EF BB BF), которые добавляет ADODB.Stream
    If stmText.Size >= 3 Then
        Dim headerBytes As Variant
        headerBytes = stmText.Read(3)
        If Not IsEmpty(headerBytes) Then
            If UBound(headerBytes) >= 2 Then
                If headerBytes(0) = &HEF And headerBytes(1) = &HBB And headerBytes(2) = &HBF Then
                    stmText.Position = 3
                Else
                    stmText.Position = 0
                End If
            Else
                stmText.Position = 0
            End If
        End If
    Else
        stmText.Position = 0
    End If

    Set stmBin = CreateObject("ADODB.Stream")
    stmBin.Type = 1 ' adTypeBinary
    stmBin.Open
    If stmText.Position < stmText.Size Then
        stmText.CopyTo stmBin
    End If
    stmBin.SaveToFile filePath, 2 ' adSaveCreateOverWrite

    stmBin.Close
    stmText.Close
    Set stmBin = Nothing
    Set stmText = Nothing
End Sub

' Нормализация всех переводов строк к LF (vbLf)
Public Function NormalizeLineEndingsLf(ByVal s As String) As String
    Dim res As String
    res = Replace(s, vbCrLf, vbLf)
    res = Replace(res, vbCr, vbLf)
    NormalizeLineEndingsLf = res
End Function

' Создание словаря таблицы со значениями по умолчанию
Public Function CreateDefaultTableDict(ByVal tableName As String) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1 ' vbTextCompare

    d("table_name") = tableName
    d("schedule_cron") = ""
    d("enabled_flag") = "true"

    d("source.connector_name") = ""
    d("source.connector_table_schema") = ""
    d("source.connector_table_name") = tableName
    d("source.connector_read_parallel_number") = "1"

    d("capture.data_capture_type") = "full"
    d("capture.pattern_capture") = ""
    d("capture.hard_delete_source_flag") = "false"
    d("capture.soft_delete_clause_trino") = ""
    d("capture.where_clause_trino") = ""
    d("capture.watermark_clause_trino") = ""
    d("capture.order_latest_clause_trino") = ""
    d("capture.cross_join_lateral_selects_trino") = ""

    d("watermark.select_wm_max_timestamp_sql") = ""
    d("watermark.select_wm_min_timestamp_sql") = ""
    d("watermark.select_wm_max_id_sql") = ""
    d("watermark.select_wm_min_id_sql") = ""
    d("watermark.select_wm_target_timestamp_sql") = ""
    d("watermark.select_wm_target_id_sql") = ""

    d("table.comment") = ""
    d("table.partitions_count") = "1"
    d("table.export_ahd_gp") = "latest"

    Dim cols As Collection
    Set cols = New Collection
    Set d("columns") = cols

    Set CreateDefaultTableDict = d
End Function

' Создание словаря колонки со значениями по умолчанию
Public Function CreateDefaultColumnDict(ByVal tableName As String) As Object
    Dim c As Object
    Set c = CreateObject("Scripting.Dictionary")
    c.CompareMode = 1 ' vbTextCompare

    c("table_name") = tableName
    c("ods_column_name") = ""
    c("ods_column_type") = "varchar"
    c("connector_mapping_sql") = ""
    c("rkh_flag") = "false"
    c("enabled_flag") = "true"
    c("comment") = ""

    Set CreateDefaultColumnDict = c
End Function

' =============================================================================
' ПАРСЕР YAML ФАЙЛА ТАБЛИЦЫ ODS
' =============================================================================
Public Function ParseOdsYamlFile(ByVal filePath As String, ByVal tableName As String) As Object
    Dim rawText As String
    rawText = ReadUtf8File(filePath)
    Set ParseOdsYamlFile = ParseOdsYamlText(rawText, tableName)
End Function

' Чтение YAML-файла, нормализация к стандартному YAML (key: val) и запись без BOM
Public Function NormalizeOdsYamlFile(ByVal srcPath As String, ByVal tableName As String, ByVal dstPath As String) As Boolean
    Dim tbl As Object
    Set tbl = ParseOdsYamlFile(srcPath, tableName)
    Dim yml1 As String
    yml1 = SerializeOdsYaml(tbl, tbl("columns"))
    WriteUtf8NoBomFile dstPath, yml1

    Dim tbl2 As Object
    Set tbl2 = ParseOdsYamlFile(dstPath, tableName)
    Dim yml2 As String
    yml2 = SerializeOdsYaml(tbl2, tbl2("columns"))
    NormalizeOdsYamlFile = (StrComp(yml1, yml2, vbBinaryCompare) = 0)
End Function

Public Function ParseOdsYamlText(ByVal rawText As String, ByVal tableName As String) As Object
    Dim tbl As Object
    Set tbl = CreateDefaultTableDict(tableName)
    ' Сбрасываем дефолты на пустые перед заполнением из файла
    tbl("source.connector_table_name") = ""
    tbl("table.export_ahd_gp") = ""

    Dim norm As String
    norm = NormalizeLineEndingsLf(rawText)
    If Len(Trim$(norm)) = 0 Then
        Set ParseOdsYamlText = tbl
        Exit Function
    End If

    Dim lines() As String
    lines = Split(norm, vbLf)

    Dim n As Long
    n = UBound(lines)

    Dim i As Long
    i = 0
    Do While i <= n
        Dim line As String
        line = lines(i)

        If Not IsBlankOrComment(line) Then
            Dim indent As Long
            indent = GetLineIndent(line)

            Dim key As String, rawVal As String
            If indent = 0 And TrySplitKeyValue(Trim$(line), key, rawVal) Then
                Dim cleanHeader As String
                cleanHeader = Trim$(StripInlineComment(rawVal))

                Select Case LCase$(key)
                    Case "schedule_cron"
                        If IsBlockScalarIndicator(cleanHeader) Then
                            tbl("schedule_cron") = ReadBlockScalar(lines, i, indent, cleanHeader)
                        Else
                            tbl("schedule_cron") = ParseScalarToCell(rawVal)
                        End If

                    Case "enabled_flag"
                        tbl("enabled_flag") = NormalizeBoolStr(ParseScalarToCell(rawVal), "true")

                    Case "source", "capture", "watermark", "table"
                        ParseYamlSection lines, i, n, LCase$(key), indent, tbl

                    Case "columns"
                        If cleanHeader = "[]" Then
                            ' Пустой массив колонок
                        Else
                            ParseColumnsSection lines, i, n, indent, tableName, tbl("columns")
                        End If

                    Case Else
                        ' Пропускаем неизвестную секцию верхнего уровня, если вдруг встретится
                        If IsBlockScalarIndicator(cleanHeader) Then
                            Call ReadBlockScalar(lines, i, indent, cleanHeader)
                        End If
                End Select
            End If
        End If

        i = i + 1
    Loop

    Set ParseOdsYamlText = tbl
End Function

' Парсинг вложенных объектных секций: source, capture, watermark, table
Private Sub ParseYamlSection(ByRef lines() As String, ByRef idx As Long, ByVal maxIdx As Long, _
                             ByVal sectionName As String, ByVal sectionIndent As Long, ByVal tbl As Object)
    Dim j As Long
    j = idx + 1

    Do While j <= maxIdx
        Dim ln As String
        ln = lines(j)

        If Not IsBlankOrComment(ln) Then
            Dim ind As Long
            ind = GetLineIndent(ln)
            If ind <= sectionIndent Then
                idx = j - 1
                Exit Sub
            End If

            Dim subKey As String, subRaw As String
            If TrySplitKeyValue(Trim$(ln), subKey, subRaw) Then
                Dim fullKey As String
                fullKey = sectionName & "." & LCase$(subKey)

                Dim cleanHdr As String
                cleanHdr = Trim$(StripInlineComment(subRaw))

                If fullKey = "capture.cross_join_lateral_selects_trino" Then
                    tbl(fullKey) = ParseYamlArrayField(lines, j, maxIdx, ind, subRaw)
                ElseIf IsBlockScalarIndicator(cleanHdr) Then
                    tbl(fullKey) = ReadBlockScalar(lines, j, ind, cleanHdr)
                Else
                    Dim parsedVal As String
                    parsedVal = ParseScalarToCell(subRaw)
                    If fullKey = "capture.hard_delete_source_flag" Then
                        parsedVal = NormalizeBoolStr(parsedVal, "false")
                    End If
                    tbl(fullKey) = parsedVal
                End If
            End If
        End If

        j = j + 1
    Loop

    idx = maxIdx
End Sub

' Парсинг массива строк (например, capture.cross_join_lateral_selects_trino)
Private Function ParseYamlArrayField(ByRef lines() As String, ByRef idx As Long, ByVal maxIdx As Long, _
                                     ByVal keyIndent As Long, ByVal rawVal As String) As String
    Dim cleanVal As String
    cleanVal = Trim$(StripInlineComment(rawVal))

    If Len(cleanVal) > 0 Then
        If LCase$(cleanVal) = "null" Or cleanVal = "~" Then
            ParseYamlArrayField = ""
            Exit Function
        ElseIf cleanVal = "[]" Then
            ParseYamlArrayField = "[]"
            Exit Function
        ElseIf Left$(cleanVal, 1) = "[" And Right$(cleanVal, 1) = "]" Then
            ParseYamlArrayField = ParseInlineYamlArray(Mid$(cleanVal, 2, Len(cleanVal) - 2))
            Exit Function
        ElseIf IsBlockScalarIndicator(cleanVal) Then
            ParseYamlArrayField = ReadBlockScalar(lines, idx, keyIndent, cleanVal)
            Exit Function
        Else
            ParseYamlArrayField = ParseScalarToCell(rawVal)
            Exit Function
        End If
    End If

    ' Многострочный YAML-список (- item)
    Dim items As Collection
    Set items = New Collection

    Dim j As Long
    j = idx + 1

    Do While j <= maxIdx
        Dim ln As String
        ln = lines(j)

        If Not IsBlankOrComment(ln) Then
            Dim ind As Long
            ind = GetLineIndent(ln)
            Dim trimmed As String
            trimmed = Trim$(ln)

            If ind < keyIndent Then
                idx = j - 1
                Exit Do
            End If
            If ind = keyIndent And Not IsSequenceItemLine(trimmed) Then
                idx = j - 1
                Exit Do
            End If

            If IsSequenceItemLine(trimmed) Then
                Dim afterDash As String
                afterDash = Trim$(Mid$(trimmed, 2))
                Dim cleanItemHdr As String
                cleanItemHdr = Trim$(StripInlineComment(afterDash))

                If IsBlockScalarIndicator(cleanItemHdr) Then
                    items.Add ReadBlockScalar(lines, j, ind, cleanItemHdr)
                Else
                    items.Add ParseScalarToCell(afterDash)
                End If
            End If
        End If

        j = j + 1
    Loop
    If j > maxIdx Then idx = maxIdx

    If items.Count = 0 Then
        ParseYamlArrayField = ""
    Else
        Dim arr() As String
        ReDim arr(1 To items.Count)
        Dim k As Long
        For k = 1 To items.Count
            arr(k) = items(k)
        Next k
        ParseYamlArrayField = Join(arr, vbLf & ARRAY_ITEM_DELIMITER & vbLf)
    End If
End Function

' Разбор компактного массива вида ['expr1', 'expr2']
Private Function ParseInlineYamlArray(ByVal inner As String) As String
    Dim items As Collection
    Set items = New Collection

    Dim cur As String
    Dim inSingle As Boolean, inDouble As Boolean
    Dim pos As Long, ch As String

    For pos = 1 To Len(inner)
        ch = Mid$(inner, pos, 1)
        If ch = "'" And Not inDouble Then
            inSingle = Not inSingle
            cur = cur & ch
        ElseIf ch = """" And Not inSingle Then
            inDouble = Not inDouble
            cur = cur & ch
        ElseIf ch = "," And Not inSingle And Not inDouble Then
            items.Add ParseScalarToCell(Trim$(cur))
            cur = ""
        Else
            cur = cur & ch
        End If
    Next pos

    If Len(Trim$(cur)) > 0 Then
        items.Add ParseScalarToCell(Trim$(cur))
    End If

    If items.Count = 0 Then
        ParseInlineYamlArray = "[]"
    Else
        Dim res() As String
        ReDim res(1 To items.Count)
        Dim m As Long
        For m = 1 To items.Count
            res(m) = items(m)
        Next m
        ParseInlineYamlArray = Join(res, vbLf & ARRAY_ITEM_DELIMITER & vbLf)
    End If
End Function

' Парсинг секции columns (массив объектов колонок)
Private Sub ParseColumnsSection(ByRef lines() As String, ByRef idx As Long, ByVal maxIdx As Long, _
                                ByVal sectionIndent As Long, ByVal tableName As String, ByVal colsColl As Collection)
    Dim j As Long
    j = idx + 1

    Do While j <= maxIdx
        Dim ln As String
        ln = lines(j)

        If Not IsBlankOrComment(ln) Then
            Dim ind As Long
            ind = GetLineIndent(ln)
            Dim trimmed As String
            trimmed = Trim$(ln)

            If ind < sectionIndent Then
                idx = j - 1
                Exit Sub
            End If
            If ind = sectionIndent And Not IsSequenceItemLine(trimmed) Then
                idx = j - 1
                Exit Sub
            End If

            If IsSequenceItemLine(trimmed) Then
                Dim dashIndent As Long
                dashIndent = ind

                Dim colDict As Object
                Set colDict = CreateDefaultColumnDict(tableName)
                colDict("ods_column_type") = ""

                ' Первое свойство может быть на той же строке после "- "
                Dim restAfterDash As String
                restAfterDash = LTrim$(Mid$(trimmed, 2))
                If Len(restAfterDash) > 0 Then
                    Dim firstPropIndent As Long
                    firstPropIndent = InStr(1, ln, restAfterDash, vbBinaryCompare) - 1
                    If firstPropIndent < dashIndent + 2 Then firstPropIndent = dashIndent + 2

                    Dim cKey As String, cRaw As String
                    If TrySplitKeyValue(restAfterDash, cKey, cRaw) Then
                        AssignColumnProperty lines, j, firstPropIndent, cKey, cRaw, colDict
                    End If
                End If

                ' Читаем остальные свойства этого элемента массива колонок
                j = j + 1
                Do While j <= maxIdx
                    Dim subLn As String
                    subLn = lines(j)
                    If Not IsBlankOrComment(subLn) Then
                        Dim subInd As Long
                        subInd = GetLineIndent(subLn)
                        Dim subTrim As String
                        subTrim = Trim$(subLn)

                        If subInd < dashIndent Then
                            Exit Do
                        End If
                        If subInd = dashIndent And IsSequenceItemLine(subTrim) Then
                            Exit Do
                        End If
                        If subInd <= sectionIndent And Not IsSequenceItemLine(subTrim) Then
                            Exit Do
                        End If

                        Dim propKey As String, propRaw As String
                        If TrySplitKeyValue(subTrim, propKey, propRaw) Then
                            AssignColumnProperty lines, j, subInd, propKey, propRaw, colDict
                        End If
                    End If
                    j = j + 1
                Loop

                colsColl.Add colDict
                j = j - 1 ' Компенсация инкремента внешнего цикла
            End If
        End If

        j = j + 1
    Loop

    idx = maxIdx
End Sub

Private Sub AssignColumnProperty(ByRef lines() As String, ByRef idx As Long, ByVal propIndent As Long, _
                                 ByVal propKey As String, ByVal propRaw As String, ByVal colDict As Object)
    Dim kLower As String
    kLower = LCase$(propKey)

    Dim cleanHdr As String
    cleanHdr = Trim$(StripInlineComment(propRaw))

    Dim val As String
    If IsBlockScalarIndicator(cleanHdr) Then
        val = ReadBlockScalar(lines, idx, propIndent, cleanHdr)
    Else
        val = ParseScalarToCell(propRaw)
    End If

    Select Case kLower
        Case "ods_column_name", "ods_column_type", "connector_mapping_sql", "comment"
            colDict(kLower) = val
        Case "rkh_flag"
            colDict("rkh_flag") = NormalizeBoolStr(val, "false")
        Case "enabled_flag"
            colDict("enabled_flag") = NormalizeBoolStr(val, "true")
    End Select
End Sub

' Чтение блочного скаляра YAML (|, |-, |+, >, >-)
' Включает защиту на случай сбитого отступа (как в примере table1.yml, где текст блока на уровне ключа)
Private Function ReadBlockScalar(ByRef lines() As String, ByRef idx As Long, _
                                 ByVal parentIndent As Long, ByVal header As String) As String
    Dim maxIdx As Long
    maxIdx = UBound(lines)

    Dim firstContentIdx As Long
    firstContentIdx = -1

    Dim probe As Long
    For probe = idx + 1 To maxIdx
        If Len(Trim$(lines(probe))) > 0 Then
            firstContentIdx = probe
            Exit For
        End If
    Next probe

    If firstContentIdx = -1 Then
        idx = maxIdx
        ReadBlockScalar = ""
        Exit Function
    End If

    Dim firstInd As Long
    firstInd = GetLineIndent(lines(firstContentIdx))

    Dim collected As Collection
    Set collected = New Collection
    Dim blockIndent As Long
    Dim k As Long

    If firstInd > parentIndent Then
        ' Стандартный YAML: все строки блока имеют отступ строго больше отступа родительского ключа
        blockIndent = firstInd
        ' Определяем минимальный непустой отступ внутри блока (если первая строка имела доп. отступ)
        For k = idx + 1 To maxIdx
            If Len(Trim$(lines(k))) > 0 Then
                Dim curInd As Long
                curInd = GetLineIndent(lines(k))
                If curInd <= parentIndent Then Exit For
                If curInd < blockIndent Then blockIndent = curInd
            End If
        Next k

        For k = idx + 1 To maxIdx
            If Len(Trim$(lines(k))) = 0 Then
                collected.Add ""
            Else
                If GetLineIndent(lines(k)) <= parentIndent Then
                    Exit For
                End If
                collected.Add StripIndentSpaces(lines(k), blockIndent)
            End If
        Next k
        idx = k - 1

    ElseIf firstInd = parentIndent And Not IsYamlKeyOrSequenceLine(lines(firstContentIdx)) Then
        ' Fallback для файлов со съехавшим отступом после |- (например, table1.yml)
        blockIndent = parentIndent
        For k = firstContentIdx To maxIdx
            If Len(Trim$(lines(k))) = 0 Then
                Exit For
            End If
            If GetLineIndent(lines(k)) <> parentIndent Or IsYamlKeyOrSequenceLine(lines(k)) Then
                Exit For
            End If
            collected.Add StripIndentSpaces(lines(k), blockIndent)
        Next k
        idx = k - 1
    Else
        ' Пустой блочный скаляр
        ReadBlockScalar = ""
        Exit Function
    End If

    ' Удаляем хвостовые пустые строки (для |- и стандартного представления в ячейке Excel)
    Dim lastNonEmpty As Long
    lastNonEmpty = collected.Count
    Do While lastNonEmpty > 0
        If Len(collected(lastNonEmpty)) = 0 Then
            lastNonEmpty = lastNonEmpty - 1
        Else
            Exit Do
        End If
    Loop

    If lastNonEmpty = 0 Then
        ReadBlockScalar = ""
        Exit Function
    End If

    Dim outArr() As String
    ReDim outArr(1 To lastNonEmpty)
    Dim m As Long
    For m = 1 To lastNonEmpty
        outArr(m) = collected(m)
    Next m

    If Left$(header, 1) = ">" Then
        ' Folded scalar: одиночные переносы сворачиваются в пробелы
        ReadBlockScalar = Join(outArr, " ")
    Else
        ReadBlockScalar = Join(outArr, vbLf)
    End If
End Function
