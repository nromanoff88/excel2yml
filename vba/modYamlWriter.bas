Attribute VB_Name = "modYamlWriter"
Option Explicit

' =============================================================================
' Модуль: modYamlWriter
' Назначение: Сериализация таблиц Excel в стандартный YAML (key: value)
'             и общие скалярные/лексические функции разбора YAML.
' =============================================================================

' =============================================================================
' СЕРИАЛИЗАЦИЯ В СТАНДАРТНЫЙ YAML (key: val, без пробела перед двоеточием)
' =============================================================================
Public Function SerializeOdsYaml(ByVal tblDict As Object, ByVal colsColl As Collection) As String
    Dim lines As Collection
    Set lines = New Collection

    lines.Add "# $schema: ../schema_ods_table.json"
    lines.Add "schedule_cron: " & FormatNullableSingleQuoted(CStr(tblDict("schedule_cron")))
    lines.Add "enabled_flag: " & NormalizeBoolStr(CStr(tblDict("enabled_flag")), "true")
    lines.Add ""

    lines.Add "source:"
    lines.Add "  connector_name: " & FormatSingleQuoted(CStr(tblDict("source.connector_name")))
    lines.Add "  connector_table_schema: " & FormatSingleQuoted(CStr(tblDict("source.connector_table_schema")))
    lines.Add "  connector_table_name: " & FormatSingleQuoted(CStr(tblDict("source.connector_table_name")))
    lines.Add "  connector_read_parallel_number: " & FormatIntStr(CStr(tblDict("source.connector_read_parallel_number")), 1)
    lines.Add ""

    lines.Add "capture:"
    lines.Add "  data_capture_type: " & FormatPlainOrDefault(CStr(tblDict("capture.data_capture_type")), "full")
    AppendNullableTextOrBlock lines, "  pattern_capture", CStr(tblDict("capture.pattern_capture")), 4, False
    lines.Add "  hard_delete_source_flag: " & NormalizeBoolStr(CStr(tblDict("capture.hard_delete_source_flag")), "false")
    AppendNullableBlockSql lines, "  soft_delete_clause_trino", CStr(tblDict("capture.soft_delete_clause_trino")), 4
    AppendNullableBlockSql lines, "  where_clause_trino", CStr(tblDict("capture.where_clause_trino")), 4
    AppendNullableBlockSql lines, "  watermark_clause_trino", CStr(tblDict("capture.watermark_clause_trino")), 4
    AppendNullableBlockSql lines, "  order_latest_clause_trino", CStr(tblDict("capture.order_latest_clause_trino")), 4
    AppendCrossJoinArray lines, "  cross_join_lateral_selects_trino", CStr(tblDict("capture.cross_join_lateral_selects_trino")), 4
    lines.Add ""

    lines.Add "watermark:"
    AppendNullableBlockSql lines, "  select_wm_max_timestamp_sql", CStr(tblDict("watermark.select_wm_max_timestamp_sql")), 4
    AppendNullableBlockSql lines, "  select_wm_min_timestamp_sql", CStr(tblDict("watermark.select_wm_min_timestamp_sql")), 4
    AppendNullableBlockSql lines, "  select_wm_max_id_sql", CStr(tblDict("watermark.select_wm_max_id_sql")), 4
    AppendNullableBlockSql lines, "  select_wm_min_id_sql", CStr(tblDict("watermark.select_wm_min_id_sql")), 4
    AppendNullableBlockSql lines, "  select_wm_target_timestamp_sql", CStr(tblDict("watermark.select_wm_target_timestamp_sql")), 4
    AppendNullableBlockSql lines, "  select_wm_target_id_sql", CStr(tblDict("watermark.select_wm_target_id_sql")), 4
    lines.Add ""

    lines.Add "table:"
    AppendCommentField lines, "  comment", CStr(tblDict("table.comment")), 4, True
    lines.Add "  partitions_count: " & FormatIntStr(CStr(tblDict("table.partitions_count")), 1)
    lines.Add "  export_ahd_gp: " & FormatNullablePlain(CStr(tblDict("table.export_ahd_gp")))
    lines.Add ""

    If colsColl Is Nothing Then
        lines.Add "columns: []"
    ElseIf colsColl.Count = 0 Then
        lines.Add "columns: []"
    Else
        lines.Add "columns:"
        Dim idx As Long
        For idx = 1 To colsColl.Count
            Dim c As Object
            Set c = colsColl(idx)

            lines.Add "  - ods_column_name: " & FormatIdentifierOrQuoted(CStr(c("ods_column_name")))
            lines.Add "    ods_column_type: " & FormatPlainOrDefault(CStr(c("ods_column_type")), "varchar")

            Dim mapSql As String
            mapSql = NormalizeLineEndingsLf(CStr(c("connector_mapping_sql")))
            If Len(Trim$(mapSql)) = 0 Then
                lines.Add "    connector_mapping_sql: ''"
            Else
                lines.Add "    connector_mapping_sql: |-"
                AppendIndentedMultiline lines, mapSql, 6
            End If

            lines.Add "    rkh_flag: " & NormalizeBoolStr(CStr(c("rkh_flag")), "false")
            lines.Add "    enabled_flag: " & NormalizeBoolStr(CStr(c("enabled_flag")), "true")
            AppendCommentField lines, "    comment", CStr(c("comment")), 6, False

            If idx < colsColl.Count Then
                lines.Add ""
            End If
        Next idx
    End If

    Dim resultArr() As String
    ReDim resultArr(1 To lines.Count)
    Dim p As Long
    For p = 1 To lines.Count
        resultArr(p) = lines(p)
    Next p

    SerializeOdsYaml = Join(resultArr, vbLf)
End Function

' =============================================================================
' ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ФОРМАТИРОВАНИЯ И РАЗБОРА СКАЛЯРОВ
' =============================================================================

Private Sub AppendNullableBlockSql(ByVal lines As Collection, ByVal keyWithIndent As String, _
                                   ByVal rawVal As String, ByVal blockSpaces As Long)
    Dim norm As String
    norm = NormalizeLineEndingsLf(rawVal)
    If IsNullToken(norm) Then
        lines.Add keyWithIndent & ": null"
    Else
        lines.Add keyWithIndent & ": |-"
        AppendIndentedMultiline lines, norm, blockSpaces
    End If
End Sub

Private Sub AppendNullableTextOrBlock(ByVal lines As Collection, ByVal keyWithIndent As String, _
                                      ByVal rawVal As String, ByVal blockSpaces As Long, ByVal alwaysBlock As Boolean)
    Dim norm As String
    norm = NormalizeLineEndingsLf(rawVal)
    If IsNullToken(norm) Then
        lines.Add keyWithIndent & ": null"
    ElseIf alwaysBlock Or InStr(1, norm, vbLf, vbBinaryCompare) > 0 Then
        lines.Add keyWithIndent & ": |-"
        AppendIndentedMultiline lines, norm, blockSpaces
    Else
        lines.Add keyWithIndent & ": " & FormatSingleQuoted(norm)
    End If
End Sub

Private Sub AppendCommentField(ByVal lines As Collection, ByVal keyWithIndent As String, _
                               ByVal rawVal As String, ByVal blockSpaces As Long, ByVal allowNullLiteral As Boolean)
    Dim norm As String
    norm = NormalizeLineEndingsLf(rawVal)
    If allowNullLiteral And LCase$(Trim$(norm)) = "null" Then
        lines.Add keyWithIndent & ": null"
    ElseIf Len(norm) = 0 Then
        lines.Add keyWithIndent & ": ''"
    ElseIf InStr(1, norm, vbLf, vbBinaryCompare) > 0 Then
        lines.Add keyWithIndent & ": |-"
        AppendIndentedMultiline lines, norm, blockSpaces
    Else
        lines.Add keyWithIndent & ": " & FormatSingleQuoted(norm)
    End If
End Sub

Private Sub AppendCrossJoinArray(ByVal lines As Collection, ByVal keyWithIndent As String, _
                                 ByVal rawVal As String, ByVal itemIndent As Long)
    Dim norm As String
    norm = NormalizeLineEndingsLf(rawVal)

    If IsNullToken(norm) Then
        lines.Add keyWithIndent & ": null"
        Exit Sub
    End If

    If Trim$(norm) = "[]" Then
        lines.Add keyWithIndent & ": []"
        Exit Sub
    End If

    lines.Add keyWithIndent & ":"

    ' Разбиваем элементы массива по строке-разделителю ---
    Dim srcLines() As String
    srcLines = Split(norm, vbLf)

    Dim items As Collection
    Set items = New Collection

    Dim buf As String
    Dim hasBuf As Boolean
    Dim i As Long
    For i = LBound(srcLines) To UBound(srcLines)
        If Trim$(srcLines(i)) = ARRAY_ITEM_DELIMITER Then
            If hasBuf Then
                items.Add buf
                buf = ""
                hasBuf = False
            End If
        Else
            If hasBuf Then
                buf = buf & vbLf & srcLines(i)
            Else
                buf = srcLines(i)
                hasBuf = True
            End If
        End If
    Next i
    If hasBuf Then items.Add buf

    Dim padDash As String
    padDash = Space$(itemIndent)
    Dim padSql As Long
    padSql = itemIndent + 2

    Dim k As Long
    For k = 1 To items.Count
        Dim itemSql As String
        itemSql = items(k)
        If Len(Trim$(itemSql)) > 0 Then
            lines.Add padDash & "- |-"
            AppendIndentedMultiline lines, itemSql, padSql
        End If
    Next k
End Sub

Private Sub AppendIndentedMultiline(ByVal lines As Collection, ByVal textLf As String, ByVal indentSpaces As Long)
    Dim pad As String
    pad = Space$(indentSpaces)
    Dim parts() As String
    parts = Split(textLf, vbLf)
    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        If Len(parts(i)) = 0 Then
            lines.Add ""
        Else
            lines.Add pad & parts(i)
        End If
    Next i
End Sub

Private Function IsNullToken(ByVal s As String) As Boolean
    Dim t As String
    t = Trim$(s)
    IsNullToken = (Len(t) = 0 Or LCase$(t) = "null" Or t = "~")
End Function

Private Function FormatSingleQuoted(ByVal s As String) As String
    FormatSingleQuoted = "'" & Replace(s, "'", "''") & "'"
End Function

Private Function FormatNullableSingleQuoted(ByVal s As String) As String
    If IsNullToken(s) Then
        FormatNullableSingleQuoted = "null"
    Else
        FormatNullableSingleQuoted = FormatSingleQuoted(Trim$(s))
    End If
End Function

Private Function FormatNullablePlain(ByVal s As String) As String
    If IsNullToken(s) Then
        FormatNullablePlain = "null"
    Else
        FormatNullablePlain = Trim$(s)
    End If
End Function

Private Function FormatPlainOrDefault(ByVal s As String, ByVal defVal As String) As String
    Dim t As String
    t = Trim$(s)
    If Len(t) = 0 Then
        FormatPlainOrDefault = defVal
    Else
        FormatPlainOrDefault = t
    End If
End Function

Private Function FormatIdentifierOrQuoted(ByVal s As String) As String
    Dim t As String
    t = Trim$(s)
    If Len(t) = 0 Then
        FormatIdentifierOrQuoted = "''"
    ElseIf IsSimpleYamlIdentifier(t) And Not IsYamlReservedWord(t) Then
        FormatIdentifierOrQuoted = t
    Else
        FormatIdentifierOrQuoted = FormatSingleQuoted(t)
    End If
End Function

Private Function FormatIntStr(ByVal s As String, ByVal defVal As Long) As String
    Dim t As String
    t = Trim$(s)
    If Len(t) = 0 Then
        FormatIntStr = CStr(defVal)
    ElseIf IsNumeric(t) Then
        FormatIntStr = CStr(CLng(Val(t)))
    Else
        FormatIntStr = t
    End If
End Function

Public Function NormalizeBoolStr(ByVal s As String, ByVal defVal As String) As String
    Dim t As String
    t = LCase$(Trim$(s))
    Select Case t
        Case "true", "истина", "1", "yes", "да"
            NormalizeBoolStr = "true"
        Case "false", "ложь", "0", "no", "нет"
            NormalizeBoolStr = "false"
        Case ""
            NormalizeBoolStr = defVal
        Case Else
            NormalizeBoolStr = t
    End Select
End Function

Public Function CellValueToString(ByVal v As Variant) As String
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then
        CellValueToString = ""
    ElseIf VarType(v) = vbBoolean Then
        If CBool(v) Then
            CellValueToString = "true"
        Else
            CellValueToString = "false"
        End If
    Else
        CellValueToString = NormalizeLineEndingsLf(CStr(v))
    End If
End Function

' Преобразование скалярного значения YAML в текст ячейки Excel
Public Function ParseScalarToCell(ByVal rawVal As String) As String
    Dim s As String
    s = Trim$(StripInlineComment(rawVal))

    If Len(s) = 0 Then
        ParseScalarToCell = ""
        Exit Function
    End If

    If LCase$(s) = "null" Or s = "~" Then
        ParseScalarToCell = ""
        Exit Function
    End If

    ' Одинарные кавычки '...'
    If Len(s) >= 2 And Left$(s, 1) = "'" And Right$(s, 1) = "'" Then
        Dim innerS As String
        innerS = Mid$(s, 2, Len(s) - 2)
        ParseScalarToCell = Replace(innerS, "''", "'")
        Exit Function
    End If

    ' Двойные кавычки "..."
    If Len(s) >= 2 And Left$(s, 1) = """" And Right$(s, 1) = """" Then
        Dim innerD As String
        innerD = Mid$(s, 2, Len(s) - 2)
        innerD = Replace(innerD, "\n", vbLf)
        innerD = Replace(innerD, "\t", vbTab)
        innerD = Replace(innerD, "\""", """")
        innerD = Replace(innerD, "\\", "\")
        ParseScalarToCell = innerD
        Exit Function
    End If

    ParseScalarToCell = s
End Function

' Удаление inline-комментария (# ...) вне кавычек
Public Function StripInlineComment(ByVal s As String) As String
    Dim inSingle As Boolean, inDouble As Boolean
    Dim i As Long, ch As String, prevCh As String

    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If i > 1 Then prevCh = Mid$(s, i - 1, 1) Else prevCh = " "

        If ch = "'" And Not inDouble Then
            inSingle = Not inSingle
        ElseIf ch = """" And Not inSingle And prevCh <> "\" Then
            inDouble = Not inDouble
        ElseIf ch = "#" And Not inSingle And Not inDouble Then
            If i = 1 Or prevCh = " " Or prevCh = vbTab Then
                StripInlineComment = RTrim$(Left$(s, i - 1))
                Exit Function
            End If
        End If
    Next i

    StripInlineComment = s
End Function

Public Function IsBlankOrComment(ByVal line As String) As Boolean
    Dim t As String
    t = Trim$(line)
    If Len(t) = 0 Then
        IsBlankOrComment = True
    ElseIf Left$(t, 1) = "#" Then
        IsBlankOrComment = True
    Else
        IsBlankOrComment = False
    End If
End Function

Public Function GetLineIndent(ByVal line As String) As Long
    Dim cnt As Long
    Dim i As Long, ch As String
    For i = 1 To Len(line)
        ch = Mid$(line, i, 1)
        If ch = " " Then
            cnt = cnt + 1
        ElseIf ch = vbTab Then
            cnt = cnt + 2
        Else
            Exit For
        End If
    Next i
    GetLineIndent = cnt
End Function

Public Function StripIndentSpaces(ByVal line As String, ByVal spacesToStrip As Long) As String
    Dim i As Long, removed As Long
    i = 1
    Do While i <= Len(line) And removed < spacesToStrip
        Dim ch As String
        ch = Mid$(line, i, 1)
        If ch = " " Then
            removed = removed + 1
            i = i + 1
        ElseIf ch = vbTab Then
            removed = removed + 2
            i = i + 1
        Else
            Exit Do
        End If
    Loop
    StripIndentSpaces = Mid$(line, i)
End Function

Public Function IsSequenceItemLine(ByVal trimmedLine As String) As Boolean
    If trimmedLine = "-" Then
        IsSequenceItemLine = True
    ElseIf Left$(trimmedLine, 2) = "- " Or Left$(trimmedLine, 2) = "-" & vbTab Then
        IsSequenceItemLine = True
    Else
        IsSequenceItemLine = False
    End If
End Function

Public Function IsYamlKeyOrSequenceLine(ByVal line As String) As Boolean
    Dim t As String
    t = Trim$(line)
    If IsSequenceItemLine(t) Then
        IsYamlKeyOrSequenceLine = True
        Exit Function
    End If
    Dim k As String, v As String
    IsYamlKeyOrSequenceLine = TrySplitKeyValue(t, k, v)
End Function

Public Function TrySplitKeyValue(ByVal trimmedLine As String, ByRef outKey As String, ByRef outVal As String) As Boolean
    outKey = ""
    outVal = ""

    Dim colonPos As Long
    colonPos = InStr(1, trimmedLine, ":", vbBinaryCompare)
    If colonPos <= 1 Then
        TrySplitKeyValue = False
        Exit Function
    End If

    ' После двоеточия должен быть конец строки или пробел/таб
    If colonPos < Len(trimmedLine) Then
        Dim nextCh As String
        nextCh = Mid$(trimmedLine, colonPos + 1, 1)
        If nextCh <> " " And nextCh <> vbTab Then
            TrySplitKeyValue = False
            Exit Function
        End If
    End If

    Dim candidateKey As String
    candidateKey = Trim$(Left$(trimmedLine, colonPos - 1))
    If Not IsSimpleYamlIdentifier(candidateKey) Then
        TrySplitKeyValue = False
        Exit Function
    End If

    outKey = candidateKey
    If colonPos < Len(trimmedLine) Then
        outVal = Trim$(Mid$(trimmedLine, colonPos + 1))
    Else
        outVal = ""
    End If
    TrySplitKeyValue = True
End Function

Public Function IsBlockScalarIndicator(ByVal cleanVal As String) As Boolean
    Select Case cleanVal
        Case "|", "|-", "|+", ">", ">-", ">+"
            IsBlockScalarIndicator = True
        Case Else
            IsBlockScalarIndicator = False
    End Select
End Function

Public Function IsSimpleYamlIdentifier(ByVal s As String) As Boolean
    If Len(s) = 0 Then
        IsSimpleYamlIdentifier = False
        Exit Function
    End If

    Dim i As Long, code As Long
    For i = 1 To Len(s)
        code = AscW(Mid$(s, i, 1))
        Dim isAlpha As Boolean, isDigit As Boolean, isUnder As Boolean
        isAlpha = (code >= 65 And code <= 90) Or (code >= 97 And code <= 122)
        isDigit = (code >= 48 And code <= 57)
        isUnder = (code = 95)

        If i = 1 Then
            If Not (isAlpha Or isUnder) Then
                IsSimpleYamlIdentifier = False
                Exit Function
            End If
        Else
            If Not (isAlpha Or isDigit Or isUnder) Then
                IsSimpleYamlIdentifier = False
                Exit Function
            End If
        End If
    Next i

    IsSimpleYamlIdentifier = True
End Function

Private Function IsYamlReservedWord(ByVal s As String) As Boolean
    Select Case LCase$(s)
        Case "true", "false", "null", "yes", "no", "on", "off", "~"
            IsYamlReservedWord = True
        Case Else
            IsYamlReservedWord = False
    End Select
End Function
