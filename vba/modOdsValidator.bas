Attribute VB_Name = "modOdsValidator"
Option Explicit

' =============================================================================
' Модуль: modOdsValidator
' Назначение: Валидация таблиц Excel по правилам schema_ods_table.json
' =============================================================================

' Возвращает коллекцию строк с описанием ошибок схемы и целостности.
' outFatalErrors = True, если есть критические ошибки файловых имён / дубликатов таблиц,
' при которых запись файлов невозможна.
Public Function ValidateOdsData(ByVal tablesColl As Collection, _
                                ByVal allColsColl As Collection, _
                                ByRef outHasFatal As Boolean) As Collection
    Dim errs As Collection
    Set errs = New Collection
    outHasFatal = False

    Dim seenTables As Object
    Set seenTables = CreateObject("Scripting.Dictionary")
    seenTables.CompareMode = 1 ' vbTextCompare

    Dim decRegex As Object
    Set decRegex = CreateObject("VBScript.RegExp")
    decRegex.Pattern = "^decimal\((?:[1-9]|[12][0-9]|3[0-8]),\s*[0-9]+\)$"
    decRegex.IgnoreCase = True
    decRegex.Global = False

    Dim wmFields As Variant
    wmFields = Array( _
        "watermark.select_wm_max_timestamp_sql", _
        "watermark.select_wm_min_timestamp_sql", _
        "watermark.select_wm_max_id_sql", _
        "watermark.select_wm_min_id_sql", _
        "watermark.select_wm_target_timestamp_sql", _
        "watermark.select_wm_target_id_sql" _
    )

    Dim tIdx As Long
    For tIdx = 1 To tablesColl.Count
        Dim t As Object
        Set t = tablesColl(tIdx)
        Dim tName As String
        tName = Trim$(CStr(t("table_name")))

        If Len(tName) = 0 Then
            errs.Add "[КРИТИЧНО] Таблица #" & tIdx & ": не заполнено поле table_name (имя YAML-файла)."
            outHasFatal = True
        ElseIf HasInvalidFileNameChars(tName) Then
            errs.Add "[КРИТИЧНО] Таблица '" & tName & "': недопустимые символы в имени файла."
            outHasFatal = True
        ElseIf seenTables.Exists(tName) Then
            errs.Add "[КРИТИЧНО] Таблица '" & tName & "': дублирующееся имя таблицы в списке таблиц."
            outHasFatal = True
        Else
            seenTables.Add tName, True
        End If

        Dim label As String
        If Len(tName) > 0 Then label = "[" & tName & "]" Else label = "[Строка #" & tIdx & "]"

        If Not IsValidBoolValue(CStr(t("enabled_flag"))) Then
            errs.Add label & " enabled_flag должен быть true или false."
        End If

        ' Секция source
        If Len(Trim$(CStr(t("source.connector_name")))) = 0 Then
            errs.Add label & " source.connector_name обязателен."
        End If
        If Len(Trim$(CStr(t("source.connector_table_schema")))) = 0 Then
            errs.Add label & " source.connector_table_schema обязателен."
        End If
        If Len(Trim$(CStr(t("source.connector_table_name")))) = 0 Then
            errs.Add label & " source.connector_table_name обязателен."
        End If
        If Not IsValidIntegerValue(CStr(t("source.connector_read_parallel_number"))) Then
            errs.Add label & " source.connector_read_parallel_number должен быть целым числом."
        End If

        ' Секция capture
        Dim capType As String
        capType = LCase$(Trim$(CStr(t("capture.data_capture_type"))))
        If capType <> "full" And capType <> "incremental" And capType <> "snapshot" Then
            errs.Add label & " capture.data_capture_type должен быть одним из: full, incremental, snapshot."
        End If

        If Not IsValidBoolValue(CStr(t("capture.hard_delete_source_flag"))) Then
            errs.Add label & " capture.hard_delete_source_flag должен быть true или false."
        End If

        ' Секция watermark: правило if/then/else из schema_ods_table.json
        Dim wIdx As Long
        If capType = "full" Then
            For wIdx = LBound(wmFields) To UBound(wmFields)
                Dim wmValFull As String
                wmValFull = Trim$(CStr(t(wmFields(wIdx))))
                If Len(wmValFull) > 0 And LCase$(wmValFull) <> "null" Then
                    errs.Add label & " " & wmFields(wIdx) & " должен быть пустым (null) при data_capture_type = full."
                End If
            Next wIdx
        ElseIf capType = "incremental" Or capType = "snapshot" Then
            For wIdx = LBound(wmFields) To UBound(wmFields)
                Dim wmValInc As String
                wmValInc = Trim$(CStr(t(wmFields(wIdx))))
                If Len(wmValInc) = 0 Or LCase$(wmValInc) = "null" Then
                    errs.Add label & " " & wmFields(wIdx) & " обязателен (строка, например <default>) при data_capture_type = " & capType & "."
                End If
            Next wIdx
        End If

        ' Секция table
        If Not IsValidIntegerValue(CStr(t("table.partitions_count"))) Then
            errs.Add label & " table.partitions_count должен быть целым числом."
        End If

        Dim expVal As String
        expVal = LCase$(Trim$(CStr(t("table.export_ahd_gp"))))
        If Len(expVal) > 0 And expVal <> "null" And expVal <> "latest" Then
            errs.Add label & " table.export_ahd_gp допускает только 'latest' или пустое значение (null)."
        End If
    Next tIdx

    ' Проверка таблицы колонок
    Dim seenColPerTable As Object
    Set seenColPerTable = CreateObject("Scripting.Dictionary")
    seenColPerTable.CompareMode = 1 ' vbTextCompare

    Dim cIdx As Long
    For cIdx = 1 To allColsColl.Count
        Dim c As Object
        Set c = allColsColl(cIdx)

        Dim colTbl As String
        colTbl = Trim$(CStr(c("table_name")))
        Dim colName As String
        colName = Trim$(CStr(c("ods_column_name")))
        Dim cLabel As String
        cLabel = "[Колонка #" & cIdx & " " & colTbl & "." & colName & "]"

        If Len(colTbl) = 0 Then
            errs.Add "[КРИТИЧНО] " & cLabel & " не указано имя таблицы (table_name)."
            outHasFatal = True
        ElseIf Not seenTables.Exists(colTbl) Then
            errs.Add cLabel & " таблица '" & colTbl & "' отсутствует в первой таблице (Tables)."
        End If

        If Len(colName) = 0 Then
            errs.Add cLabel & " не заполнено имя колонки ods_column_name."
        ElseIf Len(colTbl) > 0 Then
            Dim compositeKey As String
            compositeKey = LCase$(colTbl) & "|" & LCase$(colName)
            If seenColPerTable.Exists(compositeKey) Then
                errs.Add cLabel & " дублирующееся имя колонки '" & colName & "' в таблице '" & colTbl & "'."
            Else
                seenColPerTable.Add compositeKey, True
            End If
        End If

        Dim colType As String
        colType = Trim$(CStr(c("ods_column_type")))
        If Len(colType) > 0 Then
            If Not IsValidOdsColumnType(colType, decRegex) Then
                errs.Add cLabel & " недопустимый тип ods_column_type = '" & colType & "'."
            End If
        End If

        If Len(Trim$(CStr(c("connector_mapping_sql")))) = 0 Then
            errs.Add cLabel & " не заполнено выражение connector_mapping_sql."
        End If

        If Not IsValidBoolValue(CStr(c("rkh_flag"))) Then
            errs.Add cLabel & " rkh_flag должен быть true или false."
        End If

        If Not IsValidBoolValue(CStr(c("enabled_flag"))) Then
            errs.Add cLabel & " enabled_flag должен быть true или false."
        End If
    Next cIdx

    Set ValidateOdsData = errs
End Function

Public Function HasInvalidFileNameChars(ByVal s As String) As Boolean
    Dim badChars As String
    badChars = "\/:*?""<>|"
    Dim i As Long
    For i = 1 To Len(badChars)
        If InStr(1, s, Mid$(badChars, i, 1), vbBinaryCompare) > 0 Then
            HasInvalidFileNameChars = True
            Exit Function
        End If
    Next i
    HasInvalidFileNameChars = False
End Function

Private Function IsValidBoolValue(ByVal s As String) As Boolean
    Select Case LCase$(Trim$(s))
        Case "true", "false", "истина", "ложь", "1", "0", "yes", "no", "да", "нет"
            IsValidBoolValue = True
        Case Else
            IsValidBoolValue = False
    End Select
End Function

Private Function IsValidIntegerValue(ByVal s As String) As Boolean
    Dim t As String
    t = Trim$(s)
    If Len(t) = 0 Then
        IsValidIntegerValue = False
        Exit Function
    End If
    Dim i As Long, ch As String
    For i = 1 To Len(t)
        ch = Mid$(t, i, 1)
        If i = 1 And ch = "-" Then
            If Len(t) = 1 Then
                IsValidIntegerValue = False
                Exit Function
            End If
        ElseIf ch < "0" Or ch > "9" Then
            IsValidIntegerValue = False
            Exit Function
        End If
    Next i
    IsValidIntegerValue = True
End Function

Private Function IsValidOdsColumnType(ByVal colType As String, ByVal decRegex As Object) As Boolean
    Select Case LCase$(Trim$(colType))
        Case "boolean", "integer", "bigint", "varchar", "timestamp", _
             "timestamp(3)", "timestamp(6)", "uuid", "date", "varbinary", "time"
            IsValidOdsColumnType = True
        Case Else
            IsValidOdsColumnType = decRegex.Test(Trim$(colType))
    End Select
End Function
