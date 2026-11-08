Attribute VB_Name = "modOdsUI"
Option Explicit

' =============================================================================
' Модуль: modOdsUI
' Назначение: Отрисовка обеих таблиц на листе ODS_Config, чтение данных с листа
'             и оформление панели управления.
' =============================================================================

' =============================================================================
' ПОСТРОЕНИЕ И ОТРИСОВКА ОБЕИХ ТАБЛИЦ НА ОДНОМ ЛИСТЕ (ДРУГ ПОД ДРУГОМ)
' =============================================================================

Public Sub RenderBothTables(ByVal ws As Worksheet, ByVal tablesColl As Collection, ByVal allColsColl As Collection)
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Dim prevCalc As XlCalculation
    prevCalc = Application.Calculation
    Application.Calculation = xlCalculationManual

    On Error GoTo CleanFail

    ' Удаляем старые таблицы и полностью очищаем рабочую область ниже панели управления (с 5-й строки)
    If HasListObject(ws, TBL_COLUMNS_NAME) Then ws.ListObjects(TBL_COLUMNS_NAME).Delete
    If HasListObject(ws, TBL_TABLES_NAME) Then ws.ListObjects(TBL_TABLES_NAME).Delete

    Dim maxUsedRow As Long
    maxUsedRow = ws.UsedRange.Row + ws.UsedRange.Rows.Count + 5
    If maxUsedRow < 50 Then maxUsedRow = 50
    ws.Rows(TABLES_BANNER_ROW & ":" & maxUsedRow).Delete Shift:=xlUp

    ' --- 1. ВЕРХНЯЯ ТАБЛИЦА: АТРИБУТЫ ТАБЛИЦ ---
    DrawSectionBanner ws, TABLES_BANNER_ROW, TABLES_COL_COUNT, _
        "1. ТАБЛИЦЫ СХЕМЫ (все атрибуты кроме раздела columns)", RGB(31, 78, 120)

    DrawTableSectionGroups ws, TABLES_GROUP_ROW

    Dim tHeaders As Variant
    tHeaders = GetTablesHeaderNames()
    Dim colIdx As Long
    For colIdx = 1 To TABLES_COL_COUNT
        ws.Cells(TABLES_HEADER_ROW, colIdx).Value = tHeaders(colIdx - 1)
    Next colIdx

    Dim tRowCount As Long
    If tablesColl.Count > 0 Then tRowCount = tablesColl.Count Else tRowCount = 1

    Dim arrT() As Variant
    ReDim arrT(1 To tRowCount, 1 To TABLES_COL_COUNT)

    Dim tKeys As Variant
    tKeys = GetTablesDictKeys()

    If tablesColl.Count > 0 Then
        Dim r As Long
        For r = 1 To tablesColl.Count
            Dim td As Object
            Set td = tablesColl(r)
            For colIdx = 1 To TABLES_COL_COUNT
                arrT(r, colIdx) = CStr(td(tKeys(colIdx - 1)))
            Next colIdx
        Next r
    End If

    Dim tDataRange As Range
    Set tDataRange = ws.Range(ws.Cells(TABLES_HEADER_ROW + 1, 1), _
                              ws.Cells(TABLES_HEADER_ROW + tRowCount, TABLES_COL_COUNT))
    WriteTextArraySafe tDataRange, arrT

    Dim loTablesRange As Range
    Set loTablesRange = ws.Range(ws.Cells(TABLES_HEADER_ROW, 1), _
                                 ws.Cells(TABLES_HEADER_ROW + tRowCount, TABLES_COL_COUNT))
    Dim loTables As ListObject
    Set loTables = ws.ListObjects.Add(xlSrcRange, loTablesRange, , xlYes)
    loTables.Name = TBL_TABLES_NAME
    loTables.TableStyle = "TableStyleMedium2"

    FormatTablesListObject ws, loTables
    AddTableHeaderTooltips ws, TABLES_HEADER_ROW

    ' --- 2. НИЖНЯЯ ТАБЛИЦА: АТРИБУТЫ КОЛОНОК (РАЗДЕЛ columns) ---
    Dim colsBannerRow As Long
    colsBannerRow = TABLES_HEADER_ROW + tRowCount + 3
    Dim colsHeaderRow As Long
    colsHeaderRow = colsBannerRow + 1

    DrawSectionBanner ws, colsBannerRow, COLUMNS_COL_COUNT, _
        "2. КОЛОНКИ ТАБЛИЦ (раздел columns)", RGB(33, 89, 103)

    Dim cHeaders As Variant
    cHeaders = GetColumnsHeaderNames()
    For colIdx = 1 To COLUMNS_COL_COUNT
        ws.Cells(colsHeaderRow, colIdx).Value = cHeaders(colIdx - 1)
    Next colIdx

    Dim cRowCount As Long
    If allColsColl.Count > 0 Then cRowCount = allColsColl.Count Else cRowCount = 1

    Dim arrC() As Variant
    ReDim arrC(1 To cRowCount, 1 To COLUMNS_COL_COUNT)

    Dim cKeys As Variant
    cKeys = GetColumnsDictKeys()

    If allColsColl.Count > 0 Then
        Dim cr As Long
        For cr = 1 To allColsColl.Count
            Dim cd As Object
            Set cd = allColsColl(cr)
            For colIdx = 1 To COLUMNS_COL_COUNT
                arrC(cr, colIdx) = CStr(cd(cKeys(colIdx - 1)))
            Next colIdx
        Next cr
    End If

    Dim cDataRange As Range
    Set cDataRange = ws.Range(ws.Cells(colsHeaderRow + 1, 1), _
                              ws.Cells(colsHeaderRow + cRowCount, COLUMNS_COL_COUNT))
    WriteTextArraySafe cDataRange, arrC

    Dim loColsRange As Range
    Set loColsRange = ws.Range(ws.Cells(colsHeaderRow, 1), _
                               ws.Cells(colsHeaderRow + cRowCount, COLUMNS_COL_COUNT))
    Dim loCols As ListObject
    Set loCols = ws.ListObjects.Add(xlSrcRange, loColsRange, , xlYes)
    loCols.Name = TBL_COLUMNS_NAME
    loCols.TableStyle = "TableStyleMedium7"

    FormatColumnsListObject ws, loCols
    AddColumnsHeaderTooltips ws, colsHeaderRow

    ApplyColumnWidths ws
    RefreshColumnsTableNameValidation ws

    Application.Calculation = prevCalc
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

CleanFail:
    Application.Calculation = prevCalc
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

' =============================================================================
' СЧИТЫВАНИЕ ДАННЫХ ИЗ ОБЕИХ ТАБЛИЦ EXCEL В КОЛЛЕКЦИИ СЛОВАРЕЙ
' =============================================================================

Public Sub ReadSheetData(ByVal ws As Worksheet, ByRef outTables As Collection, ByRef outCols As Collection)
    Set outTables = New Collection
    Set outCols = New Collection

    Dim tKeys As Variant
    tKeys = GetTablesDictKeys()

    If HasListObject(ws, TBL_TABLES_NAME) Then
        Dim loT As ListObject
        Set loT = ws.ListObjects(TBL_TABLES_NAME)
        If Not loT.DataBodyRange Is Nothing Then
            Dim r As Long, c As Long
            Dim rCountT As Long
            rCountT = loT.DataBodyRange.Rows.Count

            For r = 1 To rCountT
                If Not IsRangeRowBlank(loT.DataBodyRange.Rows(r), TABLES_COL_COUNT) Then
                    Dim td As Object
                    Set td = CreateDefaultTableDict("")
                    For c = 1 To TABLES_COL_COUNT
                        td(tKeys(c - 1)) = CellValueToString(loT.DataBodyRange.Cells(r, c).Value2)
                    Next c
                    td("table_name") = Trim$(CStr(td("table_name")))
                    outTables.Add td
                End If
            Next r
        End If
    End If

    Dim cKeys As Variant
    cKeys = GetColumnsDictKeys()

    If HasListObject(ws, TBL_COLUMNS_NAME) Then
        Dim loC As ListObject
        Set loC = ws.ListObjects(TBL_COLUMNS_NAME)
        If Not loC.DataBodyRange Is Nothing Then
            Dim cr As Long, cc As Long
            Dim rCountC As Long
            rCountC = loC.DataBodyRange.Rows.Count

            For cr = 1 To rCountC
                If Not IsRangeRowBlank(loC.DataBodyRange.Rows(cr), COLUMNS_COL_COUNT) Then
                    Dim cd As Object
                    Set cd = CreateDefaultColumnDict("")
                    For cc = 1 To COLUMNS_COL_COUNT
                        cd(cKeys(cc - 1)) = CellValueToString(loC.DataBodyRange.Cells(cr, cc).Value2)
                    Next cc
                    cd("table_name") = Trim$(CStr(cd("table_name")))
                    cd("ods_column_name") = Trim$(CStr(cd("ods_column_name")))
                    outCols.Add cd
                End If
            Next cr
        End If
    End If
End Sub

' =============================================================================
' ОФОРМЛЕНИЕ ПАНЕЛИ УПРАВЛЕНИЯ, СПИСКОВ ЗНАЧЕНИЙ И ПОДСКАЗОК
' =============================================================================

Public Sub SetupTopControlPanel(ByVal ws As Worksheet)
    Application.ScreenUpdating = False

    ApplyColumnWidths ws

    ws.Rows("1:3").RowHeight = 22
    ws.Rows("4:4").RowHeight = 8

    ws.Range("B1:E3").UnMerge

    With ws.Range("A1:X3")
        .Interior.Color = RGB(242, 245, 249)
        .Font.Name = "Segoe UI"
        .Font.Size = 10
    End With
    ws.Range("A3:X3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    ws.Range("A3:X3").Borders(xlEdgeBottom).Color = RGB(180, 198, 231)
    ws.Range("A3:X3").Borders(xlEdgeBottom).Weight = xlThin

    ws.Range("A1").Value = "Папка config:"
    ws.Range("A2").Value = "Схема (папка):"
    ws.Range("A3").Value = "Статус:"
    ws.Range("A1:A3").Font.Bold = True
    ws.Range("A1:A3").HorizontalAlignment = xlRight

    ' Поля ввода пути и выбора схемы
    ws.Range("B1:D1").Merge
    ws.Range("B1:D1").Interior.Color = RGB(255, 255, 255)
    ws.Range("B1:D1").Borders.LineStyle = xlContinuous
    ws.Range("B1:D1").Borders.Color = RGB(191, 191, 191)
    ws.Range("B1").HorizontalAlignment = xlLeft

    ws.Range("B2:C2").Merge
    ws.Range("B2:C2").Interior.Color = RGB(255, 249, 230)
    ws.Range("B2:C2").Borders.LineStyle = xlContinuous
    ws.Range("B2:C2").Borders.Color = RGB(217, 150, 70)
    ws.Range("B2").Font.Bold = True
    ws.Range("B2").HorizontalAlignment = xlLeft

    ws.Range("B3:D3").Merge
    ws.Range("B3").Font.Color = RGB(64, 64, 64)
    ws.Range("B3").Font.Size = 9
    ws.Range("B3").HorizontalAlignment = xlLeft

    ' Пересоздаём кнопки панели управления (фиксируем с Placement = xlFreeFloating)
    DeleteAllButtonsOnSheet ws

    Dim baseLeft As Double
    baseLeft = ws.Cells(1, 5).Left + 8
    Dim row1Top As Double, row2Top As Double
    row1Top = 5
    row2Top = 35
    Dim btnH As Double
    btnH = 24

    CreateStyledButton ws, "btnSelectConfig", "Папка config...", "SelectConfigFolder", _
                       baseLeft, row1Top, 115, btnH, RGB(68, 114, 196)
    CreateStyledButton ws, "btnRefreshSchemas", "Обновить схемы", "RefreshSchemasButton", _
                       baseLeft + 121, row1Top, 120, btnH, RGB(68, 114, 196)
    CreateStyledButton ws, "btnLoadYaml", "Загрузить из YAML", "LoadSchemaFromYaml", _
                       baseLeft + 247, row1Top, 145, btnH, RGB(47, 117, 181)
    CreateStyledButton ws, "btnSaveYaml", "Сохранить в YAML", "SaveSchemaToYaml", _
                       baseLeft + 398, row1Top, 145, btnH, RGB(84, 130, 53)

    CreateStyledButton ws, "btnAddTable", "+ Таблица", "AddTableRow", _
                       baseLeft, row2Top, 95, btnH, RGB(58, 120, 145)
    CreateStyledButton ws, "btnDelTable", "- Таблица", "DeleteSelectedTableRow", _
                       baseLeft + 101, row2Top, 95, btnH, RGB(176, 70, 50)
    CreateStyledButton ws, "btnAddCol", "+ Колонка", "AddColumnRow", _
                       baseLeft + 202, row2Top, 95, btnH, RGB(58, 120, 145)
    CreateStyledButton ws, "btnDelCol", "- Колонка", "DeleteSelectedColumnRow", _
                       baseLeft + 303, row2Top, 95, btnH, RGB(176, 70, 50)
    CreateStyledButton ws, "btnValidate", "Проверить схему", "ValidateSheetButton", _
                       baseLeft + 404, row2Top, 139, btnH, RGB(127, 96, 0)

    ' Закрепляем верхние 4 строки (панель управления и кнопки всегда перед глазами при прокрутке)
    If Not ActiveWindow Is Nothing Then
        ws.Activate
        ActiveWindow.FreezePanes = False
        ws.Range("A5").Select
        ActiveWindow.FreezePanes = True
        ws.Range("B2").Select
    End If

    Application.ScreenUpdating = True
End Sub

Private Sub CreateStyledButton(ByVal ws As Worksheet, ByVal shpName As String, ByVal caption As String, _
                               ByVal macroName As String, ByVal leftPt As Double, ByVal topPt As Double, _
                               ByVal widthPt As Double, ByVal heightPt As Double, ByVal fillColor As Long)
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(5, leftPt, topPt, widthPt, heightPt) ' msoShapeRoundedRectangle = 5
    shp.Name = shpName
    shp.Placement = xlFreeFloating
    shp.Fill.ForeColor.RGB = fillColor
    shp.Line.Visible = 0 ' msoFalse

    With shp.TextFrame2
        .VerticalAnchor = 3 ' msoAnchorMiddle
        .MarginLeft = 4
        .MarginRight = 4
        .MarginTop = 2
        .MarginBottom = 2
        .WordWrap = 0 ' msoFalse
        With .TextRange
            .Text = caption
            .ParagraphFormat.Alignment = 2 ' msoAlignCenter
            With .Font
                .Name = "Segoe UI"
                .Size = 9.5
                .Bold = -1 ' msoTrue
                .Fill.ForeColor.RGB = RGB(255, 255, 255)
            End With
        End With
    End With

    shp.OnAction = macroName
End Sub

Private Sub DeleteAllButtonsOnSheet(ByVal ws As Worksheet)
    Dim i As Long
    For i = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(i).Name, 3) = "btn" Then
            ws.Shapes(i).Delete
        End If
    Next i
End Sub

Private Sub DrawSectionBanner(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal colSpan As Long, _
                              ByVal titleText As String, ByVal bgColor As Long)
    Dim rng As Range
    Set rng = ws.Range(ws.Cells(rowNum, 1), ws.Cells(rowNum, colSpan))
    ws.Rows(rowNum).RowHeight = 22
    With rng
        .Interior.Color = bgColor
        .Font.Name = "Segoe UI"
        .Font.Size = 10.5
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .VerticalAlignment = xlCenter
    End With
    ws.Cells(rowNum, 1).Value = "  " & titleText
End Sub

Private Sub DrawTableSectionGroups(ByVal ws As Worksheet, ByVal rowNum As Long)
    ws.Rows(rowNum).RowHeight = 18
    StyleGroupRange ws.Range(ws.Cells(rowNum, 1), ws.Cells(rowNum, 1)), "Файл", RGB(217, 225, 242)
    StyleGroupRange ws.Range(ws.Cells(rowNum, 2), ws.Cells(rowNum, 3)), "Корень YAML", RGB(226, 239, 218)
    StyleGroupRange ws.Range(ws.Cells(rowNum, 4), ws.Cells(rowNum, 7)), "source", RGB(252, 228, 214)
    StyleGroupRange ws.Range(ws.Cells(rowNum, 8), ws.Cells(rowNum, 15)), "capture", RGB(255, 242, 204)
    StyleGroupRange ws.Range(ws.Cells(rowNum, 16), ws.Cells(rowNum, 21)), "watermark", RGB(221, 235, 247)
    StyleGroupRange ws.Range(ws.Cells(rowNum, 22), ws.Cells(rowNum, 24)), "table", RGB(237, 230, 245)
End Sub

Private Sub StyleGroupRange(ByVal rng As Range, ByVal title As String, ByVal bgColor As Long)
    rng.Cells(1, 1).Value = title
    With rng
        .Interior.Color = bgColor
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Font.Bold = True
        .Font.Color = RGB(50, 50, 50)
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).Color = RGB(166, 166, 166)
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeRight).Color = RGB(166, 166, 166)
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).Color = RGB(166, 166, 166)
    End With
End Sub

Private Sub FormatTablesListObject(ByVal ws As Worksheet, ByVal lo As ListObject)
    ws.Rows(lo.HeaderRowRange.Row).RowHeight = 24
    lo.HeaderRowRange.Font.Name = "Segoe UI"
    lo.HeaderRowRange.Font.Size = 9.5

    If Not lo.DataBodyRange Is Nothing Then
        With lo.DataBodyRange
            .Font.Name = "Consolas"
            .Font.Size = 9.5
            .VerticalAlignment = xlTop
            .WrapText = False
        End With
        ' Включаем перенос строк для многострочных SQL-полей
        Dim sqlCols As Variant, sc As Variant
        sqlCols = Array(11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21)
        For Each sc In sqlCols
            lo.ListColumns(CLng(sc)).DataBodyRange.WrapText = True
        Next sc

        ' Выпадающие списки из скрытого листа _Lists
        Dim boolRef As String, capRef As String, expRef As String
        boolRef = "=ListBool"
        capRef = "=ListCapture"
        expRef = "=ListExport"

        ApplyRangeValidation lo.ListColumns(3).DataBodyRange, boolRef, False
        ApplyRangeValidation lo.ListColumns(8).DataBodyRange, capRef, False
        ApplyRangeValidation lo.ListColumns(10).DataBodyRange, boolRef, False
        ApplyRangeValidation lo.ListColumns(24).DataBodyRange, expRef, False
    End If
End Sub

Private Sub FormatColumnsListObject(ByVal ws As Worksheet, ByVal lo As ListObject)
    ws.Rows(lo.HeaderRowRange.Row).RowHeight = 24
    lo.HeaderRowRange.Font.Name = "Segoe UI"
    lo.HeaderRowRange.Font.Size = 9.5

    If Not lo.DataBodyRange Is Nothing Then
        With lo.DataBodyRange
            .Font.Name = "Consolas"
            .Font.Size = 9.5
            .VerticalAlignment = xlTop
            .WrapText = False
        End With
        lo.ListColumns(4).DataBodyRange.WrapText = True
        lo.ListColumns(7).DataBodyRange.WrapText = True

        Dim boolRef As String, typeRef As String
        boolRef = "=ListBool"
        typeRef = "=ListColTypes"

        ' ShowError = False позволяет также вводить decimal(38,4) и любые кастомные decimal(p,s)
        ApplyRangeValidation lo.ListColumns(3).DataBodyRange, typeRef, False
        ApplyRangeValidation lo.ListColumns(5).DataBodyRange, boolRef, False
        ApplyRangeValidation lo.ListColumns(6).DataBodyRange, boolRef, False
    End If
End Sub

Public Sub RefreshColumnsTableNameValidation(ByVal ws As Worksheet)
    If Not HasListObject(ws, TBL_TABLES_NAME) Or Not HasListObject(ws, TBL_COLUMNS_NAME) Then Exit Sub
    Dim loT As ListObject, loC As ListObject
    Set loT = ws.ListObjects(TBL_TABLES_NAME)
    Set loC = ws.ListObjects(TBL_COLUMNS_NAME)
    If loT.DataBodyRange Is Nothing Or loC.DataBodyRange Is Nothing Then Exit Sub

    Dim tblNamesAddr As String
    tblNamesAddr = "=" & loT.ListColumns(1).DataBodyRange.Address(True, True)
    ApplyRangeValidation loC.ListColumns(1).DataBodyRange, tblNamesAddr, False
End Sub

Public Sub ApplyColumnWidths(ByVal ws As Worksheet)
    Dim widths As Variant
    widths = Array( _
        24, 18, 15, _
        20, 22, 24, 14, _
        18, 18, 22, 30, 26, 36, 26, 30, _
        26, 26, 24, 24, 42, 24, _
        22, 16, 16 _
    )
    Dim i As Long
    For i = 0 To UBound(widths)
        ws.Columns(i + 1).ColumnWidth = widths(i)
    Next i
    ' Учитываем комфортную ширину первых 7 колонок для нижней таблицы (Columns)
    ws.Columns(2).ColumnWidth = 26  ' ods_column_name / schedule_cron
    ws.Columns(3).ColumnWidth = 18  ' ods_column_type / enabled_flag
    ws.Columns(4).ColumnWidth = 38  ' connector_mapping_sql / connector_name
    ws.Columns(5).ColumnWidth = 22  ' rkh_flag / connector_table_schema
    ws.Columns(6).ColumnWidth = 22  ' enabled_flag / connector_table_name
    ws.Columns(7).ColumnWidth = 28  ' comment / connector_read_parallel_number
End Sub

Private Sub AddTableHeaderTooltips(ByVal ws As Worksheet, ByVal hdrRow As Long)
    SetCellCommentSafe ws.Cells(hdrRow, 1), "Имя YAML-файла без расширения (имя таблицы ODS в выбранной схеме)"
    SetCellCommentSafe ws.Cells(hdrRow, 2), "Расписание загрузки (cron) или пусто (null)"
    SetCellCommentSafe ws.Cells(hdrRow, 3), "Флаг включения загрузки (true / false)"
    SetCellCommentSafe ws.Cells(hdrRow, 4), "source.connector_name: Имя коннектора в Trino"
    SetCellCommentSafe ws.Cells(hdrRow, 5), "source.connector_table_schema: Имя схемы в коннекторе Trino"
    SetCellCommentSafe ws.Cells(hdrRow, 6), "source.connector_table_name: Имя таблицы в коннекторе Trino"
    SetCellCommentSafe ws.Cells(hdrRow, 7), "source.connector_read_parallel_number: Кол-во потоков загрузки (целое число)"
    SetCellCommentSafe ws.Cells(hdrRow, 8), "capture.data_capture_type: Тип захвата данных (full / incremental / snapshot)"
    SetCellCommentSafe ws.Cells(hdrRow, 9), "capture.pattern_capture: Паттерн захвата данных (или пусто -> null)"
    SetCellCommentSafe ws.Cells(hdrRow, 10), "capture.hard_delete_source_flag: Признак физического удаления на источнике (true / false)"
    SetCellCommentSafe ws.Cells(hdrRow, 11), "capture.soft_delete_clause_trino: SQL-условие пометки строк удалёнными"
    SetCellCommentSafe ws.Cells(hdrRow, 12), "capture.where_clause_trino: Доп. фильтр источника без WHERE (диалект Trino)"
    SetCellCommentSafe ws.Cells(hdrRow, 13), "capture.watermark_clause_trino: Условие инкремента с WM без WHERE (диалект Trino)"
    SetCellCommentSafe ws.Cells(hdrRow, 14), "capture.order_latest_clause_trino: Условие сортировки дедубликации по RKH"
    SetCellCommentSafe ws.Cells(hdrRow, 15), "capture.cross_join_lateral_selects_trino: Массив SQL-выражений (несколько элементов разделяйте строкой ---)"
    SetCellCommentSafe ws.Cells(hdrRow, 16), "watermark.select_wm_max_timestamp_sql: пусто при full; иначе строка SQL или <default>"
    SetCellCommentSafe ws.Cells(hdrRow, 17), "watermark.select_wm_min_timestamp_sql: пусто при full; иначе строка SQL или <default>"
    SetCellCommentSafe ws.Cells(hdrRow, 18), "watermark.select_wm_max_id_sql: пусто при full; иначе строка SQL или <default>"
    SetCellCommentSafe ws.Cells(hdrRow, 19), "watermark.select_wm_min_id_sql: пусто при full; иначе строка SQL или <default>"
    SetCellCommentSafe ws.Cells(hdrRow, 20), "watermark.select_wm_target_timestamp_sql: пусто при full; иначе SQL вычисления wm_target_timestamp"
    SetCellCommentSafe ws.Cells(hdrRow, 21), "watermark.select_wm_target_id_sql: пусто при full; иначе SQL вычисления wm_target_id"
    SetCellCommentSafe ws.Cells(hdrRow, 22), "table.comment: Комментарий к таблице"
    SetCellCommentSafe ws.Cells(hdrRow, 23), "table.partitions_count: Количество партиций в Iceberg (целое число)"
    SetCellCommentSafe ws.Cells(hdrRow, 24), "table.export_ahd_gp: latest или пусто (null)"
End Sub

Private Sub AddColumnsHeaderTooltips(ByVal ws As Worksheet, ByVal hdrRow As Long)
    SetCellCommentSafe ws.Cells(hdrRow, 1), "Имя таблицы ODS (должно совпадать с table_name в верхней таблице)"
    SetCellCommentSafe ws.Cells(hdrRow, 2), "Имя колонки в таблице ODS (ods_column_name)"
    SetCellCommentSafe ws.Cells(hdrRow, 3), "Тип колонки в диалекте Trino (из списка или decimal(p,s), например decimal(38,4))"
    SetCellCommentSafe ws.Cells(hdrRow, 4), "SQL-выражение в диалекте Trino для извлечения колонки (connector_mapping_sql)"
    SetCellCommentSafe ws.Cells(hdrRow, 5), "Входит в бизнес-ключ RKH (true / false)"
    SetCellCommentSafe ws.Cells(hdrRow, 6), "Признак активности колонки (true / false)"
    SetCellCommentSafe ws.Cells(hdrRow, 7), "Описание колонки в таблице ODS"
End Sub

' =============================================================================
' СЛУЖЕБНЫЕ ФУНКЦИИ И МЕТАДАННЫЕ КОЛОНОК
' =============================================================================

Public Function GetTablesHeaderNames() As Variant
    Dim a(0 To 23) As Variant
    a(0) = "table_name": a(1) = "schedule_cron": a(2) = "enabled_flag"
    a(3) = "connector_name": a(4) = "connector_table_schema": a(5) = "connector_table_name"
    a(6) = "connector_read_parallel_number": a(7) = "data_capture_type": a(8) = "pattern_capture"
    a(9) = "hard_delete_source_flag": a(10) = "soft_delete_clause_trino": a(11) = "where_clause_trino"
    a(12) = "watermark_clause_trino": a(13) = "order_latest_clause_trino": a(14) = "cross_join_lateral_selects_trino"
    a(15) = "select_wm_max_timestamp_sql": a(16) = "select_wm_min_timestamp_sql": a(17) = "select_wm_max_id_sql"
    a(18) = "select_wm_min_id_sql": a(19) = "select_wm_target_timestamp_sql": a(20) = "select_wm_target_id_sql"
    a(21) = "comment": a(22) = "partitions_count": a(23) = "export_ahd_gp"
    GetTablesHeaderNames = a
End Function

Public Function GetTablesDictKeys() As Variant
    Dim a(0 To 23) As Variant
    a(0) = "table_name": a(1) = "schedule_cron": a(2) = "enabled_flag"
    a(3) = "source.connector_name": a(4) = "source.connector_table_schema"
    a(5) = "source.connector_table_name": a(6) = "source.connector_read_parallel_number"
    a(7) = "capture.data_capture_type": a(8) = "capture.pattern_capture"
    a(9) = "capture.hard_delete_source_flag": a(10) = "capture.soft_delete_clause_trino"
    a(11) = "capture.where_clause_trino": a(12) = "capture.watermark_clause_trino"
    a(13) = "capture.order_latest_clause_trino": a(14) = "capture.cross_join_lateral_selects_trino"
    a(15) = "watermark.select_wm_max_timestamp_sql": a(16) = "watermark.select_wm_min_timestamp_sql"
    a(17) = "watermark.select_wm_max_id_sql": a(18) = "watermark.select_wm_min_id_sql"
    a(19) = "watermark.select_wm_target_timestamp_sql": a(20) = "watermark.select_wm_target_id_sql"
    a(21) = "table.comment": a(22) = "table.partitions_count": a(23) = "table.export_ahd_gp"
    GetTablesDictKeys = a
End Function

Public Function GetColumnsHeaderNames() As Variant
    GetColumnsHeaderNames = Array( _
        "table_name", "ods_column_name", "ods_column_type", _
        "connector_mapping_sql", "rkh_flag", "enabled_flag", "comment" _
    )
End Function

Public Function GetColumnsDictKeys() As Variant
    GetColumnsDictKeys = Array( _
        "table_name", "ods_column_name", "ods_column_type", _
        "connector_mapping_sql", "rkh_flag", "enabled_flag", "comment" _
    )
End Function
