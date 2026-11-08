Attribute VB_Name = "modOdsHelpers"
Option Explicit

' =============================================================================
' Модуль: modOdsHelpers
' Назначение: Операции добавления/удаления строк, валидация листа
'             и вспомогательные утилиты работы с листами и диапазонами Excel.
' =============================================================================

' Кнопка: Проверить таблицы на соответствие schema_ods_table.json
Public Sub ValidateSheetButton()
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = EnsureMainSheet()

    Dim tablesColl As Collection, allColsColl As Collection
    ReadSheetData ws, tablesColl, allColsColl

    Dim hasFatal As Boolean
    Dim errs As Collection
    Set errs = ValidateOdsData(tablesColl, allColsColl, hasFatal)

    If errs.Count = 0 Then
        SetStatus ws, "Проверка пройдена успешно: таблиц=" & tablesColl.Count & ", колонок=" & allColsColl.Count
        MsgBox "Все данные корректны и соответствуют schema_ods_table.json!" & vbCrLf & vbCrLf & _
               "Проверено таблиц: " & tablesColl.Count & vbCrLf & _
               "Проверено колонок: " & allColsColl.Count, vbInformation, "Валидация пройдена"
    Else
        SetStatus ws, "Найдено замечаний при проверке: " & errs.Count
        MsgBox "Найдено замечаний по схеме (" & errs.Count & " шт.):" & vbCrLf & vbCrLf & _
               FormatErrorList(errs, 18), vbExclamation, "Результаты проверки"
    End If
    Exit Sub

ErrHandler:
    MsgBox "Ошибка при проверке данных: " & Err.Description, vbCritical, "ODS YAML Manager"
End Sub

' Кнопка: Безопасно добавить новую строку в первую таблицу (со сдвигом второй таблицы вниз)
Public Sub AddTableRow()
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = EnsureMainSheet()

    If Not HasListObject(ws, TBL_TABLES_NAME) Then
        InitOdsWorkspace
    End If

    Dim lo As ListObject
    Set lo = ws.ListObjects(TBL_TABLES_NAME)

    Application.ScreenUpdating = False

    Dim targetRow As ListRow
    If lo.ListRows.Count = 1 And IsListRowCompletelyBlank(lo.ListRows(1)) Then
        Set targetRow = lo.ListRows(1)
    Else
        ' Вставляем полную строку листа ниже таблицы, чтобы нижняя таблица tblColumns аккуратно сдвинулась вниз
        Dim insertSheetRow As Long
        insertSheetRow = lo.Range.Row + lo.Range.Rows.Count
        ws.Rows(insertSheetRow).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        Set targetRow = lo.ListRows.Add
    End If

    With targetRow.Range
        .NumberFormat = "@"
        .Cells(1, 1).Value = "new_table"
        .Cells(1, 2).Value = "12 0 * * *"
        .Cells(1, 3).Value = "true"
        .Cells(1, 7).Value = "1"
        .Cells(1, 8).Value = "full"
        .Cells(1, 10).Value = "false"
        .Cells(1, 23).Value = "1"
        .Cells(1, 24).Value = "latest"
    End With

    RefreshColumnsTableNameValidation ws
    Application.ScreenUpdating = True
    targetRow.Range.Cells(1, 1).Select
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Ошибка добавления таблицы: " & Err.Description, vbCritical, "ODS YAML Manager"
End Sub

' Кнопка: Удалить выбранную строку таблицы из первой таблицы
Public Sub DeleteSelectedTableRow()
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = EnsureMainSheet()
    If Not HasListObject(ws, TBL_TABLES_NAME) Then Exit Sub

    Dim lo As ListObject
    Set lo = ws.ListObjects(TBL_TABLES_NAME)
    If lo.DataBodyRange Is Nothing Then Exit Sub

    Dim isect As Range
    Set isect = IntersectActiveSelection(ws, lo.DataBodyRange)
    If isect Is Nothing Then
        MsgBox "Выделите ячейку в строке первой таблицы (ТАБЛИЦЫ СХЕМЫ), которую нужно удалить.", _
               vbInformation, "Удаление таблицы"
        Exit Sub
    End If

    Dim sheetRow As Long
    sheetRow = isect.Cells(1, 1).Row
    Dim relRow As Long
    relRow = sheetRow - lo.DataBodyRange.Row + 1
    Dim tblName As String
    tblName = Trim$(CellValueToString(lo.DataBodyRange.Cells(relRow, 1).Value2))

    Dim ans As VbMsgBoxResult
    If Len(tblName) > 0 Then
        ans = MsgBox("Удалить строку таблицы '" & tblName & "' из листа Excel?" & vbCrLf & _
                     "(Также будут удалены строки её колонок из нижней таблицы; " & _
                     "при сохранении файл " & tblName & ".yml будет удалён с диска).", _
                     vbYesNo + vbQuestion, "Удаление таблицы")
        If ans = vbNo Then Exit Sub
    End If

    Application.ScreenUpdating = False

    ' Удаляем связанные колонки из нижней таблицы (снизу вверх)
    If Len(tblName) > 0 And HasListObject(ws, TBL_COLUMNS_NAME) Then
        Dim loCols As ListObject
        Set loCols = ws.ListObjects(TBL_COLUMNS_NAME)
        If Not loCols.DataBodyRange Is Nothing Then
            Dim r As Long
            For r = loCols.ListRows.Count To 1 Step -1
                If StrComp(Trim$(CellValueToString(loCols.ListRows(r).Range.Cells(1, 1).Value2)), tblName, vbTextCompare) = 0 Then
                    If loCols.ListRows.Count = 1 Then
                        loCols.ListRows(1).Range.ClearContents
                    Else
                        loCols.ListRows(r).Delete
                    End If
                End If
            Next r
        End If
    End If

    ' Удаляем строку из верхней таблицы со сдвигом листа вверх
    If lo.ListRows.Count = 1 Then
        lo.ListRows(1).Range.ClearContents
    Else
        ws.Rows(sheetRow).Delete Shift:=xlUp
    End If

    RefreshColumnsTableNameValidation ws
    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Ошибка удаления строки таблицы: " & Err.Description, vbCritical, "ODS YAML Manager"
End Sub

' Кнопка: Добавить новую строку во вторую таблицу (КОЛОНКИ ТАБЛИЦ)
Public Sub AddColumnRow()
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = EnsureMainSheet()

    If Not HasListObject(ws, TBL_COLUMNS_NAME) Then
        InitOdsWorkspace
    End If

    ' Определяем имя таблицы по текущему выделению (в верхней или нижней таблице)
    Dim defaultTblName As String
    defaultTblName = DetectActiveTableName(ws)

    Dim loCols As ListObject
    Set loCols = ws.ListObjects(TBL_COLUMNS_NAME)

    Application.ScreenUpdating = False

    Dim targetRow As ListRow
    If loCols.ListRows.Count = 1 And IsListRowCompletelyBlank(loCols.ListRows(1)) Then
        Set targetRow = loCols.ListRows(1)
    Else
        Set targetRow = loCols.ListRows.Add
    End If

    With targetRow.Range
        .NumberFormat = "@"
        .Cells(1, 1).Value = defaultTblName
        .Cells(1, 2).Value = ""
        .Cells(1, 3).Value = "varchar"
        .Cells(1, 4).Value = ""
        .Cells(1, 5).Value = "false"
        .Cells(1, 6).Value = "true"
        .Cells(1, 7).Value = ""
    End With

    Application.ScreenUpdating = True
    If Len(defaultTblName) > 0 Then
        targetRow.Range.Cells(1, 2).Select
    Else
        targetRow.Range.Cells(1, 1).Select
    End If
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Ошибка добавления колонки: " & Err.Description, vbCritical, "ODS YAML Manager"
End Sub

' Кнопка: Удалить выделенные строки из второй таблицы (КОЛОНКИ ТАБЛИЦ)
Public Sub DeleteSelectedColumnRow()
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = EnsureMainSheet()
    If Not HasListObject(ws, TBL_COLUMNS_NAME) Then Exit Sub

    Dim loCols As ListObject
    Set loCols = ws.ListObjects(TBL_COLUMNS_NAME)
    If loCols.DataBodyRange Is Nothing Then Exit Sub

    Dim isect As Range
    Set isect = IntersectActiveSelection(ws, loCols.DataBodyRange)
    If isect Is Nothing Then
        MsgBox "Выделите ячейку в строке нижней таблицы (КОЛОНКИ ТАБЛИЦ), которую нужно удалить.", _
               vbInformation, "Удаление колонки"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Dim firstRow As Long, lastRow As Long, r As Long
    firstRow = isect.Rows(1).Row - loCols.DataBodyRange.Row + 1
    lastRow = firstRow + isect.Rows.Count - 1

    For r = lastRow To firstRow Step -1
        If loCols.ListRows.Count = 1 Then
            loCols.ListRows(1).Range.ClearContents
        Else
            loCols.ListRows(r).Delete
        End If
    Next r
    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Ошибка удаления строки колонки: " & Err.Description, vbCritical, "ODS YAML Manager"
End Sub

Public Function EnsureMainSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(MAIN_SHEET_NAME)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets(1)
        ws.Name = MAIN_SHEET_NAME
    End If
    Set EnsureMainSheet = ws
End Function

Public Function EnsureListsSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(LISTS_SHEET_NAME)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = LISTS_SHEET_NAME
    End If

    ws.Visible = xlSheetVeryHidden

    ' Статические справочники для Data Validation (работают независимо от локали Excel)
    ws.Cells(1, 2).Value = "true"
    ws.Cells(2, 2).Value = "false"

    ws.Cells(1, 3).Value = "full"
    ws.Cells(2, 3).Value = "incremental"
    ws.Cells(3, 3).Value = "snapshot"

    ws.Cells(1, 4).Value = "latest"

    Dim colTypes As Variant
    colTypes = Array("varchar", "integer", "bigint", "boolean", "timestamp", _
                     "timestamp(3)", "timestamp(6)", "date", "time", "uuid", "varbinary", "decimal(38,4)")
    Dim i As Long
    For i = 0 To UBound(colTypes)
        ws.Cells(i + 1, 5).Value = colTypes(i)
    Next i

    EnsureNamedRangeSafe "ListBool", "='" & LISTS_SHEET_NAME & "'!$B$1:$B$2"
    EnsureNamedRangeSafe "ListCapture", "='" & LISTS_SHEET_NAME & "'!$C$1:$C$3"
    EnsureNamedRangeSafe "ListExport", "='" & LISTS_SHEET_NAME & "'!$D$1:$D$1"
    EnsureNamedRangeSafe "ListColTypes", "='" & LISTS_SHEET_NAME & "'!$E$1:$E$12"
    Set EnsureListsSheet = ws
End Function

Public Sub AutoDetectConfigFolder(ByVal ws As Worksheet)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim eff As String
    eff = GetEffectiveConfigPath(ws)
    If Len(eff) > 0 Then
        If fso.FolderExists(eff) Then Exit Sub
    End If

    If Len(ThisWorkbook.Path) > 0 Then
        Dim candidate As String
        candidate = ThisWorkbook.Path & "\config"
        If fso.FolderExists(candidate) Then
            ws.Range(CELL_CONFIG_PATH).Value = ".\config"
        End If
    End If
End Sub

Public Function GetEffectiveConfigPath(ByVal ws As Worksheet) As String
    Dim p As String
    p = Trim$(CellValueToString(ws.Range(CELL_CONFIG_PATH).Value2))
    If Len(p) = 0 Then
        GetEffectiveConfigPath = ""
        Exit Function
    End If
    p = Replace(p, "/", "\")
    If p = "." Then
        p = ThisWorkbook.Path
    ElseIf Left$(p, 2) = ".\" Then
        If Len(ThisWorkbook.Path) > 0 Then
            p = ThisWorkbook.Path & Mid$(p, 2)
        End If
    ElseIf InStr(1, p, ":", vbBinaryCompare) = 0 And Left$(p, 2) <> "\\" Then
        If Len(ThisWorkbook.Path) > 0 Then
            p = ThisWorkbook.Path & "\" & p
        End If
    End If
    If Len(p) > 3 And Right$(p, 1) = "\" Then p = Left$(p, Len(p) - 1)
    GetEffectiveConfigPath = p
End Function

Public Function GetSelectedSchemaFolderPath(ByVal ws As Worksheet, ByVal showAlerts As Boolean) As String
    AutoDetectConfigFolder ws
    Dim cfgPath As String
    cfgPath = GetEffectiveConfigPath(ws)
    If Len(cfgPath) = 0 Then
        If showAlerts Then
            MsgBox "Укажите путь к папке config (кнопка 'Папка config...').", vbExclamation, "ODS YAML Manager"
        End If
        GetSelectedSchemaFolderPath = ""
        Exit Function
    End If

    Dim schemaName As String
    schemaName = Trim$(CellValueToString(ws.Range(CELL_SCHEMA_NAME).Value2))
    If Len(schemaName) = 0 Then
        If showAlerts Then
            MsgBox "Выберите схему (папку) в выпадающем списке в ячейке " & CELL_SCHEMA_NAME & ".", _
                   vbExclamation, "ODS YAML Manager"
        End If
        GetSelectedSchemaFolderPath = ""
        Exit Function
    End If

    GetSelectedSchemaFolderPath = cfgPath & "\" & schemaName
End Function

' Безопасная запись 2D текстового массива в диапазон без автопреобразования чисел/дат/формул
Public Sub WriteTextArraySafe(ByVal targetRange As Range, ByRef arr() As Variant)
    targetRange.NumberFormat = "@"

    Dim rMax As Long, cMax As Long
    rMax = UBound(arr, 1)
    cMax = UBound(arr, 2)

    Dim specialCells As Collection
    Set specialCells = New Collection

    Dim r As Long, c As Long
    For r = 1 To rMax
        For c = 1 To cMax
            Dim s As String
            s = CStr(arr(r, c))
            If Len(s) > 0 Then
                Dim ch As String
                ch = Left$(s, 1)
                If ch = "=" Or ch = "+" Or ch = "-" Or ch = "@" Then
                    specialCells.Add Array(r, c, s)
                    arr(r, c) = ""
                End If
            End If
        Next c
    Next r

    targetRange.Value2 = arr

    Dim idx As Long
    For idx = 1 To specialCells.Count
        Dim info As Variant
        info = specialCells(idx)
        With targetRange.Cells(CLng(info(0)), CLng(info(1)))
            .NumberFormat = "General"
            .Value = "'" & CStr(info(2))
        End With
    Next idx
End Sub

Public Sub ApplyRangeValidation(ByVal rng As Range, ByVal formulaRef As String, ByVal strictAlert As Boolean)
    On Error Resume Next
    rng.Validation.Delete
    rng.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, _
                       Operator:=xlBetween, Formula1:=formulaRef
    rng.Validation.IgnoreBlank = True
    rng.Validation.InCellDropdown = True
    rng.Validation.ShowError = strictAlert
    On Error GoTo 0
End Sub

Public Sub RemoveValidationSafe(ByVal rng As Range)
    On Error Resume Next
    rng.Validation.Delete
    On Error GoTo 0
End Sub

Public Sub SetCellCommentSafe(ByVal cell As Range, ByVal txt As String)
    On Error Resume Next
    cell.ClearComments
    Dim cmt As Comment
    Set cmt = cell.AddComment(txt)
    If Not cmt Is Nothing Then
        cmt.Shape.TextFrame.AutoSize = False
        cmt.Shape.Width = 240
        cmt.Shape.Height = 55
    End If
    On Error GoTo 0
End Sub

Public Function HasListObject(ByVal ws As Worksheet, ByVal loName As String) As Boolean
    Dim lo As ListObject
    On Error Resume Next
    Set lo = ws.ListObjects(loName)
    On Error GoTo 0
    HasListObject = Not (lo Is Nothing)
End Function

Public Function IsListRowCompletelyBlank(ByVal lr As ListRow) As Boolean
    IsListRowCompletelyBlank = IsRangeRowBlank(lr.Range, lr.Range.Columns.Count)
End Function

Public Function IsRangeRowBlank(ByVal rowRange As Range, ByVal colCnt As Long) As Boolean
    Dim c As Long
    For c = 1 To colCnt
        If Len(Trim$(CellValueToString(rowRange.Cells(1, c).Value2))) > 0 Then
            IsRangeRowBlank = False
            Exit Function
        End If
    Next c
    IsRangeRowBlank = True
End Function

Public Function DetectActiveTableName(ByVal ws As Worksheet) As String
    DetectActiveTableName = ""
    On Error Resume Next

    If HasListObject(ws, TBL_TABLES_NAME) Then
        Dim loT As ListObject
        Set loT = ws.ListObjects(TBL_TABLES_NAME)
        If Not loT.DataBodyRange Is Nothing Then
            Dim isectT As Range
            Set isectT = IntersectActiveSelection(ws, loT.DataBodyRange)
            If Not isectT Is Nothing Then
                Dim rT As Long
                rT = isectT.Cells(1, 1).Row - loT.DataBodyRange.Row + 1
                DetectActiveTableName = Trim$(CellValueToString(loT.DataBodyRange.Cells(rT, 1).Value2))
                If Len(DetectActiveTableName) > 0 Then Exit Function
            End If
        End If
    End If

    If HasListObject(ws, TBL_COLUMNS_NAME) Then
        Dim loC As ListObject
        Set loC = ws.ListObjects(TBL_COLUMNS_NAME)
        If Not loC.DataBodyRange Is Nothing Then
            Dim isectC As Range
            Set isectC = IntersectActiveSelection(ws, loC.DataBodyRange)
            If Not isectC Is Nothing Then
                Dim rC As Long
                rC = isectC.Cells(1, 1).Row - loC.DataBodyRange.Row + 1
                DetectActiveTableName = Trim$(CellValueToString(loC.DataBodyRange.Cells(rC, 1).Value2))
                If Len(DetectActiveTableName) > 0 Then Exit Function
            End If
        End If
    End If

    ' Если курсор вне таблиц и таблица всего одна — берём её имя
    If HasListObject(ws, TBL_TABLES_NAME) Then
        Set loT = ws.ListObjects(TBL_TABLES_NAME)
        If Not loT.DataBodyRange Is Nothing Then
            If loT.ListRows.Count = 1 Then
                DetectActiveTableName = Trim$(CellValueToString(loT.DataBodyRange.Cells(1, 1).Value2))
            End If
        End If
    End If
    On Error GoTo 0
End Function

Public Function FindOrphanColumnsSummary(ByVal tablesColl As Collection, ByVal allColsColl As Collection) As String
    Dim validSet As Object
    Set validSet = CreateObject("Scripting.Dictionary")
    validSet.CompareMode = 1

    Dim i As Long
    For i = 1 To tablesColl.Count
        Dim tn As String
        tn = Trim$(CStr(tablesColl(i)("table_name")))
        If Len(tn) > 0 And Not validSet.Exists(tn) Then validSet.Add tn, True
    Next i

    Dim orphanCounts As Object
    Set orphanCounts = CreateObject("Scripting.Dictionary")
    orphanCounts.CompareMode = 1

    For i = 1 To allColsColl.Count
        Dim cn As String
        cn = Trim$(CStr(allColsColl(i)("table_name")))
        If Len(cn) > 0 And Not validSet.Exists(cn) Then
            If orphanCounts.Exists(cn) Then
                orphanCounts(cn) = CLng(orphanCounts(cn)) + 1
            Else
                orphanCounts.Add cn, 1
            End If
        End If
    Next i

    If orphanCounts.Count = 0 Then
        FindOrphanColumnsSummary = ""
        Exit Function
    End If

    Dim res As String, k As Variant
    For Each k In orphanCounts.Keys
        res = res & "  • " & CStr(k) & " (" & CStr(orphanCounts(k)) & " кол.)" & vbCrLf
    Next k
    FindOrphanColumnsSummary = res
End Function

Public Function FormatErrorList(ByVal errs As Collection, ByVal maxShow As Long) As String
    Dim s As String
    Dim i As Long, lim As Long
    lim = errs.Count
    If lim > maxShow Then lim = maxShow
    For i = 1 To lim
        s = s & "• " & errs(i) & vbCrLf
    Next i
    If errs.Count > maxShow Then
        s = s & "... и ещё " & (errs.Count - maxShow) & " замечаний."
    End If
    FormatErrorList = s
End Function

Public Sub SortStringCollection(ByVal coll As Collection)
    If coll.Count <= 1 Then Exit Sub
    Dim arr() As String
    ReDim arr(1 To coll.Count)
    Dim i As Long, j As Long
    For i = 1 To coll.Count
        arr(i) = coll(i)
    Next i

    For i = 1 To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If StrComp(arr(i), arr(j), vbTextCompare) > 0 Then
                Dim tmp As String
                tmp = arr(i)
                arr(i) = arr(j)
                arr(j) = tmp
            End If
        Next j
    Next i

    Do While coll.Count > 0
        coll.Remove 1
    Loop
    For i = 1 To UBound(arr)
        coll.Add arr(i)
    Next i
End Sub

Public Sub SortFilePathsByBaseName(ByVal coll As Collection, ByVal fso As Object)
    If coll.Count <= 1 Then Exit Sub
    Dim arr() As String
    ReDim arr(1 To coll.Count)
    Dim i As Long, j As Long
    For i = 1 To coll.Count
        arr(i) = coll(i)
    Next i

    For i = 1 To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If StrComp(fso.GetBaseName(arr(i)), fso.GetBaseName(arr(j)), vbTextCompare) > 0 Then
                Dim tmp As String
                tmp = arr(i)
                arr(i) = arr(j)
                arr(j) = tmp
            End If
        Next j
    Next i

    Do While coll.Count > 0
        coll.Remove 1
    Loop
    For i = 1 To UBound(arr)
        coll.Add arr(i)
    Next i
End Sub

Public Sub SetStatus(ByVal ws As Worksheet, ByVal msg As String)
    ws.Range(CELL_STATUS_TEXT).Value = msg
End Sub

Public Sub RestoreAppState()
    On Error Resume Next
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    On Error GoTo 0
End Sub

Public Sub EnsureNamedRangeSafe(ByVal nm As String, ByVal refersToFormula As String)
    On Error Resume Next
    ThisWorkbook.Names(nm).RefersTo = refersToFormula
    If Err.Number <> 0 Then
        Err.Clear
        ThisWorkbook.Names.Add Name:=nm, RefersTo:=refersToFormula
    End If
    On Error GoTo 0
End Sub

Public Function IntersectActiveSelection(ByVal ws As Worksheet, ByVal targetRng As Range) As Range
    Set IntersectActiveSelection = Nothing
    If targetRng Is Nothing Then Exit Function
    On Error Resume Next
    Dim selRange As Range
    If TypeName(Selection) = "Range" Then
        Set selRange = Selection
    ElseIf Not ActiveCell Is Nothing Then
        Set selRange = ActiveCell
    End If
    If Not selRange Is Nothing Then
        If selRange.Worksheet.Name = ws.Name Then
            Set IntersectActiveSelection = Application.Intersect(selRange, targetRng)
        End If
    End If
    On Error GoTo 0
End Function
