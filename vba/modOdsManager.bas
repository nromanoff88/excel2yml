Attribute VB_Name = "modOdsManager"
Option Explicit

' =============================================================================
' Модуль: modOdsManager
' Назначение: Управление интерфейсом листа Excel, загрузка схем и YAML-файлов,
'             редактирование строк и полная синхронизация обратно в YAML.
' =============================================================================

Public Const MAIN_SHEET_NAME As String = "ODS_Config"
Public Const LISTS_SHEET_NAME As String = "_Lists"
Public Const TBL_TABLES_NAME As String = "tblTables"
Public Const TBL_COLUMNS_NAME As String = "tblColumns"

Public Const CELL_CONFIG_PATH As String = "B1"
Public Const CELL_SCHEMA_NAME As String = "B2"
Public Const CELL_STATUS_TEXT As String = "B3"

Public Const TABLES_BANNER_ROW As Long = 5
Public Const TABLES_GROUP_ROW As Long = 6
Public Const TABLES_HEADER_ROW As Long = 7
Public Const TABLES_COL_COUNT As Long = 24
Public Const COLUMNS_COL_COUNT As Long = 7

' Флаг разового разрешения сохранения самого .xlsm файла (для администратора шаблона)
Public gAllowWorkbookSave As Boolean

' =============================================================================
' ПУБЛИЧНЫЕ МАКРОСЫ ДЛЯ КНОПОК ИНТЕРФЕЙСА
' =============================================================================

' Инициализация книги (можно вызывать при первом запуске или для восстановления разметки)
Public Sub InitOdsWorkspace()
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = EnsureMainSheet()
    EnsureListsSheet

    SetupTopControlPanel ws
    AutoDetectConfigFolder ws
    RefreshSchemasList False

    ' Если есть выбранная схема и таблиц ещё нет — загружаем автоматически
    If Len(Trim$(CellValueToString(ws.Range(CELL_SCHEMA_NAME).Value2))) > 0 Then
        If Not HasListObject(ws, TBL_TABLES_NAME) Then
            LoadSchemaFromYamlInternal False
        End If
    Else
        Dim emptyTables As Collection, emptyCols As Collection
        Set emptyTables = New Collection
        Set emptyCols = New Collection
        RenderBothTables ws, emptyTables, emptyCols
    End If
    Exit Sub

ErrHandler:
    MsgBox "Ошибка инициализации листа: " & Err.Description, vbCritical, "ODS YAML Manager"
End Sub

' Автоматическая проверка и инициализация без диалоговых окон (при открытии / автотесте)
Public Function RunSilentInitAndTest() As String
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = EnsureMainSheet()
    EnsureListsSheet

    SetupTopControlPanel ws
    AutoDetectConfigFolder ws
    RefreshSchemasList False

    If Len(Trim$(CellValueToString(ws.Range(CELL_SCHEMA_NAME).Value2))) > 0 Then
        LoadSchemaFromYamlInternal False
    End If

    Dim tablesColl As Collection, colsColl As Collection
    ReadSheetData ws, tablesColl, colsColl
    RunSilentInitAndTest = "OK: tables=" & tablesColl.Count & ", cols=" & colsColl.Count
    Exit Function

ErrHandler:
    RunSilentInitAndTest = "ERR: " & Err.Number & " - " & Err.Description
End Function

' Вызывается при открытии книги (обновляет путь к config и выпадающий список схем без перезаписи таблиц)
Public Sub AutoRefreshOnOpen()
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = EnsureMainSheet()
    EnsureListsSheet
    AutoDetectConfigFolder ws
    RefreshSchemasList False
    ThisWorkbook.Saved = True
    On Error GoTo 0
End Sub

' Разово разрешить сохранение самого файла .xlsm (если нужно обновить шаблон в Git)
Public Sub UnlockWorkbookSaveOnce()
    gAllowWorkbookSave = True
    MsgBox "Разовое сохранение книги Excel разрешено." & vbCrLf & _
           "Нажмите Ctrl + S, чтобы обновить эталонный файл .xlsm для Git.", _
           vbInformation, "Режим администратора шаблона"
End Sub

' Вызывается при выборе другой схемы в выпадающем списке B2
Public Sub OnSchemaSelectionChanged()
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = EnsureMainSheet()
    Dim selSchema As String
    selSchema = Trim$(CellValueToString(ws.Range(CELL_SCHEMA_NAME).Value2))
    If Len(selSchema) = 0 Then Exit Sub

    Dim cfgPath As String
    cfgPath = GetEffectiveConfigPath(ws)
    If Len(cfgPath) = 0 Then Exit Sub

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(cfgPath & "\" & selSchema) Then
        If MsgBox("Загрузить YAML-файлы выбранной схемы '" & selSchema & "'?" & vbCrLf & _
                  "(Текущие данные на листе будут заменены данными из папки схемы)", _
                  vbYesNo + vbQuestion, "Смена схемы") = vbYes Then
            LoadSchemaFromYamlInternal False
        End If
    End If
    On Error GoTo 0
End Sub

' Кнопка: Выбрать папку config
Public Sub SelectConfigFolder()
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = EnsureMainSheet()

    Dim fd As Object
    Set fd = Application.FileDialog(4) ' msoFileDialogFolderPicker
    fd.Title = "Выберите папку config (содержащую schema_ods_table.json и папки схем)"
    fd.AllowMultiSelect = False

    Dim curPath As String
    curPath = GetEffectiveConfigPath(ws)
    If Len(curPath) > 0 Then
        If Right$(curPath, 1) <> "\" Then curPath = curPath & "\"
        fd.InitialFileName = curPath
    ElseIf Len(ThisWorkbook.Path) > 0 Then
        fd.InitialFileName = ThisWorkbook.Path & "\"
    End If

    If fd.Show = -1 Then
        Dim chosen As String
        chosen = fd.SelectedItems(1)
        If Right$(chosen, 1) = "\" Then chosen = Left$(chosen, Len(chosen) - 1)
        If Len(ThisWorkbook.Path) > 0 And StrComp(chosen, ThisWorkbook.Path & "\config", vbTextCompare) = 0 Then
            ws.Range(CELL_CONFIG_PATH).Value = ".\config"
        Else
            ws.Range(CELL_CONFIG_PATH).Value = chosen
        End If
        RefreshSchemasList False

        Dim selSchema As String
        selSchema = Trim$(CellValueToString(ws.Range(CELL_SCHEMA_NAME).Value2))
        If Len(selSchema) > 0 Then
            If MsgBox("Загрузить YAML-файлы из схемы '" & selSchema & "'?", _
                      vbYesNo + vbQuestion, "Загрузка схемы") = vbYes Then
                LoadSchemaFromYamlInternal False
            End If
        End If
    End If
    Exit Sub

ErrHandler:
    MsgBox "Ошибка выбора папки: " & Err.Description, vbCritical, "ODS YAML Manager"
End Sub

' Кнопка: Обновить список схем (подпапок внутри config)
Public Sub RefreshSchemasButton()
    RefreshSchemasList True
End Sub

Public Sub RefreshSchemasList(Optional ByVal showNotification As Boolean = False)
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = EnsureMainSheet()
    Dim wsLists As Worksheet
    Set wsLists = EnsureListsSheet()

    AutoDetectConfigFolder ws
    Dim cfgPath As String
    cfgPath = GetEffectiveConfigPath(ws)

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    wsLists.Columns(1).ClearContents

    If Len(cfgPath) = 0 Or Not fso.FolderExists(cfgPath) Then
        RemoveValidationSafe ws.Range(CELL_SCHEMA_NAME)
        SetStatus ws, "Папка config не найдена. Нажмите 'Папка config...' для выбора."
        If showNotification Then
            MsgBox "Папка config не найдена по пути:" & vbCrLf & cfgPath, vbExclamation, "ODS YAML Manager"
        End If
        Exit Sub
    End If

    Dim fld As Object, subFld As Object
    Set fld = fso.GetFolder(cfgPath)

    Dim schemaNames As Collection
    Set schemaNames = New Collection

    For Each subFld In fld.SubFolders
        If Left$(subFld.Name, 1) <> "." Then
            schemaNames.Add subFld.Name
        End If
    Next subFld

    SortStringCollection schemaNames

    Dim prevSelected As String
    prevSelected = Trim$(CellValueToString(ws.Range(CELL_SCHEMA_NAME).Value2))
    Dim prevStillExists As Boolean
    prevStillExists = False

    If schemaNames.Count > 0 Then
        Dim i As Long
        For i = 1 To schemaNames.Count
            wsLists.Cells(i, 1).Value = schemaNames(i)
            If StrComp(schemaNames(i), prevSelected, vbTextCompare) = 0 Then
                prevStillExists = True
                prevSelected = schemaNames(i)
            End If
        Next i

        EnsureNamedRangeSafe "ListSchemas", "='" & LISTS_SHEET_NAME & "'!$A$1:$A$" & schemaNames.Count
        ApplyRangeValidation ws.Range(CELL_SCHEMA_NAME), "=ListSchemas", True

        Dim prevEvents As Boolean
        prevEvents = Application.EnableEvents
        Application.EnableEvents = False
        If prevStillExists Then
            ws.Range(CELL_SCHEMA_NAME).Value = prevSelected
        Else
            ws.Range(CELL_SCHEMA_NAME).Value = schemaNames(1)
        End If
        Application.EnableEvents = prevEvents

        SetStatus ws, "Найдено схем в config: " & schemaNames.Count & ". Выберите схему и нажмите 'Загрузить из YAML'."
        If showNotification Then
            MsgBox "Список схем обновлён. Найдено папок схем: " & schemaNames.Count, vbInformation, "ODS YAML Manager"
        End If
    Else
        RemoveValidationSafe ws.Range(CELL_SCHEMA_NAME)
        Application.EnableEvents = False
        ws.Range(CELL_SCHEMA_NAME).Value = ""
        Application.EnableEvents = True
        SetStatus ws, "В папке config нет подпапок со схемами."
        If showNotification Then
            MsgBox "В папке config не найдено ни одной подпапки схемы.", vbExclamation, "ODS YAML Manager"
        End If
    End If
    Exit Sub

ErrHandler:
    MsgBox "Ошибка обновления списка схем: " & Err.Description, vbCritical, "ODS YAML Manager"
End Sub

' Кнопка: Загрузить все YAML выбранной схемы в две таблицы
Public Sub LoadSchemaFromYaml()
    LoadSchemaFromYamlInternal True
End Sub

Public Sub LoadSchemaFromYamlInternal(Optional ByVal showDonePopup As Boolean = True)
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = EnsureMainSheet()
    EnsureListsSheet

    Dim schemaFolder As String
    schemaFolder = GetSelectedSchemaFolderPath(ws, True)
    If Len(schemaFolder) = 0 Then Exit Sub

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(schemaFolder) Then
        MsgBox "Папка выбранной схемы не существует:" & vbCrLf & schemaFolder, vbExclamation, "ODS YAML Manager"
        Exit Sub
    End If

    ' Собираем и сортируем имена YAML-файлов (*.yml, *.yaml)
    Dim filePaths As Collection
    Set filePaths = New Collection
    Dim fld As Object, fil As Object
    Set fld = fso.GetFolder(schemaFolder)

    For Each fil In fld.Files
        Dim ext As String
        ext = LCase$(fso.GetExtensionName(fil.Name))
        If (ext = "yml" Or ext = "yaml") And Left$(fil.Name, 1) <> "." Then
            filePaths.Add fil.Path
        End If
    Next fil

    SortFilePathsByBaseName filePaths, fso

    Dim tablesColl As Collection
    Set tablesColl = New Collection
    Dim allColsColl As Collection
    Set allColsColl = New Collection

    Dim idx As Long
    For idx = 1 To filePaths.Count
        Dim fp As String
        fp = filePaths(idx)
        Dim baseName As String
        baseName = fso.GetBaseName(fp)

        Dim tblDict As Object
        Set tblDict = ParseOdsYamlFile(fp, baseName)
        tablesColl.Add tblDict

        Dim tblCols As Collection
        Set tblCols = tblDict("columns")
        Dim cIdx As Long
        For cIdx = 1 To tblCols.Count
            allColsColl.Add tblCols(cIdx)
        Next cIdx
    Next idx

    RenderBothTables ws, tablesColl, allColsColl

    Dim schemaName As String
    schemaName = Trim$(CellValueToString(ws.Range(CELL_SCHEMA_NAME).Value2))
    SetStatus ws, "Загружена схема '" & schemaName & "': таблиц = " & tablesColl.Count & _
                  ", колонок = " & allColsColl.Count & " (" & Format$(Now, "dd.mm.yyyy hh:nn:ss") & ")"
    ThisWorkbook.Saved = True

    If showDonePopup Then
        MsgBox "Схема '" & schemaName & "' успешно загружена!" & vbCrLf & vbCrLf & _
               "Таблиц (YAML-файлов): " & tablesColl.Count & vbCrLf & _
               "Колонок всего: " & allColsColl.Count, vbInformation, "Загрузка завершена"
    End If
    Exit Sub

ErrHandler:
    RestoreAppState
    MsgBox "Ошибка при чтении YAML-файлов:" & vbCrLf & Err.Description, vbCritical, "ODS YAML Manager"
End Sub

' Кнопка: Сохранить изменения из Excel в YAML (полная синхронизация папки схемы)
Public Sub SaveSchemaToYaml()
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = EnsureMainSheet()

    Dim schemaName As String
    schemaName = Trim$(CellValueToString(ws.Range(CELL_SCHEMA_NAME).Value2))
    If Len(schemaName) = 0 Then
        MsgBox "Укажите имя схемы в ячейке " & CELL_SCHEMA_NAME & ".", vbExclamation, "ODS YAML Manager"
        Exit Sub
    End If

    If HasInvalidFileNameChars(schemaName) Then
        MsgBox "Имя схемы содержит недопустимые символы.", vbExclamation, "ODS YAML Manager"
        Exit Sub
    End If

    Dim cfgPath As String
    cfgPath = GetEffectiveConfigPath(ws)
    If Len(cfgPath) = 0 Then
        MsgBox "Сначала укажите путь к папке config в ячейке " & CELL_CONFIG_PATH & ".", vbExclamation, "ODS YAML Manager"
        Exit Sub
    End If

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(cfgPath) Then
        MsgBox "Папка config не найдена:" & vbCrLf & cfgPath, vbExclamation, "ODS YAML Manager"
        Exit Sub
    End If

    Dim schemaFolder As String
    schemaFolder = cfgPath & "\" & schemaName
    If Not fso.FolderExists(schemaFolder) Then
        If MsgBox("Папка схемы '" & schemaName & "' ещё не существует." & vbCrLf & _
                  "Создать новую папку схемы?" & vbCrLf & schemaFolder, _
                  vbYesNo + vbQuestion, "Создание новой схемы") = vbNo Then
            Exit Sub
        End If
        fso.CreateFolder schemaFolder
        RefreshSchemasList False
        Application.EnableEvents = False
        ws.Range(CELL_SCHEMA_NAME).Value = schemaName
        Application.EnableEvents = True
    End If

    ' Считываем данные обеих таблиц с листа
    Dim tablesColl As Collection
    Dim allColsColl As Collection
    ReadSheetData ws, tablesColl, allColsColl

    ' Проверяем целостность и валидацию по schema_ods_table.json
    Dim hasFatal As Boolean
    Dim valErrors As Collection
    Set valErrors = ValidateOdsData(tablesColl, allColsColl, hasFatal)

    If hasFatal Then
        MsgBox "Сохранение остановлено из-за критических ошибок:" & vbCrLf & vbCrLf & _
               FormatErrorList(valErrors, 15), vbCritical, "Ошибки данных"
        Exit Sub
    End If

    ' Проверка осиротевших колонок (колонки, у которых table_name нет в первой таблице)
    Dim orphanSummary As String
    orphanSummary = FindOrphanColumnsSummary(tablesColl, allColsColl)
    If Len(orphanSummary) > 0 Then
        If MsgBox("Внимание! Во второй таблице (Columns) есть строки для таблиц, " & _
                  "которых нет в первой таблице (Tables):" & vbCrLf & orphanSummary & vbCrLf & vbCrLf & _
                  "Эти колонки НЕ будут записаны ни в один YAML-файл." & vbCrLf & _
                  "Продолжить сохранение?", vbYesNo + vbExclamation, "Несвязанные колонки") = vbNo Then
            Exit Sub
        End If
    End If

    If valErrors.Count > 0 Then
        If MsgBox("Обнаружены несоответствия схеме schema_ods_table.json (" & valErrors.Count & " шт.):" & _
                  vbCrLf & vbCrLf & FormatErrorList(valErrors, 12) & vbCrLf & vbCrLf & _
                  "Всё равно сохранить файлы в YAML?", _
                  vbYesNo + vbExclamation, "Предупреждение валидации") = vbNo Then
            Exit Sub
        End If
    End If

    ' Группируем колонки по имени таблицы с сохранением порядка строк
    Dim colsByTable As Object
    Set colsByTable = CreateObject("Scripting.Dictionary")
    colsByTable.CompareMode = 1 ' vbTextCompare

    Dim tIdx As Long
    For tIdx = 1 To tablesColl.Count
        Dim tObj As Object
        Set tObj = tablesColl(tIdx)
        Dim tKey As String
        tKey = Trim$(CStr(tObj("table_name")))
        If Not colsByTable.Exists(tKey) Then
            Dim newColList As Collection
            Set newColList = New Collection
            colsByTable.Add tKey, newColList
        End If
    Next tIdx

    Dim cIdx As Long
    For cIdx = 1 To allColsColl.Count
        Dim cObj As Object
        Set cObj = allColsColl(cIdx)
        Dim cTblName As String
        cTblName = Trim$(CStr(cObj("table_name")))
        If colsByTable.Exists(cTblName) Then
            colsByTable(cTblName).Add cObj
        End If
    Next cIdx

    ' Сканируем существующие YAML-файлы на диске для полной синхронизации
    Dim existingFilesByBase As Object
    Set existingFilesByBase = CreateObject("Scripting.Dictionary")
    existingFilesByBase.CompareMode = 1 ' vbTextCompare

    Dim fld As Object, fil As Object
    Set fld = fso.GetFolder(schemaFolder)
    For Each fil In fld.Files
        Dim ext As String
        ext = LCase$(fso.GetExtensionName(fil.Name))
        If (ext = "yml" Or ext = "yaml") And Left$(fil.Name, 1) <> "." Then
            Dim bName As String
            bName = fso.GetBaseName(fil.Name)
            If Not existingFilesByBase.Exists(bName) Then
                existingFilesByBase.Add bName, fil.Path
            End If
        End If
    Next fil

    ' Определяем файлы к удалению (которых больше нет в tblTables)
    Dim filesToDelete As Collection
    Set filesToDelete = New Collection
    Dim diskKey As Variant
    For Each diskKey In existingFilesByBase.Keys
        If Not colsByTable.Exists(CStr(diskKey)) Then
            filesToDelete.Add CStr(existingFilesByBase(diskKey))
        End If
    Next diskKey

    ' Подтверждение удаления файлов с диска (защита от случайного удаления)
    If filesToDelete.Count > 0 Then
        Dim delNames As String
        Dim dIdx As Long
        For dIdx = 1 To filesToDelete.Count
            If dIdx <= 15 Then
                delNames = delNames & "  • " & fso.GetFileName(filesToDelete(dIdx)) & vbCrLf
            End If
        Next dIdx
        If filesToDelete.Count > 15 Then
            delNames = delNames & "  ... и ещё " & (filesToDelete.Count - 15) & " файл(ов)" & vbCrLf
        End If

        If MsgBox("Полная синхронизация схемы '" & schemaName & "':" & vbCrLf & vbCrLf & _
                  "Следующие YAML-файлы отсутствуют в таблице Excel и БУДУТ УДАЛЕНЫ с диска (" & _
                  filesToDelete.Count & " шт.):" & vbCrLf & delNames & vbCrLf & _
                  "Подтверждаете запись изменений и удаление этих файлов?", _
                  vbYesNo + vbExclamation + vbDefaultButton2, "Подтверждение удаления YAML") = vbNo Then
            Exit Sub
        End If
    End If

    ' Записываем YAML-файлы и удаляем лишние
    Dim createdCnt As Long, updatedCnt As Long, unchangedCnt As Long, deletedCnt As Long

    For tIdx = 1 To tablesColl.Count
        Dim curTbl As Object
        Set curTbl = tablesColl(tIdx)
        Dim curName As String
        curName = Trim$(CStr(curTbl("table_name")))

        Dim targetPath As String
        Dim isExisting As Boolean
        isExisting = existingFilesByBase.Exists(curName)
        If isExisting Then
            targetPath = CStr(existingFilesByBase(curName))
        Else
            targetPath = schemaFolder & "\" & curName & ".yml"
        End If

        Dim yamlLf As String
        yamlLf = SerializeOdsYaml(curTbl, colsByTable(curName))
        Dim yamlCrLf As String
        yamlCrLf = Replace(yamlLf, vbLf, vbCrLf)

        If isExisting Then
            Dim oldContent As String
            oldContent = ReadUtf8File(targetPath)
            If StrComp(oldContent, yamlCrLf, vbBinaryCompare) = 0 Then
                unchangedCnt = unchangedCnt + 1
            Else
                WriteUtf8NoBomFile targetPath, yamlLf
                updatedCnt = updatedCnt + 1
            End If
        Else
            WriteUtf8NoBomFile targetPath, yamlLf
            createdCnt = createdCnt + 1
        End If
    Next tIdx

    For dIdx = 1 To filesToDelete.Count
        If fso.FileExists(filesToDelete(dIdx)) Then
            fso.DeleteFile filesToDelete(dIdx), True
            deletedCnt = deletedCnt + 1
        End If
    Next dIdx

    SetStatus ws, "Сохранено в '" & schemaName & "': создано=" & createdCnt & _
                  ", обновлено=" & updatedCnt & ", без изменений=" & unchangedCnt & _
                  ", удалено=" & deletedCnt & " (" & Format$(Now, "dd.mm.yyyy hh:nn:ss") & ")"
    ThisWorkbook.Saved = True

    MsgBox "Синхронизация со схемой '" & schemaName & "' завершена!" & vbCrLf & vbCrLf & _
           "• Создано новых файлов: " & createdCnt & vbCrLf & _
           "• Обновлено файлов: " & updatedCnt & vbCrLf & _
           "• Без изменений: " & unchangedCnt & vbCrLf & _
           "• Удалено файлов: " & deletedCnt, vbInformation, "Сохранение выполнено"
    Exit Sub

ErrHandler:
    MsgBox "Ошибка при сохранении YAML-файлов:" & vbCrLf & Err.Description, vbCritical, "ODS YAML Manager"
End Sub
