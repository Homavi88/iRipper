Option Explicit

' ============================================================
' Экспорт уникальных строк F:M с листа "Ripper I" в архив.
' Привязка только к имени листа - имя файла источника не важно.
'
' Архив ВСЕГДА один и тот же файл (DEST_PATH): макрос только
' ДОПИСЫВАЕТ новые Batch Id, никогда не очищает и не
' перезаписывает уже сохранённые строки.
'
' ВАЖНО: не оставляйте в модуле строку Attribute VB_Name = ...
' Она нужна только при File > Import File, при вставке кода
' даёт Compile error: Expected: line number or label...
' ============================================================

' Путь к файлу-архиву (фиксированный; не зависит от имени исходной книги)
Private Const DEST_PATH As String = "C:\AUTO\iRipper\Ripper_I_PackedBatches.xlsx"

' Имя листа-источника (фиксировано; имя книги может меняться)
Private Const SOURCE_SHEET_NAME As String = "Ripper I"

' Имя листа в файле-архиве
Private Const DEST_SHEET_NAME As String = "Ripper I"

' Диапазон F:M
Private Const COL_FIRST As Long = 6   ' F - Packed batch
Private Const COL_LAST As Long = 13   ' M - Ripper weight
Private Const COL_PACKED As Long = 6  ' F - критерий "есть значение"
Private Const COL_BATCH As Long = 10  ' J - Batch Id (уникальность)

' В архиве колонки сдвигаются: F->A ... M->H, Batch Id оказывается в колонке E
Private Const DEST_BATCH_COL As Long = 5
Private Const DEST_COL_COUNT As Long = 8 ' F:M = 8 колонок

' ------------------------------------------------------------
' Точка входа: запустите этот макрос (Alt+F8)
' ------------------------------------------------------------
Public Sub ExportUniquePackedBatches()
    Dim srcWb As Workbook
    Dim srcWs As Worksheet
    Dim destWb As Workbook
    Dim destWs As Worksheet
    Dim existing As Object
    Dim lastSrcRow As Long
    Dim lastDestRow As Long
    Dim r As Long
    Dim destRow As Long
    Dim batchKey As String
    Dim packedVal As Variant
    Dim added As Long
    Dim skippedEmpty As Long
    Dim skippedDup As Long
    Dim createdNew As Boolean
    Dim weOpenedDest As Boolean
    Dim screenWasUpdating As Boolean
    Dim calcMode As XlCalculation

    On Error GoTo CleanFail

    screenWasUpdating = Application.ScreenUpdating
    calcMode = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    weOpenedDest = False
    createdNew = False

    Set srcWb = ThisWorkbook
    Set srcWs = FindSheetByName(srcWb, SOURCE_SHEET_NAME)
    If srcWs Is Nothing Then
        MsgBox "Лист """ & SOURCE_SHEET_NAME & """ не найден в текущей книге.", _
               vbExclamation, "Export Ripper I"
        GoTo CleanExit
    End If

    lastSrcRow = LastUsedRowInColumns(srcWs, COL_FIRST, COL_LAST)
    If lastSrcRow < 2 Then
        MsgBox "На листе """ & SOURCE_SHEET_NAME & """ нет данных для копирования.", _
               vbInformation, "Export Ripper I"
        GoTo CleanExit
    End If

    ' Один и тот же архив независимо от имени исходной книги
    Set destWb = OpenOrCreateArchive(DEST_PATH, createdNew, weOpenedDest)
    Set destWs = EnsureDestSheet(destWb, DEST_SHEET_NAME)

    ' Никогда не очищаем лист: только дописываем после последней строки
    lastDestRow = LastDataRow(destWs)
    If lastDestRow < 1 Then
        CopyHeaderAsValues srcWs, destWs
        lastDestRow = 1
    End If

    Set existing = BuildBatchIdMap(destWs, lastDestRow)

    destRow = lastDestRow
    added = 0
    skippedEmpty = 0
    skippedDup = 0

    For r = 2 To lastSrcRow
        packedVal = srcWs.Cells(r, COL_PACKED).Value2
        If IsBlankCell(packedVal) Then
            skippedEmpty = skippedEmpty + 1
        Else
            batchKey = NormalizeBatchId(srcWs.Cells(r, COL_BATCH).Value2)
            If Len(batchKey) = 0 Then
                skippedEmpty = skippedEmpty + 1
            ElseIf existing.Exists(batchKey) Then
                skippedDup = skippedDup + 1
            Else
                destRow = destRow + 1
                ' Копируем только значения (формулы не переносятся)
                destWs.Range(destWs.Cells(destRow, 1), destWs.Cells(destRow, DEST_COL_COUNT)).Value2 = _
                    srcWs.Range(srcWs.Cells(r, COL_FIRST), srcWs.Cells(r, COL_LAST)).Value2
                existing.Add batchKey, True
                added = added + 1
            End If
        End If
    Next r

    ' Только Save - никогда SaveAs поверх существующего архива
    destWb.Save

    MsgBox "Готово (дозапись в общий архив)." & vbCrLf & _
           "Исходная книга: " & srcWb.Name & vbCrLf & _
           "Добавлено строк: " & added & vbCrLf & _
           "Пропущено (пустое Packed batch / Batch Id): " & skippedEmpty & vbCrLf & _
           "Пропущено (уже есть Batch Id): " & skippedDup & vbCrLf & _
           "Всего строк в архиве (с шапкой): " & destRow & vbCrLf & _
           "Файл: " & destWb.FullName, _
           vbInformation, "Export Ripper I"

CleanExit:
    On Error Resume Next
    If weOpenedDest Then
        If Not destWb Is Nothing Then
            If destWb.FullName <> srcWb.FullName Then destWb.Close SaveChanges:=False
        End If
    End If
    On Error GoTo 0
    Application.EnableEvents = True
    Application.Calculation = calcMode
    Application.ScreenUpdating = screenWasUpdating
    Exit Sub

CleanFail:
    On Error Resume Next
    If weOpenedDest Then
        If Not destWb Is Nothing Then
            If destWb.FullName <> srcWb.FullName Then destWb.Close SaveChanges:=False
        End If
    End If
    On Error GoTo 0
    Application.EnableEvents = True
    Application.Calculation = calcMode
    Application.ScreenUpdating = screenWasUpdating
    MsgBox "Ошибка VBA " & Err.Number & ": " & Err.Description, vbCritical, "Export Ripper I"
End Sub

' ======================== Вспомогательные ========================

Private Function FindSheetByName(wb As Workbook, sheetName As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If StrComp(ws.Name, sheetName, vbTextCompare) = 0 Then
            Set FindSheetByName = ws
            Exit Function
        End If
    Next ws
    Set FindSheetByName = Nothing
End Function

' Открывает существующий архив или создаёт новый.
' НИКОГДА не перезаписывает уже существующий файл через SaveAs.
Private Function OpenOrCreateArchive(ByVal path As String, _
                                     ByRef createdNew As Boolean, _
                                     ByRef weOpened As Boolean) As Workbook
    Dim wb As Workbook
    Dim fso As Object
    Dim normalized As String
    Dim saveErrNum As Long
    Dim saveErrDesc As String

    createdNew = False
    weOpened = False
    normalized = NormalizePath(path)
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' 1) Уже открыт в Excel - используем его (не открываем второй раз)
    Set wb = FindOpenWorkbookByPath(normalized)
    If Not wb Is Nothing Then
        Set OpenOrCreateArchive = wb
        Exit Function
    End If

    ' 2) Файл на диске есть - только Open, без SaveAs
    If ArchiveFileExists(normalized, fso) Then
        Set wb = Workbooks.Open(Filename:=normalized, UpdateLinks:=0, ReadOnly:=False, IgnoreReadOnlyRecommended:=True)
        weOpened = True
        Set OpenOrCreateArchive = wb
        Exit Function
    End If

    ' 3) Файла нет - создаём новый и сохраняем по фиксированному пути
    EnsureParentFolder normalized, fso
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Application.DisplayAlerts = False
    On Error GoTo SaveAsFail
    wb.SaveAs Filename:=normalized, FileFormat:=xlOpenXMLWorkbook
    On Error GoTo 0
    Application.DisplayAlerts = True
    createdNew = True
    weOpened = True
    Set OpenOrCreateArchive = wb
    Exit Function

SaveAsFail:
    saveErrNum = Err.Number
    saveErrDesc = Err.Description
    Application.DisplayAlerts = True
    ' Если между проверкой и SaveAs файл появился - открываем его, не затираем
    On Error Resume Next
    wb.Close SaveChanges:=False
    On Error GoTo 0
    If ArchiveFileExists(normalized, fso) Then
        Set wb = Workbooks.Open(Filename:=normalized, UpdateLinks:=0, ReadOnly:=False)
        weOpened = True
        createdNew = False
        Set OpenOrCreateArchive = wb
        Exit Function
    End If
    Err.Raise saveErrNum, "OpenOrCreateArchive", _
              "Не удалось создать архив: " & normalized & " (" & saveErrDesc & ")"
End Function

Private Function FindOpenWorkbookByPath(ByVal path As String) As Workbook
    Dim wb As Workbook
    Dim fullName As String
    Dim target As String

    target = UCase$(NormalizePath(path))
    For Each wb In Application.Workbooks
        On Error Resume Next
        fullName = wb.FullName
        On Error GoTo 0
        If Len(fullName) > 0 Then
            If UCase$(NormalizePath(fullName)) = target Then
                Set FindOpenWorkbookByPath = wb
                Exit Function
            End If
        End If
    Next wb
    Set FindOpenWorkbookByPath = Nothing
End Function

Private Function ArchiveFileExists(ByVal path As String, ByVal fso As Object) As Boolean
    On Error Resume Next
    ArchiveFileExists = fso.FileExists(path)
    If Not ArchiveFileExists Then
        ArchiveFileExists = (Len(Dir$(path)) > 0)
    End If
    On Error GoTo 0
End Function

Private Function NormalizePath(ByVal path As String) As String
    NormalizePath = Replace(Trim$(path), "/", "\")
End Function

Private Sub EnsureParentFolder(ByVal path As String, ByVal fso As Object)
    Dim folder As String
    folder = fso.GetParentFolderName(path)
    If Len(folder) > 0 Then
        If Not fso.FolderExists(folder) Then
            fso.CreateFolder folder
        End If
    End If
End Sub

Private Function EnsureDestSheet(wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    Set ws = FindSheetByName(wb, sheetName)
    If ws Is Nothing Then
        ' Для новой книги переименуем единственный лист, чтобы не плодить пустые
        If wb.Worksheets.Count = 1 And LastDataRow(wb.Worksheets(1)) < 1 Then
            Set ws = wb.Worksheets(1)
            On Error Resume Next
            ws.Name = sheetName
            On Error GoTo 0
        Else
            Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
            On Error Resume Next
            ws.Name = sheetName
            On Error GoTo 0
        End If
    End If
    Set EnsureDestSheet = ws
End Function

' Последняя строка с данными; 0 = лист пустой (шапки нет)
Private Function LastDataRow(ws As Worksheet) As Long
    Dim lastRow As Long
    lastRow = LastUsedRowInColumns(ws, 1, DEST_COL_COUNT)
    If lastRow <= 1 And IsBlankCell(ws.Cells(1, 1).Value2) And IsBlankCell(ws.Cells(1, DEST_BATCH_COL).Value2) Then
        LastDataRow = 0
    Else
        LastDataRow = lastRow
    End If
End Function

Private Sub CopyHeaderAsValues(srcWs As Worksheet, destWs As Worksheet)
    destWs.Range(destWs.Cells(1, 1), destWs.Cells(1, DEST_COL_COUNT)).Value2 = _
        srcWs.Range(srcWs.Cells(1, COL_FIRST), srcWs.Cells(1, COL_LAST)).Value2
End Sub

Private Function BuildBatchIdMap(ws As Worksheet, ByVal lastRow As Long) As Object
    Dim dict As Object
    Dim r As Long
    Dim key As String

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1 ' TextCompare

    If lastRow >= 2 Then
        For r = 2 To lastRow
            key = NormalizeBatchId(ws.Cells(r, DEST_BATCH_COL).Value2)
            If Len(key) > 0 Then
                If Not dict.Exists(key) Then dict.Add key, True
            End If
        Next r
    End If

    Set BuildBatchIdMap = dict
End Function

Private Function LastUsedRowInColumns(ws As Worksheet, ByVal colFrom As Long, ByVal colTo As Long) As Long
    Dim c As Long
    Dim lastRow As Long
    Dim found As Long

    lastRow = 1
    For c = colFrom To colTo
        found = ws.Cells(ws.Rows.Count, c).End(xlUp).Row
        If found > lastRow Then lastRow = found
    Next c
    LastUsedRowInColumns = lastRow
End Function

Private Function IsBlankCell(ByVal v As Variant) As Boolean
    If IsError(v) Then
        IsBlankCell = True
    ElseIf IsNull(v) Or IsEmpty(v) Then
        IsBlankCell = True
    ElseIf VarType(v) = vbString Then
        IsBlankCell = (Len(Trim$(CStr(v))) = 0)
    Else
        IsBlankCell = False
    End If
End Function

Private Function NormalizeBatchId(ByVal v As Variant) As String
    If IsBlankCell(v) Then
        NormalizeBatchId = ""
    ElseIf IsNumeric(v) Then
        ' Числа без экспоненты / хвоста .0
        If Fix(CDbl(v)) = CDbl(v) Then
            NormalizeBatchId = CStr(CLng(v))
        Else
            NormalizeBatchId = Trim$(CStr(v))
        End If
    Else
        NormalizeBatchId = Trim$(CStr(v))
    End If
End Function
