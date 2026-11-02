Option Explicit

' ============================================================
' Экспорт уникальных строк F:M с листа "Ripper I" в архив.
' Привязка только к имени листа - имя файла источника не важно.
'
' ВАЖНО: не оставляйте в модуле строку Attribute VB_Name = ...
' Она нужна только при File > Import File, при вставке кода
' даёт Compile error: Expected: line number or label...
' ============================================================

' Путь к файлу-архиву (измените при необходимости)
Private Const DEST_PATH As String = "C:\AUTO\iRipper\Ripper_I_PackedBatches.xlsx"

' Имя листа-источника (фиксировано; имя книги может меняться)
Private Const SOURCE_SHEET_NAME As String = "Ripper I"

' Имя листа в файле-архиве
Private Const DEST_SHEET_NAME As String = "Ripper I"

' Диапазон F:M
Private Const COL_FIRST As Long = 6   ' F — Packed batch
Private Const COL_LAST As Long = 13   ' M — Ripper weight
Private Const COL_PACKED As Long = 6  ' F — критерий «есть значение»
Private Const COL_BATCH As Long = 10  ' J — Batch Id (уникальность)

' В архиве колонки сдвигаются: F->A ... M->H, Batch Id оказывается в колонке E
Private Const DEST_BATCH_COL As Long = 5

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
    Dim screenWasUpdating As Boolean
    Dim calcMode As XlCalculation

    On Error GoTo CleanFail

    screenWasUpdating = Application.ScreenUpdating
    calcMode = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

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

    Set destWb = OpenOrCreateArchive(DEST_PATH, createdNew)
    Set destWs = EnsureDestSheet(destWb, DEST_SHEET_NAME)

    ' Шапка: копируем только если архив пустой / лист новый
    If WorksheetIsEmpty(destWs) Then
        CopyHeaderAsValues srcWs, destWs
        lastDestRow = 1
    Else
        lastDestRow = LastUsedRowInColumns(destWs, 1, COL_LAST - COL_FIRST + 1)
        If lastDestRow < 1 Then lastDestRow = 1
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
                destWs.Range(destWs.Cells(destRow, 1), destWs.Cells(destRow, COL_LAST - COL_FIRST + 1)).Value2 = _
                    srcWs.Range(srcWs.Cells(r, COL_FIRST), srcWs.Cells(r, COL_LAST)).Value2
                existing.Add batchKey, True
                added = added + 1
            End If
        End If
    Next r

    destWb.Save
    If createdNew Or destWb.FullName <> srcWb.FullName Then
        ' Архив остаётся открытым, чтобы сразу видеть результат; при желании закройте вручную
    End If

    MsgBox "Готово." & vbCrLf & _
           "Добавлено строк: " & added & vbCrLf & _
           "Пропущено (пустое Packed batch / Batch Id): " & skippedEmpty & vbCrLf & _
           "Пропущено (уже есть Batch Id): " & skippedDup & vbCrLf & _
           "Файл: " & DEST_PATH, _
           vbInformation, "Export Ripper I"

CleanExit:
    Application.EnableEvents = True
    Application.Calculation = calcMode
    Application.ScreenUpdating = screenWasUpdating
    Exit Sub

CleanFail:
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

Private Function OpenOrCreateArchive(ByVal path As String, ByRef createdNew As Boolean) As Workbook
    Dim wb As Workbook
    Dim fso As Object

    createdNew = False
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FileExists(path) Then
        Set wb = Workbooks.Open(Filename:=path, UpdateLinks:=0, ReadOnly:=False)
    Else
        ' Создаём папку, если её ещё нет (например M:\ уже смонтирован)
        EnsureParentFolder path, fso
        Set wb = Workbooks.Add
        Application.DisplayAlerts = False
        wb.SaveAs Filename:=path, FileFormat:=xlOpenXMLWorkbook
        Application.DisplayAlerts = True
        createdNew = True
    End If

    Set OpenOrCreateArchive = wb
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
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        On Error Resume Next
        ws.Name = sheetName
        On Error GoTo 0
    End If
    Set EnsureDestSheet = ws
End Function

Private Function WorksheetIsEmpty(ws As Worksheet) As Boolean
    WorksheetIsEmpty = (ws.UsedRange.Cells.CountLarge <= 1 And Len(CStr(ws.Cells(1, 1).Value2 & "")) = 0)
End Function

Private Sub CopyHeaderAsValues(srcWs As Worksheet, destWs As Worksheet)
    destWs.Range(destWs.Cells(1, 1), destWs.Cells(1, COL_LAST - COL_FIRST + 1)).Value2 = _
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
