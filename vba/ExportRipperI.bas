Option Explicit

' ============================================================
' ExportRipperI - v3
' Экспорт F:M с листа "Ripper I" в ОДИН общий архив.
' Имя исходной книги не важно - архив только ДОПИСЫВАЕТСЯ.
'
' Если в MsgBox нет строки Version: v3 - в Excel всё ещё старый код.
' ============================================================

Private Const MACRO_VERSION As String = "v3-append-2026-11-02"

' Фиксированный путь архива (не зависит от имени исходной книги)
Private Const DEST_PATH As String = "C:\AUTO\iRipper\Ripper_I_PackedBatches.xlsx"

Private Const SOURCE_SHEET_NAME As String = "Ripper I"
Private Const DEST_SHEET_NAME As String = "Ripper I"

Private Const COL_FIRST As Long = 6   ' F - Packed batch
Private Const COL_LAST As Long = 13   ' M - Ripper weight
Private Const COL_PACKED As Long = 6
Private Const COL_BATCH As Long = 10  ' J - Batch Id

' В архиве: F->A ... J->E ... M->H
Private Const DEST_BATCH_COL As Long = 5
Private Const DEST_COL_COUNT As Long = 8

Public Sub ExportUniquePackedBatches()
    Dim srcWb As Workbook
    Dim srcWs As Worksheet
    Dim destWb As Workbook
    Dim destWs As Worksheet
    Dim existing As Object
    Dim lastSrcRow As Long
    Dim lastDestRow As Long
    Dim rowsBefore As Long
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
    Dim writeCell As Range

    On Error GoTo CleanFail

    screenWasUpdating = Application.ScreenUpdating
    calcMode = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    weOpenedDest = False
    createdNew = False
    Set destWb = Nothing

    Set srcWb = ThisWorkbook
    Set srcWs = FindSheetByName(srcWb, SOURCE_SHEET_NAME)
    If srcWs Is Nothing Then
        MsgBox "Version: " & MACRO_VERSION & vbCrLf & _
               "Лист """ & SOURCE_SHEET_NAME & """ не найден.", _
               vbExclamation, "Export Ripper I"
        GoTo CleanExit
    End If

    lastSrcRow = LastNonEmptyRow(srcWs, COL_FIRST, COL_LAST)
    If lastSrcRow < 2 Then
        MsgBox "Version: " & MACRO_VERSION & vbCrLf & _
               "Нет данных на листе """ & SOURCE_SHEET_NAME & """.", _
               vbInformation, "Export Ripper I"
        GoTo CleanExit
    End If

    Set destWb = OpenArchiveAppendOnly(DEST_PATH, createdNew, weOpenedDest)
    Set destWs = EnsureDestSheet(destWb, DEST_SHEET_NAME)

    lastDestRow = LastNonEmptyRow(destWs, 1, DEST_COL_COUNT)
    rowsBefore = lastDestRow

    If lastDestRow < 1 Then
        CopyHeaderAsValues srcWs, destWs
        lastDestRow = 1
        rowsBefore = 1
    End If

    Set existing = BuildBatchIdMap(destWs, lastDestRow)

    ' Пишем ТОЛЬКО ниже уже занятых строк - никогда поверх
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
                destRow = NextEmptyArchiveRow(destWs, destRow)
                If destRow <= rowsBefore And rowsBefore > 1 Then
                    Err.Raise vbObjectError + 1, "ExportUniquePackedBatches", _
                              "Защита: попытка записи в занятую строку " & destRow & _
                              " (в архиве уже " & rowsBefore & " строк). Операция отменена."
                End If

                Set writeCell = destWs.Cells(destRow, 1)
                ' Ещё раз: не затираем непустую строку
                If Not IsBlankCell(writeCell.Value2) Or Not IsBlankCell(destWs.Cells(destRow, DEST_BATCH_COL).Value2) Then
                    Err.Raise vbObjectError + 2, "ExportUniquePackedBatches", _
                              "Защита: строка " & destRow & " в архиве не пустая. Операция отменена."
                End If

                destWs.Range(destWs.Cells(destRow, 1), destWs.Cells(destRow, DEST_COL_COUNT)).Value2 = _
                    srcWs.Range(srcWs.Cells(r, COL_FIRST), srcWs.Cells(r, COL_LAST)).Value2

                existing.Add batchKey, True
                added = added + 1
            End If
        End If
    Next r

    ' Контроль: число строк не должно уменьшиться
    If LastNonEmptyRow(destWs, 1, DEST_COL_COUNT) < rowsBefore Then
        Err.Raise vbObjectError + 3, "ExportUniquePackedBatches", _
                  "Защита: число строк в архиве уменьшилось. Сохранение отменено."
    End If

    destWb.Save

    MsgBox "Version: " & MACRO_VERSION & vbCrLf & _
           "Режим: только дозапись (без перезаписи)" & vbCrLf & _
           "Исходная книга: " & srcWb.Name & vbCrLf & _
           "Строк в архиве до: " & rowsBefore & vbCrLf & _
           "Добавлено: " & added & vbCrLf & _
           "Пропущено (пусто): " & skippedEmpty & vbCrLf & _
           "Пропущено (дубликат Batch Id): " & skippedDup & vbCrLf & _
           "Строк в архиве после: " & LastNonEmptyRow(destWs, 1, DEST_COL_COUNT) & vbCrLf & _
           "Архив: " & destWb.FullName, _
           vbInformation, "Export Ripper I"

CleanExit:
    On Error Resume Next
    If weOpenedDest Then
        If Not destWb Is Nothing Then
            If Not srcWb Is Nothing Then
                If destWb.FullName <> srcWb.FullName Then destWb.Close SaveChanges:=False
            End If
        End If
    End If
    On Error GoTo 0
    Application.EnableEvents = True
    Application.Calculation = calcMode
    Application.ScreenUpdating = screenWasUpdating
    Exit Sub

CleanFail:
    Dim errNum As Long
    Dim errDesc As String
    errNum = Err.Number
    errDesc = Err.Description
    On Error Resume Next
    If weOpenedDest Then
        If Not destWb Is Nothing Then
            If Not srcWb Is Nothing Then
                If destWb.FullName <> srcWb.FullName Then destWb.Close SaveChanges:=False
            End If
        End If
    End If
    On Error GoTo 0
    Application.EnableEvents = True
    Application.Calculation = calcMode
    Application.ScreenUpdating = screenWasUpdating
    MsgBox "Version: " & MACRO_VERSION & vbCrLf & _
           "Ошибка VBA " & errNum & ": " & errDesc, vbCritical, "Export Ripper I"
End Sub

' ======================== Архив: только open/append ========================

' 1) Уже открыт  2) Пробуем Open  3) Создаём ТОЛЬКО если файла нет
' SaveAs поверх существующего файла - ЗАПРЕЩЁН
Private Function OpenArchiveAppendOnly(ByVal path As String, _
                                       ByRef createdNew As Boolean, _
                                       ByRef weOpened As Boolean) As Workbook
    Dim wb As Workbook
    Dim fso As Object
    Dim normalized As String
    Dim openErr As Long
    Dim openDesc As String

    createdNew = False
    weOpened = False
    normalized = NormalizePath(path)
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set wb = Nothing

    Set wb = FindOpenWorkbookByPath(normalized)
    If Not wb Is Nothing Then
        Set OpenArchiveAppendOnly = wb
        Exit Function
    End If

    ' Главный путь: пытаемся ОТКРЫТЬ. Если файл есть - он откроется.
    On Error Resume Next
    Set wb = Workbooks.Open(Filename:=normalized, _
                            UpdateLinks:=0, _
                            ReadOnly:=False, _
                            IgnoreReadOnlyRecommended:=True)
    openErr = Err.Number
    openDesc = Err.Description
    On Error GoTo 0

    If Not wb Is Nothing Then
        weOpened = True
        createdNew = False
        Set OpenArchiveAppendOnly = wb
        Exit Function
    End If

    ' Файл есть на диске, но Open не удался - НЕ создаём новый и НЕ SaveAs
    If FileExistsOnDisk(normalized, fso) Then
        Err.Raise vbObjectError + 10, "OpenArchiveAppendOnly", _
                  "Архив существует, но не удалось открыть:" & vbCrLf & normalized & vbCrLf & _
                  openDesc & vbCrLf & _
                  "Закройте файл если он занят другим процессом. Перезапись запрещена."
    End If

    ' Файла точно нет - создаём новый
    EnsureParentFolder normalized, fso

    If FileExistsOnDisk(normalized, fso) Then
        Err.Raise vbObjectError + 11, "OpenArchiveAppendOnly", _
                  "Архив появился на диске. Перезапись запрещена. Откройте его вручную:" & vbCrLf & normalized
    End If

    Set wb = Workbooks.Add(xlWBATWorksheet)
    On Error GoTo SaveFail
    wb.SaveAs Filename:=normalized, FileFormat:=xlOpenXMLWorkbook
    On Error GoTo 0

    createdNew = True
    weOpened = True
    Set OpenArchiveAppendOnly = wb
    Exit Function

SaveFail:
    openErr = Err.Number
    openDesc = Err.Description
    On Error Resume Next
    wb.Close SaveChanges:=False
    On Error GoTo 0
    Err.Raise openErr, "OpenArchiveAppendOnly", _
              "Не удалось создать архив (без перезаписи): " & normalized & vbCrLf & openDesc
End Function

Private Function FindOpenWorkbookByPath(ByVal path As String) As Workbook
    Dim wb As Workbook
    Dim fullName As String
    Dim target As String
    Dim targetName As String

    target = UCase$(NormalizePath(path))
    targetName = UCase$(FileNameOnly(path))

    For Each wb In Application.Workbooks
        fullName = vbNullString
        On Error Resume Next
        fullName = wb.FullName
        On Error GoTo 0
        If Len(fullName) > 0 Then
            If UCase$(NormalizePath(fullName)) = target Then
                Set FindOpenWorkbookByPath = wb
                Exit Function
            End If
            ' Запасной критерий: то же имя файла в той же папке AUTO\iRipper
            If UCase$(FileNameOnly(fullName)) = targetName Then
                If InStr(1, UCase$(NormalizePath(fullName)), "\AUTO\IRIPPER\", vbBinaryCompare) > 0 Then
                    Set FindOpenWorkbookByPath = wb
                    Exit Function
                End If
            End If
        End If
    Next wb
    Set FindOpenWorkbookByPath = Nothing
End Function

Private Function FileExistsOnDisk(ByVal path As String, ByVal fso As Object) As Boolean
    Dim d As String
    On Error Resume Next
    FileExistsOnDisk = fso.FileExists(path)
    If Not FileExistsOnDisk Then
        d = Dir$(path)
        FileExistsOnDisk = (Len(d) > 0)
    End If
    On Error GoTo 0
End Function

Private Function NormalizePath(ByVal path As String) As String
    NormalizePath = Replace(Trim$(path), "/", "\")
End Function

Private Function FileNameOnly(ByVal path As String) As String
    Dim p As String
    Dim pos As Long
    p = NormalizePath(path)
    pos = InStrRev(p, "\")
    If pos > 0 Then
        FileNameOnly = Mid$(p, pos + 1)
    Else
        FileNameOnly = p
    End If
End Function

Private Sub EnsureParentFolder(ByVal path As String, ByVal fso As Object)
    Dim folder As String
    folder = fso.GetParentFolderName(path)
    If Len(folder) = 0 Then Exit Sub
    If Not fso.FolderExists(folder) Then
        fso.CreateFolder folder
    End If
End Sub

Private Function EnsureDestSheet(wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    Set ws = FindSheetByName(wb, sheetName)
    If ws Is Nothing Then
        If wb.Worksheets.Count = 1 And LastNonEmptyRow(wb.Worksheets(1), 1, DEST_COL_COUNT) < 1 Then
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

Private Function FindSheetByName(wb As Workbook, sheetName As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If StrComp(Trim$(ws.Name), Trim$(sheetName), vbTextCompare) = 0 Then
            Set FindSheetByName = ws
            Exit Function
        End If
    Next ws
    Set FindSheetByName = Nothing
End Function

' ======================== Строки / значения ========================

Private Function LastNonEmptyRow(ws As Worksheet, ByVal colFrom As Long, ByVal colTo As Long) As Long
    Dim rng As Range
    Dim area As Range
    Dim lastRow As Long
    Dim c As Long
    Dim found As Long

    lastRow = 0

    On Error Resume Next
    Set rng = ws.Range(ws.Cells(1, colFrom), ws.Cells(ws.Rows.Count, colTo)).Find( _
                What:="*", _
                LookIn:=xlFormulas, _
                LookAt:=xlPart, _
                SearchOrder:=xlByRows, _
                SearchDirection:=xlPrevious, _
                MatchCase:=False)
    On Error GoTo 0

    If Not rng Is Nothing Then
        lastRow = rng.Row
    End If

    ' Дополнительно по каждой колонке (надёжнее после фильтров / UsedRange)
    For c = colFrom To colTo
        found = ws.Cells(ws.Rows.Count, c).End(xlUp).Row
        If Not IsBlankCell(ws.Cells(found, c).Value2) Then
            If found > lastRow Then lastRow = found
        End If
    Next c

    ' UsedRange как третий сигнал
    On Error Resume Next
    Set area = ws.UsedRange
    On Error GoTo 0
    If Not area Is Nothing Then
        found = area.Row + area.Rows.Count - 1
        If found > lastRow Then
            ' UsedRange иногда раздут - принимаем только если в этой строке есть значение в наших колонках
            For c = colFrom To colTo
                If Not IsBlankCell(ws.Cells(found, c).Value2) Then
                    lastRow = found
                    Exit For
                End If
            Next c
        End If
    End If

    LastNonEmptyRow = lastRow
End Function

Private Function NextEmptyArchiveRow(ws As Worksheet, ByVal afterRow As Long) As Long
    Dim r As Long
    r = afterRow + 1
    If r < 2 Then r = 2
    Do While Not IsBlankCell(ws.Cells(r, 1).Value2) Or Not IsBlankCell(ws.Cells(r, DEST_BATCH_COL).Value2)
        r = r + 1
        If r >= ws.Rows.Count Then
            Err.Raise vbObjectError + 12, "NextEmptyArchiveRow", "Нет свободных строк в архиве."
        End If
    Loop
    NextEmptyArchiveRow = r
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
    dict.CompareMode = 1

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
        If Fix(CDbl(v)) = CDbl(v) Then
            NormalizeBatchId = CStr(CLng(v))
        Else
            NormalizeBatchId = Trim$(CStr(v))
        End If
    Else
        NormalizeBatchId = Trim$(CStr(v))
    End If
End Function
