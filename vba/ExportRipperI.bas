Option Explicit

' ============================================================
' ExportRipperI - v5
' Экспорт F:M с листа "Ripper I" в ОДИН общий архив.
' Имя исходной книги не важно - архив только ДОПИСЫВАЕТСЯ.
'
' v5: исправлен Type mismatch на шаге find source sheet
'     (безопасный поиск листа + ThisWorkbook/ActiveWorkbook)
' v4: Type mismatch при разборе Batch Id / spill
' ============================================================

Private Const MACRO_VERSION As String = "v5-append-2026-11-23"

Private Const DEST_PATH As String = "C:\AUTO\iRipper\Ripper_I_PackedBatches.xlsx"

Private Const SOURCE_SHEET_NAME As String = "Ripper I"
Private Const DEST_SHEET_NAME As String = "Ripper I"

Private Const COL_FIRST As Long = 6   ' F - Packed batch
Private Const COL_LAST As Long = 13   ' M - Ripper weight
Private Const COL_PACKED As Long = 6
Private Const COL_BATCH As Long = 10  ' J - Batch Id

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
    Dim calcMode As Long
    Dim stepName As String
    Dim errNum As Long
    Dim errDesc As String
    Dim sheetList As String

    On Error GoTo CleanFail

    stepName = "init"
    screenWasUpdating = Application.ScreenUpdating
    calcMode = CLng(Application.Calculation)
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    weOpenedDest = False
    createdNew = False
    Set destWb = Nothing
    Set srcWb = Nothing
    Set srcWs = Nothing

    stepName = "resolve source workbook"
    Set srcWb = ResolveSourceWorkbook(SOURCE_SHEET_NAME)

    stepName = "find source sheet"
    If srcWb Is Nothing Then
        MsgBox "Version: " & MACRO_VERSION & vbCrLf & _
               "Лист """ & SOURCE_SHEET_NAME & """ не найден" & vbCrLf & _
               "ни в ThisWorkbook, ни в ActiveWorkbook." & vbCrLf & _
               "Откройте книгу с листом Ripper I и запустите снова.", _
               vbExclamation, "Export Ripper I"
        GoTo CleanExit
    End If

    Set srcWs = FindSheetByName(srcWb, SOURCE_SHEET_NAME)
    If srcWs Is Nothing Then
        sheetList = ListWorksheetNames(srcWb)
        MsgBox "Version: " & MACRO_VERSION & vbCrLf & _
               "Лист """ & SOURCE_SHEET_NAME & """ не найден в книге:" & vbCrLf & _
               srcWb.Name & vbCrLf & vbCrLf & _
               "Доступные листы:" & vbCrLf & sheetList, _
               vbExclamation, "Export Ripper I"
        GoTo CleanExit
    End If

    stepName = "scan source rows"
    lastSrcRow = LastNonEmptyRow(srcWs, COL_FIRST, COL_LAST)
    If lastSrcRow < 2 Then
        MsgBox "Version: " & MACRO_VERSION & vbCrLf & _
               "Нет данных на листе """ & SOURCE_SHEET_NAME & """.", _
               vbInformation, "Export Ripper I"
        GoTo CleanExit
    End If

    stepName = "open archive"
    Set destWb = OpenArchiveAppendOnly(DEST_PATH, createdNew, weOpenedDest)

    stepName = "ensure dest sheet"
    Set destWs = EnsureDestSheet(destWb, DEST_SHEET_NAME)

    stepName = "scan archive rows"
    lastDestRow = LastNonEmptyRow(destWs, 1, DEST_COL_COUNT)
    rowsBefore = lastDestRow

    If lastDestRow < 1 Then
        stepName = "copy header"
        CopyHeaderAsValues srcWs, destWs
        lastDestRow = 1
        rowsBefore = 1
    End If

    stepName = "build Batch Id map"
    Set existing = BuildBatchIdMap(destWs, lastDestRow)

    destRow = lastDestRow
    added = 0
    skippedEmpty = 0
    skippedDup = 0

    For r = 2 To lastSrcRow
        stepName = "source row " & CStr(r)
        packedVal = CellValueScalar(srcWs.Cells(r, COL_PACKED))

        If IsBlankCell(packedVal) Then
            skippedEmpty = skippedEmpty + 1
        Else
            batchKey = NormalizeBatchId(CellValueScalar(srcWs.Cells(r, COL_BATCH)))
            If Len(batchKey) = 0 Then
                skippedEmpty = skippedEmpty + 1
            ElseIf existing.Exists(batchKey) Then
                skippedDup = skippedDup + 1
            Else
                stepName = "append row " & CStr(r)
                destRow = NextEmptyArchiveRow(destWs, destRow)

                If destRow <= rowsBefore And rowsBefore > 1 Then
                    Err.Raise vbObjectError + 1, "ExportUniquePackedBatches", _
                              "Защита: попытка записи в занятую строку " & destRow & _
                              " (в архиве уже " & rowsBefore & " строк). Операция отменена."
                End If

                If Not IsBlankCell(CellValueScalar(destWs.Cells(destRow, 1))) Or _
                   Not IsBlankCell(CellValueScalar(destWs.Cells(destRow, DEST_BATCH_COL))) Then
                    Err.Raise vbObjectError + 2, "ExportUniquePackedBatches", _
                              "Защита: строка " & destRow & " в архиве не пустая. Операция отменена."
                End If

                CopyRowAsValues srcWs, r, destWs, destRow
                existing.Add batchKey, True
                added = added + 1
            End If
        End If
    Next r

    stepName = "verify row count"
    If LastNonEmptyRow(destWs, 1, DEST_COL_COUNT) < rowsBefore Then
        Err.Raise vbObjectError + 3, "ExportUniquePackedBatches", _
                  "Защита: число строк в архиве уменьшилось. Сохранение отменено."
    End If

    stepName = "save archive"
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
                If StrComp(CStr(destWb.FullName), CStr(srcWb.FullName), vbTextCompare) <> 0 Then
                    destWb.Close SaveChanges:=False
                End If
            Else
                destWb.Close SaveChanges:=False
            End If
        End If
    End If
    On Error GoTo 0
    Application.EnableEvents = True
    Application.Calculation = calcMode
    Application.ScreenUpdating = screenWasUpdating
    Exit Sub

CleanFail:
    errNum = Err.Number
    errDesc = Err.Description
    On Error Resume Next
    If weOpenedDest Then
        If Not destWb Is Nothing Then
            If Not srcWb Is Nothing Then
                If StrComp(CStr(destWb.FullName), CStr(srcWb.FullName), vbTextCompare) <> 0 Then
                    destWb.Close SaveChanges:=False
                End If
            Else
                destWb.Close SaveChanges:=False
            End If
        End If
    End If
    On Error GoTo 0
    Application.EnableEvents = True
    Application.Calculation = calcMode
    Application.ScreenUpdating = screenWasUpdating
    MsgBox "Version: " & MACRO_VERSION & vbCrLf & _
           "Шаг: " & stepName & vbCrLf & _
           "Ошибка VBA " & errNum & ": " & errDesc, vbCritical, "Export Ripper I"
End Sub

' ======================== Поиск листа (без For Each Worksheet) ========================

' ThisWorkbook, затем ActiveWorkbook - макрос может жить не в той же книге
Private Function ResolveSourceWorkbook(ByVal sheetName As String) As Workbook
    Dim wb As Workbook
    Dim ws As Worksheet

    Set ResolveSourceWorkbook = Nothing

    On Error Resume Next
    Set wb = ThisWorkbook
    On Error GoTo 0
    If Not wb Is Nothing Then
        Set ws = FindSheetByName(wb, sheetName)
        If Not ws Is Nothing Then
            Set ResolveSourceWorkbook = wb
            Exit Function
        End If
    End If

    Set wb = Nothing
    On Error Resume Next
    Set wb = Application.ActiveWorkbook
    On Error GoTo 0
    If Not wb Is Nothing Then
        Set ws = FindSheetByName(wb, sheetName)
        If Not ws Is Nothing Then
            Set ResolveSourceWorkbook = wb
            Exit Function
        End If
    End If

    ' Последний шанс: любая открытая книга с нужным листом
    On Error Resume Next
    For Each wb In Application.Workbooks
        Set ws = FindSheetByName(wb, sheetName)
        If Not ws Is Nothing Then
            Set ResolveSourceWorkbook = wb
            Exit Function
        End If
    Next wb
    On Error GoTo 0
End Function

' Прямой доступ по имени + обход по индексу через Object (без typed For Each)
Private Function FindSheetByName(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    Dim sh As Object
    Dim i As Long
    Dim target As String
    Dim nm As String

    Set FindSheetByName = Nothing
    If wb Is Nothing Then Exit Function

    target = Trim$(sheetName)
    If Len(target) = 0 Then Exit Function

    ' 1) Прямой доступ - Excel ищет без учёта регистра
    Set ws = Nothing
    On Error Resume Next
    Set ws = wb.Worksheets(target)
    On Error GoTo 0
    If Not ws Is Nothing Then
        Set FindSheetByName = ws
        Exit Function
    End If

    ' 2) Обход по индексу без For Each ws As Worksheet (давал Error 13)
    On Error Resume Next
    For i = 1 To wb.Worksheets.Count
        Set sh = Nothing
        Set sh = wb.Worksheets(i)
        If Not sh Is Nothing Then
            nm = vbNullString
            nm = CStr(sh.Name)
            If StrComp(Trim$(nm), target, vbTextCompare) = 0 Then
                Set ws = Nothing
                Set ws = sh
                If Not ws Is Nothing Then
                    Set FindSheetByName = ws
                    On Error GoTo 0
                    Exit Function
                End If
            End If
        End If
    Next i
    On Error GoTo 0
End Function

Private Function ListWorksheetNames(ByVal wb As Workbook) As String
    Dim i As Long
    Dim sh As Object
    Dim nm As String
    Dim acc As String

    acc = ""
    On Error Resume Next
    For i = 1 To wb.Worksheets.Count
        Set sh = Nothing
        nm = vbNullString
        Set sh = wb.Worksheets(i)
        nm = CStr(sh.Name)
        If Len(nm) > 0 Then
            If Len(acc) = 0 Then
                acc = "- " & nm
            Else
                acc = acc & vbCrLf & "- " & nm
            End If
        End If
    Next i
    On Error GoTo 0

    If Len(acc) = 0 Then acc = "(нет листов)"
    ListWorksheetNames = acc
End Function

' ======================== Архив ========================

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

    If FileExistsOnDisk(normalized, fso) Then
        Err.Raise vbObjectError + 10, "OpenArchiveAppendOnly", _
                  "Архив существует, но не удалось открыть:" & vbCrLf & normalized & vbCrLf & _
                  openDesc & vbCrLf & "Перезапись запрещена."
    End If

    EnsureParentFolder normalized, fso

    If FileExistsOnDisk(normalized, fso) Then
        Err.Raise vbObjectError + 11, "OpenArchiveAppendOnly", _
                  "Архив появился на диске. Перезапись запрещена:" & vbCrLf & normalized
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
              "Не удалось создать архив: " & normalized & vbCrLf & openDesc
End Function

Private Function FindOpenWorkbookByPath(ByVal path As String) As Workbook
    Dim wb As Workbook
    Dim fullName As String
    Dim target As String
    Dim targetName As String

    target = UCase$(NormalizePath(path))
    targetName = UCase$(FileNameOnly(path))
    Set FindOpenWorkbookByPath = Nothing

    On Error Resume Next
    For Each wb In Application.Workbooks
        fullName = vbNullString
        fullName = wb.FullName
        If Len(fullName) > 0 Then
            If UCase$(NormalizePath(fullName)) = target Then
                Set FindOpenWorkbookByPath = wb
                On Error GoTo 0
                Exit Function
            End If
            If UCase$(FileNameOnly(fullName)) = targetName Then
                If InStr(1, UCase$(NormalizePath(fullName)), "\AUTO\IRIPPER\", vbBinaryCompare) > 0 Then
                    Set FindOpenWorkbookByPath = wb
                    On Error GoTo 0
                    Exit Function
                End If
            End If
        End If
    Next wb
    On Error GoTo 0
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

Private Function EnsureDestSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    Dim firstWs As Worksheet

    Set ws = FindSheetByName(wb, sheetName)
    If ws Is Nothing Then
        Set firstWs = Nothing
        On Error Resume Next
        Set firstWs = wb.Worksheets(1)
        On Error GoTo 0

        If Not firstWs Is Nothing Then
            If wb.Worksheets.Count = 1 And LastNonEmptyRow(firstWs, 1, DEST_COL_COUNT) < 1 Then
                On Error Resume Next
                firstWs.Name = sheetName
                On Error GoTo 0
                Set ws = FindSheetByName(wb, sheetName)
            End If
        End If

        If ws Is Nothing Then
            Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
            On Error Resume Next
            ws.Name = sheetName
            On Error GoTo 0
        End If
    End If
    Set EnsureDestSheet = ws
End Function

' ======================== Строки / значения ========================

Private Function LastNonEmptyRow(ByVal ws As Worksheet, ByVal colFrom As Long, ByVal colTo As Long) As Long
    Dim rng As Range
    Dim area As Range
    Dim lastRow As Long
    Dim c As Long
    Dim found As Long

    lastRow = 0
    If ws Is Nothing Then
        LastNonEmptyRow = 0
        Exit Function
    End If

    On Error Resume Next
    Set rng = ws.Range(ws.Cells(1, colFrom), ws.Cells(ws.Rows.Count, colTo)).Find( _
                What:="*", _
                LookIn:=xlFormulas, _
                LookAt:=xlPart, _
                SearchOrder:=xlByRows, _
                SearchDirection:=xlPrevious, _
                MatchCase:=False)
    On Error GoTo 0

    If Not rng Is Nothing Then lastRow = rng.Row

    For c = colFrom To colTo
        found = ws.Cells(ws.Rows.Count, c).End(xlUp).Row
        If Not IsBlankCell(CellValueScalar(ws.Cells(found, c))) Then
            If found > lastRow Then lastRow = found
        End If
    Next c

    On Error Resume Next
    Set area = ws.UsedRange
    On Error GoTo 0
    If Not area Is Nothing Then
        found = area.Row + area.Rows.Count - 1
        If found > lastRow Then
            For c = colFrom To colTo
                If Not IsBlankCell(CellValueScalar(ws.Cells(found, c))) Then
                    lastRow = found
                    Exit For
                End If
            Next c
        End If
    End If

    LastNonEmptyRow = lastRow
End Function

Private Function NextEmptyArchiveRow(ByVal ws As Worksheet, ByVal afterRow As Long) As Long
    Dim r As Long
    r = afterRow + 1
    If r < 2 Then r = 2
    Do While Not IsBlankCell(CellValueScalar(ws.Cells(r, 1))) Or _
             Not IsBlankCell(CellValueScalar(ws.Cells(r, DEST_BATCH_COL)))
        r = r + 1
        If r >= ws.Rows.Count Then
            Err.Raise vbObjectError + 12, "NextEmptyArchiveRow", "Нет свободных строк в архиве."
        End If
    Loop
    NextEmptyArchiveRow = r
End Function

Private Sub CopyHeaderAsValues(ByVal srcWs As Worksheet, ByVal destWs As Worksheet)
    CopyRowAsValues srcWs, 1, destWs, 1
End Sub

Private Sub CopyRowAsValues(ByVal srcWs As Worksheet, ByVal srcRow As Long, _
                            ByVal destWs As Worksheet, ByVal destRow As Long)
    Dim c As Long
    Dim v As Variant
    For c = 0 To DEST_COL_COUNT - 1
        v = CellValueScalar(srcWs.Cells(srcRow, COL_FIRST + c))
        If IsError(v) Then
            destWs.Cells(destRow, 1 + c).Value2 = Empty
        Else
            On Error Resume Next
            destWs.Cells(destRow, 1 + c).Value2 = v
            If Err.Number <> 0 Then
                Err.Clear
                destWs.Cells(destRow, 1 + c).Value = CStr(v & "")
            End If
            On Error GoTo 0
        End If
    Next c
End Sub

Private Function BuildBatchIdMap(ByVal ws As Worksheet, ByVal lastRow As Long) As Object
    Dim dict As Object
    Dim r As Long
    Dim key As String

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1

    If lastRow >= 2 Then
        For r = 2 To lastRow
            key = NormalizeBatchId(CellValueScalar(ws.Cells(r, DEST_BATCH_COL)))
            If Len(key) > 0 Then
                If Not dict.Exists(key) Then dict.Add key, True
            End If
        Next r
    End If

    Set BuildBatchIdMap = dict
End Function

Private Function CellValueScalar(ByVal cell As Range) As Variant
    Dim v As Variant
    On Error GoTo FailSafe
    v = cell.Value2
    If IsArray(v) Then
        CellValueScalar = v(LBound(v, 1), LBound(v, 2))
    Else
        CellValueScalar = v
    End If
    Exit Function
FailSafe:
    On Error Resume Next
    CellValueScalar = cell.Text
    If Err.Number <> 0 Then CellValueScalar = Empty
    On Error GoTo 0
End Function

Private Function IsBlankCell(ByVal v As Variant) As Boolean
    Dim t As Integer
    On Error GoTo SafeBlank

    If IsArray(v) Then
        IsBlankCell = IsBlankCell(v(LBound(v)))
        Exit Function
    End If

    t = VarType(v)
    Select Case t
        Case vbEmpty, vbNull, vbError
            IsBlankCell = True
        Case vbString
            IsBlankCell = (Len(Trim$(v)) = 0)
        Case Else
            IsBlankCell = False
    End Select
    Exit Function

SafeBlank:
    IsBlankCell = True
End Function

Private Function NormalizeBatchId(ByVal v As Variant) As String
    Dim s As String
    Dim t As Integer
    On Error GoTo FailSafe

    If IsArray(v) Then
        NormalizeBatchId = NormalizeBatchId(v(LBound(v)))
        Exit Function
    End If

    t = VarType(v)
    Select Case t
        Case vbEmpty, vbNull, vbError
            NormalizeBatchId = ""
            Exit Function
        Case vbString
            s = Trim$(v)
        Case vbByte, vbInteger, vbLong
            s = CStr(v)
        Case vbSingle, vbDouble, vbCurrency, vbDecimal
            On Error Resume Next
            If Abs(CDbl(v) - Fix(CDbl(v))) < 0.000000001 Then
                s = Format$(v, "0")
            Else
                s = Trim$(Str$(v))
            End If
            If Err.Number <> 0 Or Len(s) = 0 Then s = Trim$(CStr(v & ""))
            On Error GoTo FailSafe
        Case vbBoolean
            If v Then s = "1" Else s = "0"
        Case vbDate
            s = Format$(v, "yyyymmdd")
        Case Else
            s = Trim$(CStr(v & ""))
    End Select

    NormalizeBatchId = Trim$(s)
    Exit Function

FailSafe:
    On Error Resume Next
    NormalizeBatchId = Trim$(CStr(v & ""))
    If Err.Number <> 0 Then NormalizeBatchId = ""
    On Error GoTo 0
End Function
