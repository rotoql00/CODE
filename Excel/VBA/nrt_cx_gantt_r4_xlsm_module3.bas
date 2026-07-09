Option Explicit

' ★ 핵심 보완: 메인 파일명을 전달받을 수 있도록 HostName 인자 배치
Public Sub Cx_Gantt_SungJun_Choi(Optional ByVal HostName As String = "")

    Dim wbTarget As Workbook
    Dim wsData As Worksheet
    Dim wsCfg As Worksheet

    Dim rngInput As Range
    Dim rngMode As Range
    Dim rngHdr As Range
    Dim rngOut As Range

    Dim vInput As Variant
    Dim vMode As Variant
    Dim vHdr As Variant
    Dim outArr As Variant

    Dim colDates() As Double
    Dim nRows As Long
    Dim nCols As Long

    Dim r As Long, c As Long
    Dim calN As Double
    Dim headerLastDate As Double

    Dim modeKey As String
    Dim defaultMode As String
    Dim joinDelim As String
    Dim openEndedPolicy As String

    Dim headerStartCellAddr As String
    Dim inputRangeAddr As String
    Dim modeRangeAddr As String
    Dim outputStartCellAddr As String

    Dim lastHdrCol As Long

    Dim prevCalc As XlCalculation
    Dim prevScreen As Boolean
    Dim prevEvents As Boolean
    Dim prevStatus As Variant

    Dim cache As Object
    Set cache = CreateObject("Scripting.Dictionary")

    On Error GoTo EH

    '-------------------------------------------------------------------------
    ' 🔥 [근본 해결]: 매크로 위치 필터가 "모든 통합 문서"일 때 발생하는 포커스 유실 차단
    '-------------------------------------------------------------------------
    If HostName <> "" Then
        ' 메인 파일에서 넘겨준 이름으로 원본 워크북 객체를 명확하게 고정합니다.
        Set wbTarget = Workbooks(HostName)
        ' 임시 파일의 빈 시트가 아니라, 원본 파일의 현재 활성화된 간트차트 시트를 정조준합니다.
        Set wsData = wbTarget.ActiveSheet
    Else
        ' 단독 실행 등 예외 상황 방어용
        Set wbTarget = ThisWorkbook
        Set wsData = ActiveSheet
    End If
    
    ' Config 시트 또한 임시 새 문서가 아닌 원본 파일 내부에서 정확하게 가져옵니다.
    Set wsCfg = wbTarget.Worksheets("Config")
    '-------------------------------------------------------------------------

    defaultMode = SafeTrimUpper(GetAppConfigValue(wsCfg, "DefaultMode", "E"))
    joinDelim = GetAppConfigValue(wsCfg, "JoinDelimiter", "/")
    openEndedPolicy = SafeTrimUpper(GetAppConfigValue(wsCfg, "OpenEndedPolicy", "ONE_DAY"))

    headerStartCellAddr = GetAppConfigValue(wsCfg, "HeaderStartCell", "O3")
    inputRangeAddr = GetAppConfigValue(wsCfg, "InputRange", "D5:M350")
    modeRangeAddr = GetAppConfigValue(wsCfg, "ModeRange", "B5:B350")
    outputStartCellAddr = GetAppConfigValue(wsCfg, "OutputStartCell", "O5")

    Set rngInput = wsData.Range(inputRangeAddr)
    Set rngMode = wsData.Range(modeRangeAddr)

    nRows = rngInput.Rows.Count

    lastHdrCol = FindHeaderLastCol(wsData, headerStartCellAddr)

    Set rngHdr = wsData.Range(wsData.Range(headerStartCellAddr), wsData.Cells(wsData.Range(headerStartCellAddr).Row, lastHdrCol))

    nCols = rngHdr.Columns.Count

    Set rngOut = wsData.Range(outputStartCellAddr).Resize(nRows, nCols)

    vInput = rngInput.Value2
    vMode = rngMode.Value2
    vHdr = rngHdr.Value2

    ReDim colDates(1 To nCols)

    For c = 1 To nCols
        colDates(c) = ToDateSerialFast(vHdr(1, c))
    Next c

    headerLastDate = 0
    If nCols >= 1 Then headerLastDate = colDates(nCols)

    ReDim outArr(1 To nRows, 1 To nCols)

    prevCalc = Application.Calculation
    prevScreen = Application.ScreenUpdating
    prevEvents = Application.EnableEvents
    prevStatus = Application.StatusBar

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Config-driven Stage grid 계산 중..."

    For r = 1 To nRows

        modeKey = SafeTrimUpper(vMode(r, 1))
        If Len(modeKey) = 0 Then modeKey = defaultMode

        Dim stageCount As Long
        Dim stageName() As String
        Dim relStartIdx() As Long
        Dim relFinishIdx() As Long

        LoadStagesForModeCached wsCfg, rngInput, modeKey, cache, stageCount, stageName, relStartIdx, relFinishIdx

        If stageCount = 0 Then
            For c = 1 To nCols
                outArr(r, c) = ""
            Next c
            GoTo NEXT_ROW
        End If

        Dim sArr() As Double
        Dim fArr() As Double
        ReDim sArr(1 To stageCount)
        ReDim fArr(1 To stageCount)

        Dim i As Long
        Dim s As Double, f As Double

        For i = 1 To stageCount

            s = 0
            f = 0

            If relStartIdx(i) > 0 And relStartIdx(i) <= UBound(vInput, 2) Then
                s = ToDateSerialFast(vInput(r, relStartIdx(i)))
            End If

            If relFinishIdx(i) > 0 And relFinishIdx(i) <= UBound(vInput, 2) Then
                f = ToDateSerialFast(vInput(r, relFinishIdx(i)))
            End If

            NormalizeRangeWithPolicy s, f, openEndedPolicy, headerLastDate

            sArr(i) = s
            fArr(i) = f

        Next i

        For c = 1 To nCols

            calN = colDates(c)

            If calN = 0 Then
                outArr(r, c) = ""
            Else
                outArr(r, c) = BuildStageString(calN, stageCount, stageName, sArr, fArr, joinDelim)
            End If

        Next c

NEXT_ROW:
        If (r Mod 20) = 0 Then
            Application.StatusBar = "Config-driven Stage grid 계산 중... (" & r & "/" & nRows & ")"
        End If

    Next r

    rngOut.Value2 = outArr

    Application.StatusBar = False
    Application.Calculation = prevCalc
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreen

    MsgBox "WEB code Progress completion.", vbInformation, "완료"

    Exit Sub

EH:
    Application.StatusBar = False
    Application.Calculation = prevCalc
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreen
    MsgBox "오류: " & Err.Description, vbExclamation, "Cx_Gantt_ConfigDriven_Run"

End Sub


Private Function BuildStageString( _
    ByVal calN As Double, _
    ByVal stageCount As Long, _
    ByRef stageName() As String, _
    ByRef sArr() As Double, _
    ByRef fArr() As Double, _
    ByVal delimiter As String _
) As String

    Dim res As String: res = ""
    Dim i As Long

    For i = 1 To stageCount
        If IsInRange(calN, sArr(i), fArr(i)) Then
            res = AppendStage(res, stageName(i), delimiter)
        End If
    Next i

    BuildStageString = res

End Function


Private Sub LoadStagesForModeCached( _
    ByVal wsCfg As Worksheet, _
    ByVal rngInput As Range, _
    ByVal modeKey As String, _
    ByVal cache As Object, _
    ByRef stageCount As Long, _
    ByRef stageName() As String, _
    ByRef relStartIdx() As Long, _
    ByRef relFinishIdx() As Long _
)

    Dim pack As Variant

    If cache.Exists(modeKey) Then
        pack = cache(modeKey)
        stageCount = pack(0)
        stageName = pack(1)
        relStartIdx = pack(2)
        relFinishIdx = pack(3)
        Exit Sub
    End If

    LoadStagesFromConfig wsCfg, rngInput, modeKey, stageCount, stageName, relStartIdx, relFinishIdx
    pack = Array(stageCount, stageName, relStartIdx, relFinishIdx)
    cache(modeKey) = pack

End Sub


Private Sub LoadStagesFromConfig( _
    ByVal wsCfg As Worksheet, _
    ByVal rngInput As Range, _
    ByVal modeKey As String, _
    ByRef stageCount As Long, _
    ByRef stageName() As String, _
    ByRef relStartIdx() As Long, _
    ByRef relFinishIdx() As Long _
)

    Dim lo As ListObject
    Set lo = wsCfg.ListObjects("StageConfig")

    Dim colMode As Long, colSeq As Long, colName As Long
    Dim colS As Long, colF As Long

    colMode = lo.ListColumns("Mode").Index
    colSeq = lo.ListColumns("Seq").Index
    colName = lo.ListColumns("StageName").Index
    colS = lo.ListColumns("StartCol").Index
    colF = lo.ListColumns("FinishCol").Index

    Dim r As Long
    Dim matchCount As Long: matchCount = 0

    For r = 1 To lo.DataBodyRange.Rows.Count
        If SafeTrimUpper(lo.DataBodyRange.Cells(r, colMode).Value2) = SafeTrimUpper(modeKey) Then
            matchCount = matchCount + 1
        End If
    Next r

    stageCount = matchCount
    If stageCount = 0 Then Exit Sub

    ReDim stageName(1 To stageCount)
    ReDim relStartIdx(1 To stageCount)
    ReDim relFinishIdx(1 To stageCount)

    Dim seqArr() As Long: ReDim seqArr(1 To stageCount)

    Dim i As Long: i = 0
    Dim absS As Long, absF As Long
    Dim relS As Long, relF As Long

    For r = 1 To lo.DataBodyRange.Rows.Count

        If SafeTrimUpper(lo.DataBodyRange.Cells(r, colMode).Value2) = SafeTrimUpper(modeKey) Then

            i = i + 1

            seqArr(i) = CLng(lo.DataBodyRange.Cells(r, colSeq).Value2)
            stageName(i) = CStr(lo.DataBodyRange.Cells(r, colName).Value2)

            absS = ColLetterToIndex(CStr(lo.DataBodyRange.Cells(r, colS).Value2))
            absF = ColLetterToIndex(CStr(lo.DataBodyRange.Cells(r, colF).Value2))

            relS = absS - rngInput.Column + 1
            relF = absF - rngInput.Column + 1

            relStartIdx(i) = relS
            relFinishIdx(i) = relF

        End If

    Next r

    Dim a As Long, b As Long
    Dim tmpL As Long, tmpS As String

    For a = 1 To stageCount - 1
        For b = a + 1 To stageCount
            If seqArr(b) < seqArr(a) Then

                tmpL = seqArr(a): seqArr(a) = seqArr(b): seqArr(b) = tmpL
                tmpS = stageName(a): stageName(a) = stageName(b): stageName(b) = tmpS
                tmpL = relStartIdx(a): relStartIdx(a) = relStartIdx(b): relStartIdx(b) = tmpL
                tmpL = relFinishIdx(a): relFinishIdx(a) = relFinishIdx(b): relFinishIdx(b) = tmpL

            End If
        Next b
    Next a

End Sub


Private Function FindHeaderLastCol(ByVal ws As Worksheet, ByVal headerStartCellAddr As String) As Long

    Dim firstCell As Range
    Dim curCell As Range
    Dim lastCol As Long
    Dim v As Variant

    Set firstCell = ws.Range(headerStartCellAddr)
    Set curCell = firstCell
    lastCol = firstCell.Column

    Do While True

        v = curCell.Value2

        If ToDateSerialFast(v) = 0 Then Exit Do

        lastCol = curCell.Column
        Set curCell = curCell.Offset(0, 1)

    Loop

    FindHeaderLastCol = lastCol

End Function


Private Function AppendStage(ByVal res As String, ByVal stageName As String, ByVal delimiter As String) As String

    If Len(res) = 0 Then
        AppendStage = stageName
        Exit Function
    End If

    AppendStage = res & delimiter & stageName

End Function


Private Sub NormalizeRangeWithPolicy(ByRef s As Double, ByRef f As Double, ByVal policy As String, ByVal headerLastDate As Double)

    If s <> 0 And f = 0 Then

        If policy = "TO_END" Then
            f = headerLastDate
        ElseIf policy = "IGNORE" Then
            s = 0
            f = 0
        Else
            f = s
        End If

        Exit Sub
    End If

    If s = 0 And f <> 0 Then

        If policy = "IGNORE" Then
            s = 0
            f = 0
        Else
            s = f
        End If

        Exit Sub
    End If

    If s <> 0 And f <> 0 Then
        If f < s Then
            Dim tmp As Double
            tmp = s
            s = f
            f = tmp
        End If
    End If

End Sub


Private Function IsInRange(ByVal calN As Double, ByVal s As Double, ByVal f As Double) As Boolean

    If s = 0 And f = 0 Then
        IsInRange = False
        Exit Function
    End If

    IsInRange = (calN >= s And calN <= f)

End Function


Private Function ToDateSerialFast(ByVal v As Variant) As Double

    If IsError(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    If v = "" Then Exit Function

    If IsNumeric(v) Then
        ToDateSerialFast = CDbl(v)
        Exit Function
    End If

    If IsDate(v) Then
        ToDateSerialFast = CDbl(CDate(v))
        Exit Function
    End If

End Function


Private Function SafeTrimUpper(ByVal v As Variant) As String

    If IsError(v) Then
        SafeTrimUpper = ""
        Exit Function
    End If

    If IsEmpty(v) Then
        SafeTrimUpper = ""
        Exit Function
    End If

    SafeTrimUpper = UCase$(Trim$(CStr(v)))

End Function


Private Function ColLetterToIndex(ByVal colLetter As String) As Long

    Dim i As Long
    Dim ch As String
    Dim res As Long

    colLetter = UCase$(Trim$(colLetter))

    If Len(colLetter) = 0 Then
        ColLetterToIndex = 0
        Exit Function
    End If

    res = 0

    For i = 1 To Len(colLetter)
        ch = Mid$(colLetter, i, 1)

        If ch < "A" Or ch > "Z" Then
            ColLetterToIndex = 0
            Exit Function
        End If

        res = res * 26 + (Asc(ch) - Asc("A") + 1)
    Next i

    ColLetterToIndex = res

End Function


Private Function GetAppConfigValue(ByVal wsConfig As Worksheet, ByVal keyName As String, ByVal defaultValue As String) As String

    Dim lo As ListObject
    Dim keyCol As Long, valCol As Long
    Dim r As Long
    Dim v As Variant

    keyName = SafeTrimUpper(keyName)

    On Error GoTo EH

    Set lo = wsConfig.ListObjects("AppConfig")

    keyCol = lo.ListColumns("Key").Index
    valCol = lo.ListColumns("Value").Index

    For r = 1 To lo.DataBodyRange.Rows.Count
        v = lo.DataBodyRange.Cells(r, keyCol).Value2
        If SafeTrimUpper(v) = keyName Then
            GetAppConfigValue = CStr(lo.DataBodyRange.Cells(r, valCol).Value2)
            Exit Function
        End If
    Next r

    GetAppConfigValue = defaultValue
    Exit Function

EH:
    GetAppConfigValue = defaultValue

End Function
