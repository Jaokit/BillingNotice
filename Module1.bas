Attribute VB_Name = "Module1"
Option Explicit

Public Const P1_ROOM As String = "E2"
Public Const P1_DATE As String = "E3"
Public Const P1_WU   As String = "C5"
Public Const P1_WAMT As String = "E5"
Public Const P1_EU   As String = "C6"
Public Const P1_EAMT As String = "E6"
Public Const P1_GARB As String = "E7"
Public Const P1_RFIN As String = "C8"
Public Const P1_RFAM As String = "E8"
Public Const P1_FINEIN  As String = "C9"
Public Const P1_FINEAMT As String = "E9"
Public Const P1_TOT  As String = "E11"
Public Const P1_OWNER As String = "B3"

Public Const P2_ROOM As String = "E13"
Public Const P2_DATE As String = "E14"
Public Const P2_WU   As String = "C16"
Public Const P2_WAMT As String = "E16"
Public Const P2_EU   As String = "C17"
Public Const P2_EAMT As String = "E17"
Public Const P2_GARB As String = "E18"
Public Const P2_RFIN As String = "C19"
Public Const P2_RFAM As String = "E19"
Public Const P2_FINEIN  As String = "C20"
Public Const P2_FINEAMT As String = "E20"
Public Const P2_TOT  As String = "E22"
Public Const P2_OWNER As String = "B14"

Public Const P3_ROOM As String = "E24"
Public Const P3_DATE As String = "E25"
Public Const P3_WU   As String = "C27"
Public Const P3_WAMT As String = "E27"
Public Const P3_EU   As String = "C28"
Public Const P3_EAMT As String = "E28"
Public Const P3_GARB As String = "E29"
Public Const P3_RFIN As String = "C30"
Public Const P3_RFAM As String = "E30"
Public Const P3_FINEIN  As String = "C31"
Public Const P3_FINEAMT As String = "E31"
Public Const P3_TOT  As String = "E33"
Public Const P3_OWNER As String = "B25"

Public Const WATER_RATE  As Double = 28
Public Const ELEC_RATE   As Double = 10
Public Const GARBAGE_FEE As Double = 20

Public Const P1_WPREV As String = "G5"
Public Const P1_WCURR As String = "H5"
Public Const P1_EPREV As String = "G6"
Public Const P1_ECURR As String = "H6"

Public Const P2_WPREV As String = "G16"
Public Const P2_WCURR As String = "H16"
Public Const P2_EPREV As String = "G17"
Public Const P2_ECURR As String = "H17"

Public Const P3_WPREV As String = "G27"
Public Const P3_WCURR As String = "H27"
Public Const P3_EPREV As String = "G28"
Public Const P3_ECURR As String = "H28"

Private Function BahtFmt() As String
    BahtFmt = ChrW(3647) & "#,##0"
End Function

Private Sub ClearEach(ws As Worksheet, ParamArray addrs())
    Dim i As Long
    For i = LBound(addrs) To UBound(addrs)
        If Len(addrs(i) & vbNullString) > 0 Then
            ws.Range(CStr(addrs(i))).MergeArea.ClearContents
        End If
    Next i
End Sub

Private Sub SetValueMerged(ws As Worksheet, ByVal addr As String, ByVal v As Variant)
    ws.Range(addr).MergeArea.Value = v
End Sub

Public Sub ForceBillDateFormat(Optional ByVal ws As Worksheet)
    If ws Is Nothing Then Set ws = ThisWorkbook.Worksheets("Bill")
    On Error Resume Next
    ws.Range(P1_DATE & "," & P2_DATE & "," & P3_DATE).NumberFormat = "mm/yyyy"
    On Error GoTo 0
End Sub

Public Sub AutoFillCurrentMonth(Optional ByVal ws As Worksheet)
    If ws Is Nothing Then Set ws = ThisWorkbook.Worksheets("Bill")
    Dim curMonth As Date: curMonth = DateSerial(Year(Date), Month(Date), 1)

    Application.EnableEvents = False
    On Error GoTo Done

    Dim c As Range
    For Each c In ws.Range(P1_DATE & "," & P2_DATE & "," & P3_DATE).Cells
        If Len(Trim$(c.Value & "")) = 0 Or Not IsDate(c.Value) Then
            c.Value = curMonth
        End If
        c.NumberFormat = "mm/yyyy"
    Next c

Done:
    Application.EnableEvents = True
End Sub

Public Function OwnerFromNameSheet(ByVal roomValue As Variant) As String
    Dim ws As Worksheet, f As Range, key As String, v As Variant
    On Error GoTo EH

    Set ws = ThisWorkbook.Worksheets("Name")
    key = Trim$(CStr(roomValue))

    Set f = ws.Columns("A").Find(What:=key, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    If Not f Is Nothing Then
        OwnerFromNameSheet = Trim$(CStr(f.Offset(0, 1).Value))
        Exit Function
    End If

    v = Application.VLookup(key, ws.Range("A:B"), 2, False)
    If IsError(v) Then
        OwnerFromNameSheet = ""
    Else
        OwnerFromNameSheet = Trim$(CStr(v))
    End If
    Exit Function
EH:
    OwnerFromNameSheet = ""
End Function

Public Sub RefreshOwnerNames(Optional ByVal ws As Worksheet)
    If ws Is Nothing Then Set ws = ThisWorkbook.Worksheets("Bill")
    Application.EnableEvents = False
    On Error GoTo Done

    Dim v1 As String, v2 As String, v3 As String
    v1 = IIf(Len(Trim$(ws.Range(P1_ROOM).Value)) = 0, "", OwnerFromNameSheet(ws.Range(P1_ROOM).Value))
    v2 = IIf(Len(Trim$(ws.Range(P2_ROOM).Value)) = 0, "", OwnerFromNameSheet(ws.Range(P2_ROOM).Value))
    v3 = IIf(Len(Trim$(ws.Range(P3_ROOM).Value)) = 0, "", OwnerFromNameSheet(ws.Range(P3_ROOM).Value))

    SetValueMerged ws, P1_OWNER, v1
    SetValueMerged ws, P2_OWNER, v2
    SetValueMerged ws, P3_OWNER, v3

Done:
    Application.EnableEvents = True
End Sub

Private Function RoomRate(ByVal room As String, ByRef isManual As Boolean) As Double
    Dim letter As String, num As Long, i As Long
    room = UCase$(Trim$(room))
    If room = "" Then isManual = True: RoomRate = 0: Exit Function

    letter = Left$(room, 1)
    For i = 2 To Len(room)
        If Mid$(room, i, 1) Like "[0-9]" Then
            num = num * 10 + CLng(Mid$(room, i, 1))
        Else
            Exit For
        End If
    Next i

    isManual = False
    Select Case letter
        Case "A"
            If num >= 1 And num <= 12 Then
                isManual = True: RoomRate = 0
            ElseIf num <= 24 Then
                RoomRate = 1400
            Else
                isManual = True: RoomRate = 0
            End If
        Case "B"
            If num >= 1 And num <= 12 Then
                RoomRate = 1600
            ElseIf num <= 24 Then
                RoomRate = 1400
            Else
                isManual = True: RoomRate = 0
            End If
        Case Else
            isManual = True: RoomRate = 0
    End Select
End Function

Private Sub ComputeUsage(ByVal ws As Worksheet, ByVal prevAddr As String, _
                         ByVal currAddr As String, ByVal outUsageAddr As String)
    Dim p As Variant, c As Variant
    p = ws.Range(prevAddr).Value
    c = ws.Range(currAddr).Value

    On Error Resume Next
    ws.Range(prevAddr).Interior.ColorIndex = xlNone
    ws.Range(currAddr).Interior.ColorIndex = xlNone
    On Error GoTo 0

    If IsNumeric(p) And IsNumeric(c) Then
        If Val(c) >= Val(p) Then
            ws.Range(outUsageAddr).Value = Val(c) - Val(p)
        Else
            ws.Range(outUsageAddr).ClearContents
            ws.Range(prevAddr).Interior.Color = RGB(255, 220, 220)
            ws.Range(currAddr).Interior.Color = RGB(255, 220, 220)
        End If
    End If
End Sub

Public Sub UpdateUsageFromReadings(Optional ByVal ws As Worksheet = Nothing)
    If ws Is Nothing Then Set ws = ThisWorkbook.Worksheets("Bill")
    ' ???
    ComputeUsage ws, P1_WPREV, P1_WCURR, P1_WU
    ComputeUsage ws, P2_WPREV, P2_WCURR, P2_WU
    ComputeUsage ws, P3_WPREV, P3_WCURR, P3_WU
    ' ??
    ComputeUsage ws, P1_EPREV, P1_ECURR, P1_EU
    ComputeUsage ws, P2_EPREV, P2_ECURR, P2_EU
    ComputeUsage ws, P3_EPREV, P3_ECURR, P3_EU
End Sub

Private Sub CalcPanel(roomCell As String, dateCell As String, _
                      wuCell As String, wAmtCell As String, _
                      euCell As String, eAmtCell As String, _
                      garbCell As String, _
                      roomFeeInCell As String, roomFeeAmtCell As String, _
                      fineInCell As String, fineAmtCell As String, _
                      totalCell As String, panelName As String, _
                      Optional ws As Worksheet = Nothing)

    If ws Is Nothing Then Set ws = ThisWorkbook.Worksheets("Bill")

    Dim room As String, wu As Double, eu As Double, fine As Double
    Dim isManual As Boolean, rr As Double, rf As Double, total As Double

    room = Trim$(ws.Range(roomCell).Value)
    wu = Val(ws.Range(wuCell).Value)
    eu = Val(ws.Range(euCell).Value)
    fine = Val(ws.Range(fineInCell).Value)

    ws.Range(wAmtCell).Value = wu * WATER_RATE
    ws.Range(eAmtCell).Value = eu * ELEC_RATE
    ws.Range(garbCell).Value = GARBAGE_FEE

    rr = RoomRate(room, isManual)
    If isManual Then
        If Trim$(CStr(ws.Range(roomFeeInCell).Value)) = "" And room <> "" Then
            Dim v As Variant
            v = Application.InputBox( _
                "Room " & room & " is in A1–A12 (shop range)." & vbCrLf & _
                "Please enter the room fee for " & panelName & ":", _
                "Room fee", Type:=1)
            If v <> False Then ws.Range(roomFeeInCell).Value = Val(v)
        End If
        rf = Val(ws.Range(roomFeeInCell).Value)
    Else
        rf = rr
        If room <> "" Then ws.Range(roomFeeInCell).Value = rf
    End If
    ws.Range(roomFeeAmtCell).Value = rf

    ws.Range(fineAmtCell).Value = fine

    total = Val(ws.Range(wAmtCell).Value) + _
            Val(ws.Range(eAmtCell).Value) + _
            Val(ws.Range(garbCell).Value) + _
            Val(ws.Range(roomFeeAmtCell).Value) + _
            Val(ws.Range(fineAmtCell).Value)
    ws.Range(totalCell).Value = total

    ws.Range(wAmtCell & "," & eAmtCell & "," & garbCell & "," & _
             roomFeeAmtCell & "," & fineAmtCell & "," & totalCell).NumberFormat = BahtFmt
End Sub

Public Sub CalcP1(): CalcPanel P1_ROOM, P1_DATE, P1_WU, P1_WAMT, P1_EU, P1_EAMT, P1_GARB, P1_RFIN, P1_RFAM, P1_FINEIN, P1_FINEAMT, P1_TOT, "Panel 1": End Sub
Public Sub CalcP2(): CalcPanel P2_ROOM, P2_DATE, P2_WU, P2_WAMT, P2_EU, P2_EAMT, P2_GARB, P2_RFIN, P2_RFAM, P2_FINEIN, P2_FINEAMT, P2_TOT, "Panel 2": End Sub
Public Sub CalcP3(): CalcPanel P3_ROOM, P3_DATE, P3_WU, P3_WAMT, P3_EU, P3_EAMT, P3_GARB, P3_RFIN, P3_RFAM, P3_FINEIN, P3_FINEAMT, P3_TOT, "Panel 3": End Sub
Public Sub CalcAll():  CalcP1: CalcP2: CalcP3: End Sub

Private Sub SaveOneRow(wsH As Worksheet, d As Variant, room As String, _
                       wu As Double, wAmt As Double, eu As Double, eAmt As Double, _
                       garb As Double, roomFee As Double, fine As Double, total As Double)

    Dim r As Long: r = wsH.Cells(wsH.Rows.Count, "A").End(xlUp).Row
    If r < 2 Then r = 2 Else r = r + 1

    wsH.Cells(r, 1).Value = IIf(IsDate(d), CDate(d), Date)
    wsH.Cells(r, 2).Value = room
    wsH.Cells(r, 3).Value = wu
    wsH.Cells(r, 4).Value = wAmt
    wsH.Cells(r, 5).Value = eu
    wsH.Cells(r, 6).Value = eAmt
    wsH.Cells(r, 7).Value = garb
    wsH.Cells(r, 8).Value = roomFee
    wsH.Cells(r, 9).Value = fine
    wsH.Cells(r, 10).Value = total

    wsH.Range(wsH.Cells(r, 4), wsH.Cells(r, 10)).NumberFormat = BahtFmt
End Sub

Public Sub SaveAllPanelsToHistorAndPrint()
    Dim wsB As Worksheet, wsH As Worksheet
    Set wsB = ThisWorkbook.Worksheets("Bill")

    AutoFillCurrentMonth wsB
    ForceBillDateFormat wsB
    UpdateUsageFromReadings wsB
    CalcAll

    On Error Resume Next
    Set wsH = ThisWorkbook.Worksheets("Histor")
    On Error GoTo 0
    If wsH Is Nothing Then
        Set wsH = ThisWorkbook.Worksheets.Add(After:=wsB)
        wsH.Name = "Histor"
    End If

    If Trim$(wsB.Range(P1_ROOM).Value) <> "" Then SaveOneRow wsH, wsB.Range(P1_DATE).Value, wsB.Range(P1_ROOM).Value, _
        Val(wsB.Range(P1_WU).Value), Val(wsB.Range(P1_WAMT).Value), Val(wsB.Range(P1_EU).Value), Val(wsB.Range(P1_EAMT).Value), _
        Val(wsB.Range(P1_GARB).Value), Val(wsB.Range(P1_RFAM).Value), Val(wsB.Range(P1_FINEAMT).Value), Val(wsB.Range(P1_TOT).Value)

    If Trim$(wsB.Range(P2_ROOM).Value) <> "" Then SaveOneRow wsH, wsB.Range(P2_DATE).Value, wsB.Range(P2_ROOM).Value, _
        Val(wsB.Range(P2_WU).Value), Val(wsB.Range(P2_WAMT).Value), Val(wsB.Range(P2_EU).Value), Val(wsB.Range(P2_EAMT).Value), _
        Val(wsB.Range(P2_GARB).Value), Val(wsB.Range(P2_RFAM).Value), Val(wsB.Range(P2_FINEAMT).Value), Val(wsB.Range(P2_TOT).Value)

    If Trim$(wsB.Range(P3_ROOM).Value) <> "" Then SaveOneRow wsH, wsB.Range(P3_DATE).Value, wsB.Range(P3_ROOM).Value, _
        Val(wsB.Range(P3_WU).Value), Val(wsB.Range(P3_WAMT).Value), Val(wsB.Range(P3_EU).Value), Val(wsB.Range(P3_EAMT).Value), _
        Val(wsB.Range(P3_GARB).Value), Val(wsB.Range(P3_RFAM).Value), Val(wsB.Range(P3_FINEAMT).Value), Val(wsB.Range(P3_TOT).Value)

    wsB.PrintOut
    ClearAllPanels wsB
End Sub

Public Sub ClearAllPanels(Optional ws As Worksheet = Nothing)
    If ws Is Nothing Then Set ws = ThisWorkbook.Worksheets("Bill")
    Application.EnableEvents = False

    ClearEach ws, P1_DATE, P1_ROOM, P1_WU, P1_EU, P1_RFIN, P1_FINEIN, P1_OWNER
    ws.Range(P1_WAMT & "," & P1_EAMT & "," & P1_GARB & "," & P1_RFAM & "," & P1_FINEAMT & "," & P1_TOT).ClearContents

    ClearEach ws, P2_DATE, P2_ROOM, P2_WU, P2_EU, P2_RFIN, P2_FINEIN, P2_OWNER
    ws.Range(P2_WAMT & "," & P2_EAMT & "," & P2_GARB & "," & P2_RFAM & "," & P2_FINEAMT & "," & P2_TOT).ClearContents

    ClearEach ws, P3_DATE, P3_ROOM, P3_WU, P3_EU, P3_RFIN, P3_FINEIN, P3_OWNER
    ws.Range(P3_WAMT & "," & P3_EAMT & "," & P3_GARB & "," & P3_RFAM & "," & P3_FINEAMT & "," & P3_TOT).ClearContents

    ws.Range(P1_WPREV & "," & P1_WCURR & "," & P1_EPREV & "," & P1_ECURR & "," & _
             P2_WPREV & "," & P2_WCURR & "," & P2_EPREV & "," & P2_ECURR & "," & _
             P3_WPREV & "," & P3_WCURR & "," & P3_EPREV & "," & P3_ECURR).ClearContents
    On Error Resume Next
    ws.Range(P1_WPREV & "," & P1_WCURR & "," & P1_EPREV & "," & P1_ECURR & "," & _
             P2_WPREV & "," & P2_WCURR & "," & P2_EPREV & "," & P2_ECURR & "," & _
             P3_WPREV & "," & P3_WCURR & "," & P3_EPREV & "," & P3_ECURR).Interior.ColorIndex = xlNone
    On Error GoTo 0

    Application.EnableEvents = True
    AutoFillCurrentMonth ws
End Sub

Public Sub ResetAppState()
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Events ON & AutoCalc ON", vbInformation
End Sub


