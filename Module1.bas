Attribute VB_Name = "Module1"
Option Explicit

'==================== CELL MAP (place these FIRST) ====================
' Panel 1 (Top)
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

' Panel 2 (Middle)
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

' Panel 3 (Bottom)
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
'=====================================================================

'==================== RATES & FORMAT ====================
Public Const WATER_RATE  As Double = 28
Public Const ELEC_RATE   As Double = 10
Public Const GARBAGE_FEE As Double = 20

Private Function BahtFmt() As String
    BahtFmt = ChrW(3647) & "#,##0"
End Function
'=======================================================

'==================== ROOM RULE =========================
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
'=======================================================

'==================== CORE CALC =========================
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
'=======================================================

'==================== SAVE + PRINT + CLEAR ===============
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
    ClearAllPanels
End Sub

Public Sub ClearAllPanels(Optional ws As Worksheet = Nothing)
    If ws Is Nothing Then Set ws = ThisWorkbook.Worksheets("Bill")
    Application.EnableEvents = False
    ' P1
    ws.Range(P1_DATE & "," & P1_ROOM & "," & P1_WU & "," & P1_EU & "," & P1_RFIN & "," & P1_FINEIN).ClearContents
    ws.Range(P1_WAMT & "," & P1_EAMT & "," & P1_GARB & "," & P1_RFAM & "," & P1_FINEAMT & "," & P1_TOT).ClearContents
    ' P2
    ws.Range(P2_DATE & "," & P2_ROOM & "," & P2_WU & "," & P2_EU & "," & P2_RFIN & "," & P2_FINEIN).ClearContents
    ws.Range(P2_WAMT & "," & P2_EAMT & "," & P2_GARB & "," & P2_RFAM & "," & P2_FINEAMT & "," & P2_TOT).ClearContents
    ' P3
    ws.Range(P3_DATE & "," & P3_ROOM & "," & P3_WU & "," & P3_EU & "," & P3_RFIN & "," & P3_FINEIN).ClearContents
    ws.Range(P3_WAMT & "," & P3_EAMT & "," & P3_GARB & "," & P3_RFAM & "," & P3_FINEAMT & "," & P3_TOT).ClearContents
    Application.EnableEvents = True
End Sub

' Utilities
Public Sub ResetAppState()
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Events ON & AutoCalc ON", vbInformation
End Sub


