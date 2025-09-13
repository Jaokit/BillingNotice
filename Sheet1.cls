VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Activate()
    ForceBillDateFormat Me
    AutoFillCurrentMonth Me
    RefreshOwnerNames Me
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    If Application.EnableEvents = False Then Exit Sub
    Application.EnableEvents = False
    On Error GoTo Done

    Dim hitRooms As Range, c As Range
    Set hitRooms = Intersect(Target, Me.Range(P1_ROOM & "," & P2_ROOM & "," & P3_ROOM))
    If Not hitRooms Is Nothing Then
        For Each c In hitRooms.Cells
            Select Case c.Address(False, False)
                Case Replace(P1_ROOM, "$", ""): Me.Range(P1_OWNER).Value = IIf(Len(Trim$(c.Value & "")) = 0, "", OwnerFromNameSheet(c.Value))
                Case Replace(P2_ROOM, "$", ""): Me.Range(P2_OWNER).Value = IIf(Len(Trim$(c.Value & "")) = 0, "", OwnerFromNameSheet(c.Value))
                Case Replace(P3_ROOM, "$", ""): Me.Range(P3_OWNER).Value = IIf(Len(Trim$(c.Value & "")) = 0, "", OwnerFromNameSheet(c.Value))
            End Select
        Next c
        AutoFillCurrentMonth Me
    End If

    Dim hitDates As Range
    Set hitDates = Intersect(Target, Me.Range(P1_DATE & "," & P2_DATE & "," & P3_DATE))
    If Not hitDates Is Nothing Then
        hitDates.NumberFormat = "mm/yyyy"
    End If

    CalcAll

Done:
    Application.EnableEvents = True
End Sub


