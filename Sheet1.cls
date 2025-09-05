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

Private Sub Worksheet_Change(ByVal Target As Range)
    If Application.EnableEvents = False Then Exit Sub
    Application.EnableEvents = False
    On Error GoTo Done

    CalcAll

Done:
    Application.EnableEvents = True
End Sub
