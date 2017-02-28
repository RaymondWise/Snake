VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GameSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Private Sub Worksheet_SelectionChange(ByVal target As Range)
    If Intersect(target, GameSheet.Range("Board")) Is Nothing Then Exit Sub
    If Range("FirstMove").value > 0 Then
        GameSheet.Activate
        Range("FirstMove").value = 0
        Application.OnTime Now + TimeValue("00:00:02"), "Start_Timing"
        Exit Sub
    End If
    If Range("LegalMove") = 0 Then
        Range("LegalMove") = 1
        Exit Sub
    End If

    Dim storedLocation() As String
    storedLocation = Split(Range("Position"), ",")
    Dim currentLocation As Range
    Set currentLocation = Cells(CLng(storedLocation(0)), CLng(storedLocation(1)))
    Main currentLocation, target



End Sub

