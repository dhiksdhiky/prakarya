Option Explicit
Dim irange, icells As Range

Private Sub Worksheet_Change(ByVal Target As Range)
Set irange = Range("spt")
If Target.Address = "$C$4" Then
    Range("harta").ClearContents
    Target.Replace ".", ""
    Target.Replace "-", ""
    For Each icells In irange
    icells.BorderAround LineStyle:=xlContinuous, Weight:=xlMedium
    Next icells
End If
End Sub
