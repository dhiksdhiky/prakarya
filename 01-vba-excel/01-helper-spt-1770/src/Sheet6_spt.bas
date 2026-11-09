Attribute VB_Name = "Sheet6"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
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
