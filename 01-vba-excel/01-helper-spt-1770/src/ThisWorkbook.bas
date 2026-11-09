Attribute VB_Name = "ThisWorkbook"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Private Sub Workbook_Activate()
    With Application
        .DisplayFullScreen = True
        .DisplayFormulaBar = False
        .DisplayStatusBar = False
    End With
    With ActiveWindow
        .DisplayHorizontalScrollBar = False     'Make space and clear screen
        .DisplayVerticalScrollBar = False
        .DisplayWorkbookTabs = False
        .DisplayHeadings = False
        .DisplayGridlines = False
    End With
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
Application.DisplayAlerts = False
ThisWorkbook.Worksheets("setting").Range("C8").Value = ThisWorkbook.path
ThisWorkbook.Save
Application.DisplayAlerts = True
End Sub

Private Sub Workbook_Deactivate()
    With Application
        .DisplayFullScreen = False
        .DisplayFormulaBar = True
        .DisplayStatusBar = True
    End With
    With ActiveWindow
        .DisplayHorizontalScrollBar = True
        .DisplayVerticalScrollBar = True
        .DisplayWorkbookTabs = True
        .DisplayHeadings = True
        .DisplayGridlines = True
    End With
End Sub

Private Sub Workbook_Open()
If ThisWorkbook.Worksheets("setting").Range("C8").Value <> ThisWorkbook.path Then
ThisWorkbook.Worksheets("setting").Range("C3:C8").ClearContents
ThisWorkbook.Worksheets("spt").Range("C22").Value = vbNullString
End If
ThisWorkbook.Worksheets("spt").Activate
End Sub
