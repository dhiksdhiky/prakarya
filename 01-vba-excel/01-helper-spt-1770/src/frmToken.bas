
Option Explicit
Dim wb1 As Workbook
Dim wspark, wsspt, wsset As Worksheet
Dim objinputs As Object
Dim mulai, lebar, jawab As Integer
Dim lama, selesai As Single
Dim fra As Frame
Dim lasttoken As Long

Private Sub cmdOK_Click()
Set wsspt = ThisWorkbook.Worksheets("spt")
Set wsset = ThisWorkbook.Worksheets("setting")

If wsset.Range("C3").Value = vbNullString Then

    MsgBox "Harap lengkapi dulu data SIKKA di menu setting"
    wsset.Activate
    Unload Me
    Exit Sub

End If

lasttoken = wsspt.Cells(Rows.Count, "C").End(xlUp).Row



wsspt.Range("B27:C" & lasttoken + 1).ClearContents

With Me

    .lblTimer.Visible = True
    
    .lblKet.Visible = True
    
    .fraTimer.Visible = True
    
    .cmdOK.Visible = False
    
    .lblTunggu.Font.Size = 10
        
    .lblTunggu.Caption = "mohon ditunggu ya. sing sabar... gak lama kok"
    
    .lblKet.Caption = ""
    
    .Height = 1000
        
End With

'If wsspt.Range("C26").Value = "TOKEN EFORM" Then

Call GetCode

'Else

'Call getefilingcode

'End If

Unload Me

End Sub


Private Sub UserForm_Initialize()

With Me

    .cmdOK.Visible = True
    
    .cmdOK.Left = 54
    
    .cmdOK.Top = 27
    
    .Height = 89
    
    .Width = 500
    
    .lblTimer.Visible = False
    
    .lblKet.Visible = False
    
    .fraTimer.Visible = False
    
    .fraTimer.Caption = Format(0, "0%")
    
    .lblTimer.Width = 0
    
    .lblTunggu.Caption = "Proses akan membutuhkan waktu beberapa detik"

    .lblTunggu.Font.Size = 9.5
            
End With
    
End Sub
