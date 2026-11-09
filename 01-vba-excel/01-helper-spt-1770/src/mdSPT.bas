Attribute VB_Name = "mdSPT"
Option Explicit
Dim wb, wb1, wbpegawai, wbform As Workbook
Dim ws, wshome, wsform, wsbil, wsspt, wsinv, wssik, wsregis, wsdb, wsset, wspark, wspegawai As Worksheet
Dim i, shini, lastdata, lasttoken As Long
Dim xrng, xcell As Range
Dim path, nip, pass, npwp, eform, selesai As String
Dim objinputs, htmldoc As Object
Dim fra As Frame
Dim web As WebBrowser

Sub kuncish()
If ThisWorkbook.Worksheets("spt").ProtectContents = True Then
ThisWorkbook.Worksheets("spt").Unprotect
Else
ThisWorkbook.Worksheets("spt").Protect
End If
End Sub

Sub locksh()
ThisWorkbook.Worksheets("spt").Protect
End Sub
Sub nextsh() '__next ke spt halaman berikutnya
shini = ActiveSheet.Index
If shini = ThisWorkbook.Worksheets("1770").Index + 5 Then
    
    MsgBox "sudah halaman terakhir ini bro"
    
Else

    ThisWorkbook.Worksheets(shini + 1).Activate
    
End If
End Sub

Sub prevsh() '__mundur ke spt halaman sebelumnya
shini = ActiveSheet.Index
If shini = ThisWorkbook.Worksheets("1770").Index Then
    MsgBox "sudah di SPT Induk"
Else
    ThisWorkbook.Worksheets(shini - 1).Activate
End If
End Sub

Sub home() '__buka home
ThisWorkbook.Worksheets("spt").Activate
End Sub

Sub pp()  '__buka pp23
ThisWorkbook.Worksheets("PP23").Activate
End Sub

Sub induk()  '__buka spt induk
ThisWorkbook.Worksheets("1770").Activate
End Sub

Sub hitung() '___menghitung pph final pp23 (otomatis mana yang terisi apakah bruto atau pph nyaa, kereeennn
Set wb1 = ThisWorkbook
Set wsspt = wb1.Worksheets("spt")
Set xrng = wsspt.Range("I7:J" & wsspt.Cells(Rows.Count, "J").End(xlUp).Row - 1)

For Each xcell In xrng

    If xcell.Value <> vbNullString Then
        
        If xcell.Column = 9 Then
        
            xcell.Offset(0, 1).Value = xcell.Value / 200
        
        ElseIf xcell.Column = 10 Then
        
            xcell.Offset(0, -1).Value = xcell.Value * 200
            
        End If
        
    End If
    
    Debug.Print xcell.Column
    
Next xcell

End Sub

Sub makecsv() '___make csv to upload
Application.DisplayAlerts = False
Dim pathfolder As String
Set wb1 = ThisWorkbook
Set wspark = wb1.Worksheets("print_area")
Set wsspt = wb1.Worksheets("spt")
pathfolder = wsspt.Range("C22").Value '___path folder simpan file csv

If pathfolder = vbNullString Then

    MsgBox "Mohon seting folder lokasi file CSV terlebih dahulu" & vbNewLine & "Klik tombol Set Folder", vbInformation
    Exit Sub
    
End If

lastdata = wsspt.Range("J" & Rows.Count).End(xlUp).Row - 1 '___data bruto terakhir

With wspark

    .UsedRange.Delete

    .Cells(1, 1).Value = "NPWP"
    
    .Cells(1, 2).Value = "Masa Pajak"
    
    .Cells(1, 3).Value = "Alamat"
    
    .Cells(1, 4).Value = "Bruto"
    
    wsspt.Range("C4").Copy

    .Range("A2:A" & lastdata - 5).PasteSpecial xlPasteValues '___paste value npwp
    
    wsspt.Range("C8").Copy
    
    .Range("C2:C" & lastdata - 5).PasteSpecial xlPasteValues '___paste value alamat
    
    wsspt.Range("G7:G" & lastdata).Copy
    
    .Range("B2").PasteSpecial xlPasteValues '___paste value masa pajak
    
    wsspt.Range("I7:I" & lastdata).Copy
    
    .Range("D2").PasteSpecial xlPasteValues '___paste value penghasilan bruto

    .Copy
    
End With

With ActiveWorkbook
    .SaveAs pathfolder & "\Upload PP23.csv", FileFormat:=xlCSV, local:=True '___tanpa filename:= malah bisa semicolon ;
    .Close
End With

If wsspt.Range("U8").Value <> vbNullString Then

    lastdata = wsspt.Range("u" & Rows.Count).End(xlUp).Row '___data bruto terakhir

    With wspark

        .UsedRange.Delete
        
        .Cells(1, 1).Value = "Kode Harta"
    
        .Cells(1, 2).Value = "Nama Harta"
    
        .Cells(1, 3).Value = "Tahun Perolehan"
    
        .Cells(1, 4).Value = "Tahun Perolehan"
        
        .Cells(1, 5).Value = "Keterangan"

        wsspt.Range("R8:V" & lastdata).Copy

        .Range("A2").PasteSpecial xlPasteValues '___paste value kolom harta
    
        .Copy
    
    End With

    With ActiveWorkbook
        .SaveAs pathfolder & "\Upload Harta.csv", FileFormat:=xlCSV, local:=True '___tanpa filename:= malah bisa semicolon ;
        .Close
    End With

        MsgBox "Berhasil membuat file upload CSV penghasilan bruto dan harta."

wsspt.Activate

Else

    MsgBox "Berhasil membuat file upload CSV penghasilan bruto"

wsspt.Activate

End If

Application.DisplayAlerts = True

End Sub

Sub hapusbruto()
Application.DisplayAlerts = False
Set wb1 = ThisWorkbook
Set wspark = wb1.Worksheets("print_area")
Set wsspt = wb1.Worksheets("spt")
lastdata = wsspt.Range("J" & Rows.Count).End(xlUp).Row - 1 '___data bruto terakhir

If Range("M2").Value = 1 Then

    With wsspt
        .Range("I7:J" & lastdata).name = "bruto"
        .Range("bruto").ClearContents
    End With

ElseIf Range("M2").Value = 2 Then

    lastdata = wsspt.Range("R" & Rows.Count).End(xlUp).Row '___data harta terakhir
    wsspt.Range("R8:V28").name = "harta"
    wsspt.Range("harta").ClearContents
    
End If

Application.DisplayAlerts = True
End Sub

Sub bukaEFORM()
path = Range("C22").Value

If path = vbNullString Then GoTo error

    eform = "explorer.exe " & path & "\FORM_" & Range("C11") & "_" & Range("C4").Value & "_" & Range("C9") & "_" & Range("C12").Value & ".pdf"

    VBA.Shell eform

    MsgBox "Mohon ditunggu form akan dibuka..." & vbNewLine & "Pastikan :" & vbNewLine & "1. NPWP, Jenis, Tahun, status SPT sudah sesuai" & vbNewLine & "2. EFORM sudah terdownload" & vbNewLine & "Apabila dalam 6 detik tidak terbuka, silakan buka secara manual."

    Exit Sub
    
error:

    MsgBox "Mohon seting folder lokasi file EFORM terlebih dahulu" & vbNewLine & "Klik tombol Set Folder", vbInformation

End Sub

Sub FldrDownload()
Set wb1 = ThisWorkbook
Set wsspt = wb1.Worksheets("spt")
Dim PDFFldr As FileDialog
Set PDFFldr = Application.FileDialog(msoFileDialogFolderPicker)
With PDFFldr
    .Title = "Pilih folder lokasi EFORM"
    If .Show <> -1 Then GoTo Nosel
    wsspt.Range("C22").Value = .SelectedItems(1)
End With
MsgBox "Berhasil!" & vbNewLine & "Lokasi penyimpangan file CSV, file PDF lampiran, dan file EFORM adalah: " & wsspt.Range("C22").Value
Nosel:
End Sub

Sub showOmzet()
Set wsspt = ThisWorkbook.Worksheets("spt")

wsspt.Columns("F:BZ").EntireColumn.Hidden = True
wsspt.Shapes("Rectangle 23").Visible = True

With wsspt

    If wsspt.Range("M2").Value = 2 Then

        .Columns("E:K").EntireColumn.Hidden = False '___show peredaran bruto
        .Range("M2").Value = 1
    
    ElseIf wsspt.Range("M2").Value = 1 Then
    
        .Columns("Q:AF").EntireColumn.Hidden = False '___show daftar harta
        .Range("M2").Value = 2
    
    End If

End With

End Sub

Sub showGrossup()
Set wsspt = ThisWorkbook.Worksheets("spt")
'Call kuncish
wsspt.Columns("F:BZ").EntireColumn.Hidden = True
wsspt.Shapes("Rectangle 23").Visible = False

With wsspt

    If wsspt.Range("AR2").Value = 2 Then

        .Columns("AH:AO").EntireColumn.Hidden = False '___show rumus nihil
        .Range("AR2").Value = 1
    
    ElseIf wsspt.Range("AR2").Value = 1 Then
    
        .Columns("AU:BI").EntireColumn.Hidden = False '___show rumus grossup
        .Range("AR2").Value = 2
    
    End If

End With

'Call locksh

End Sub

Sub ShowSPTmasa()

ThisWorkbook.Worksheets("1721").Activate

End Sub

Sub pdfomzet() '___buat pdf file penghasilan bruto, sheet pp23
Application.DisplayAlerts = False
Dim wspp23 As Worksheet
Set wsspt = ThisWorkbook.Worksheets("spt")
Set wspp23 = ThisWorkbook.Worksheets("PP23")

path = wsspt.Range("C22").Value

If path = vbNullString Then

    MsgBox "Mohon seting folder lokasi file PDF terlebih dahulu" & vbNewLine & "Klik tombol Set Folder", vbInformation
    
    Exit Sub
    
End If

On Error GoTo masihkebukabro

    wspp23.ExportAsFixedFormat Type:=xlTypePDF, Filename:=path & "\LAPOR.pdf", Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False

MsgBox "Berhasil membuat file upload penghasilan bruto."

Application.DisplayAlerts = True

Exit Sub

masihkebukabro:

MsgBox "Terdapat file PDF dengan nama LAPOR yang terbuka, tutup dulu", vbCritical, "Gagal membuat file upload"

Application.DisplayAlerts = True

End Sub

Sub printSPT()
Dim pss1770 As Long

pss1770 = ThisWorkbook.Worksheets("1770").Index

If MsgBox("Formulir SPT Tahunan 1770 akan dicetak, pastikan ukuran kertas sudah sesuai (F4).", vbOKCancel + vbExclamation) = vbOK Then
    
For i = pss1770 To pss1770 + 6 '___trik supaya bisa print dinamis tanpa khawatir salah print ketika sheet bertambah

    ThisWorkbook.Worksheets(i).PrintOut copies:=1, collate:=True, IgnorePrintAreas:=False

Next i

End If
    
End Sub

Sub buatNIHIL() '___code untuk mencari nilai penghasilan dari pekerjaan bebas/usaha agar PPh nya NIHIL
'Call kuncish
Set wsspt = ThisWorkbook.Worksheets("spt")
    Application.CutCopyMode = False
    wsspt.Range("AN23").GoalSeek goal:=0, changingcell:=wsspt.Range("AN5")
'Call kuncish
End Sub

Sub grossup()
'Call kuncish
Set wsspt = ThisWorkbook.Worksheets("spt")

With wsspt
    .Range("AZ15").ClearContents
    .Range("AZ34").GoalSeek goal:=Range("AZ24"), changingcell:=Range("AZ15")
End With
'Call locksh
End Sub

Sub seting()
Set wsset = ThisWorkbook.Worksheets("setting")
wsset.Activate
End Sub

'============================================================================================================================================================================================
'BATAS SERIUS BOS.....by Eka Adhikara Rahim.....BATAS SERIUS BOS.....by Eka Adhikara Rahim.....BATAS SERIUS BOS.....by Eka Adhikara Rahim.....BATAS SERIUS BOS.....by Eka Adhikara Rahim.....
'============================================================================================================================================================================================


Sub getToken()

frmToken.Show

End Sub

Sub GetCode()
Set wsspt = ThisWorkbook.Worksheets("spt")
Set wsset = ThisWorkbook.Worksheets("setting")
Set wspark = ThisWorkbook.Worksheets("print_area")
Set web = frmToken.WebBrowser1

web.Silent = True

'On Error Resume Next

With wsset

    nip = .Range("C3").Value
    
    pass = .Range("C4").Value
    
End With

    npwp = wsspt.Range("C4").Value
    
web.Navigate "http://INTERNAL_MONITORING_HOST/cektoken/login.php" '---mengarahkan web ke link (URL disamarkan untuk keamanan)
frmToken.lblKet.Caption = "Menuju laman monitoring..."  '___keterangan loading

Do While web.Busy Or web.ReadyState <> 4 '--tunggu sampai loading selesai
    DoEvents
Loop

selesai = 0.1
frmToken.fraTimer.Caption = Format(selesai, "0%")
frmToken.lblTimer.Width = 20

'On Error GoTo skipaja

Set htmldoc = web.Document

    'htmldoc.all.overridelink.Click
    
Do While web.Busy Or web.ReadyState <> 4 '--tunggu sampai loading selesai
    DoEvents
Loop

'skipaja:

'-- masukkan username dan password
frmToken.lblKet.Caption = "Login akun SIKKA..."  '___keterangan loading

On Error GoTo vpn

htmldoc.forms(0).UserName.Value = nip
    
htmldoc.forms(0).Password.Value = pass

selesai = 0.2
frmToken.fraTimer.Caption = Format(selesai, "0%")
frmToken.lblTimer.Width = 40
    
Set objinputs = web.Document.getElementsByclassname("btn btn-focus m-btn m-btn--pill m-btn--custom m-btn--air m-login__btn m-login__btn--primary")(0)

objinputs.Click '---klik tombol login sesuai sikka (data pegawai)
    
Do While web.Busy Or web.ReadyState <> 4 '--tunggu sampai loading selesai
    DoEvents
Loop

selesai = 0.3
frmToken.fraTimer.Caption = Format(selesai, "0%")
frmToken.lblTimer.Width = 60

frmToken.lblKet.Caption = "Membuka laman token e-Filing..."  '___keterangan loading

'==============================================BATAS EFORM

If wsspt.Range("C26").Value = "TOKEN EFORM" Then

    frmToken.lblKet.Caption = "Membuka laman token EFORM..."  '___keterangan loading

    web.Navigate "http://INTERNAL_MONITORING_HOST/cektoken/tokeneform.php" '---mengarahkan web ke link (URL disamarkan untuk keamanan)
            
    selesai = 0.4
    frmToken.fraTimer.Caption = Format(selesai, "0%")
    frmToken.lblTimer.Width = 80

End If

'==============================================BATAS

    On Error GoTo ulang3

    Application.Wait (Now + TimeValue("00:00:01"))

    Do While web.Busy Or web.ReadyState <> 4 '--tunggu sampai loading selesai
        DoEvents
    Loop

Set htmldoc = web.Document
    
    htmldoc.forms(0).npwp15.Value = npwp '---input npwp
    
selesai = 0.5
frmToken.fraTimer.Caption = Format(selesai, "0%")
frmToken.lblTimer.Width = 100
frmToken.lblKet.Caption = "Input NPWP..."  '___keterangan loading

Set objinputs = web.Document.getElementsByclassname("btn btn-success")(0) '---cari tombol by class name html

    objinputs.Click '---klik cari token
    
    Application.Wait (Now + TimeValue("00:00:01"))
    
selesai = 0.8
frmToken.fraTimer.Caption = Format(selesai, "0%")
frmToken.lblTimer.Width = 120

frmToken.lblKet.Caption = "Tarik mang.., eh tarik data token..."  '___keterangan loading
    
    Application.Wait (Now + TimeValue("00:00:01"))
    
Do While web.Busy Or web.ReadyState <> 4 '--tunggu sampai loading selesai
    DoEvents
Loop

selesai = 0.85
frmToken.fraTimer.Caption = Format(selesai, "0%")
frmToken.lblTimer.Width = 140
    
    wspark.Activate
    
selesai = 0.9
frmToken.fraTimer.Caption = Format(selesai, "0%")
frmToken.lblTimer.Width = 160

    'MsgBox "copy token berhasil, silakan paste ke eform"
    
    wspark.UsedRange.ClearContents
    
    web.ExecWB 17, 0
    
    web.ExecWB 12, 0 '---copy text di webbrowser
    
    wspark.Range("A1").Select
    
frmToken.lblKet.Caption = "Parkir kode aktivasi..."  '___keterangan loading
    
selesai = 0.95
frmToken.fraTimer.Caption = Format(selesai, "0%")
frmToken.lblTimer.Width = 180
    
    'On Error GoTo ulang
'ulang:
    ActiveSheet.PasteSpecial displayasicon:=False, nohtmlformatting:=True
    
frmToken.lblKet.Caption = "Pindah token..."  '___keterangan loading
    
    Application.CutCopyMode = False
    
    wspark.Rows("1:26").Select

    Selection.Delete Shift:=xlUp
    
    lasttoken = wspark.Cells(Rows.Count, "E").End(xlUp).Row
    
    If wsspt.Range("C26").Value = "TOKEN EFORM" Then
            
        Range("G2:H" & lasttoken).Copy
        
        wsspt.Range("B27").PasteSpecial
    
    Else
    
        wspark.Range("E2:E" & lasttoken).Copy
        'Selection.Copy
        wsspt.Range("B27").PasteSpecial
        wspark.Range("I2:I" & lasttoken).Copy
        wsspt.Range("C27").PasteSpecial
    
    End If
    
    Application.CutCopyMode = False
    
    Application.CutCopyMode = False

selesai = 1
frmToken.fraTimer.Caption = Format(selesai, "0%")
frmToken.lblTimer.Width = 200

wsspt.Activate

Exit Sub

ulang3:
MsgBox "Pastikan NIP/Pass SIKKA sudah benar, close file dan coba lagi", vbInformation, "Gagal generate token..."
Exit Sub

vpn:
MsgBox "Pastikan sudah terhubung intranet, NIP/Pass SIKKA sudah benar, dan ulangi lagi", vbInformation, "Gagal terhubung ke intranet..."

wsspt.Activate

End Sub
