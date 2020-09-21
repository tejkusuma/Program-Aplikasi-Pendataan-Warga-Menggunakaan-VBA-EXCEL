VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} BukuInduk 
   Caption         =   "BUKU INDUK"
   ClientHeight    =   7245
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14115
   OleObjectBlob   =   "BukuInduk.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "BukuInduk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub keluar_Click()
Unload Me
End Sub

Private Sub submit_Click()
Dim Kolom As Long
Dim lembar As Worksheet
Set lembar = Worksheets("Buku_Penduduk")
Kolom = lembar.Cells(Rows.Count, 11) _
    .End(xlUp).Offset(1, 0).Row


'ngecek kolom kosong
If Trim(Me.nama_lengkap.Value) = "" Then
Me.nama_lengkap.SetFocus
MsgBox "Nama Lengkap Wajib Diisi"
Exit Sub
End If

If Trim(Me.nik.Value) = "" Then
Me.nik.SetFocus
MsgBox "NIK Wajib Diisi"
Exit Sub
End If

If Trim(Me.no_kk.Value) = "" Then
Me.no_kk.SetFocus
MsgBox "No KK Wajib Diisi"
Exit Sub
End If

If Trim(Me.tgl_lahir.Value) = "" Then
Me.tgl_lahir.SetFocus
MsgBox "Tanggal Lahir Wajib Diisi"
Exit Sub
End If

If Trim(Me.agama.Value) = "" Then
Me.agama.SetFocus
MsgBox "Agama Wajib Diisi"
Exit Sub
End If

If Trim(Me.pend_terakhir.Value) = "" Then
Me.pend_terakhir.SetFocus
MsgBox "Pendidikan Terakhir Wajib Diisi"
Exit Sub
End If

If Trim(Me.pekerjaan.Value) = "" Then
Me.pekerjaan.SetFocus
MsgBox "Pekerjaan Wajib Diisi"
Exit Sub
End If

If Trim(Me.kedudukan.Value) = "" Then
Me.kedudukan.SetFocus
MsgBox "Kedudukan dalam Keluarga Wajib Diisi"
Exit Sub
End If

If Trim(Me.nama_ayah.Value) = "" Then
Me.nama_ayah.SetFocus
MsgBox "Nama Ayah Wajib Diisi"
Exit Sub
End If

If Trim(Me.nama_ibu.Value) = "" Then
Me.nama_ibu.SetFocus
MsgBox "Nama Ibu Wajib Diisi"
Exit Sub
End If

'coba
If OptionButton1.Value = True Then
lembar.Cells(Kolom, 22).Value = "Laki-laki"
Else
lembar.Cells(Kolom, 22).Value = "Perempuan"
End If

'fungsi memasukkan data ke Cells
lembar.Cells(Kolom, 11).Value = Me.nama_lengkap
lembar.Cells(Kolom, 12).Value = Me.nik
lembar.Cells(Kolom, 13).Value = Me.no_kk
lembar.Cells(Kolom, 14).Value = Me.tgl_lahir
lembar.Cells(Kolom, 15).Value = Me.pindah_alamat
lembar.Cells(Kolom, 16).Value = Me.tgl_wafat
lembar.Cells(Kolom, 17).Value = Me.wafat_usia
lembar.Cells(Kolom, 18).Value = Me.jdw_pilkada
lembar.Cells(Kolom, 23).Value = statuskawin.Value
lembar.Cells(Kolom, 24).Value = Me.agama
lembar.Cells(Kolom, 25).Value = Me.pend_terakhir
lembar.Cells(Kolom, 26).Value = Me.pekerjaan
lembar.Cells(Kolom, 27).Value = Me.kedudukan
lembar.Cells(Kolom, 28).Value = Me.nama_ayah
lembar.Cells(Kolom, 29).Value = Me.nama_ibu
lembar.Cells(Kolom, 30).Value = Me.tgl_kk
lembar.Cells(Kolom, 31).Value = Me.no_hp


'hapus data
Me.nama_lengkap.Value = ""
Me.nik.Value = ""
Me.no_kk.Value = ""
Me.tgl_lahir.Value = ""
Me.pindah_alamat.Value = ""
Me.tgl_wafat.Value = ""
Me.wafat_usia.Value = ""
Me.agama.Value = ""
Me.pend_terakhir.Value = ""
Me.pekerjaan.Value = ""
Me.kedudukan.Value = ""
Me.nama_ayah.Value = ""
Me.nama_ibu.Value = ""
Me.tgl_kk.Value = ""
Me.no_hp.Value = ""

End Sub

Private Sub tgl_kk_Change()
If tgl_kk.TextLength = 2 Or tgl_kk.TextLength = 5 Then
    tgl_kk.Text = tgl_kk.Text + "-"
End If
End Sub

Private Sub tgl_lahir_Change()
If tgl_lahir.TextLength = 2 Or tgl_lahir.TextLength = 5 Then
    tgl_lahir.Text = tgl_lahir.Text + "-"
End If
End Sub

Private Sub tgl_wafat_Change()
If tgl_wafat.TextLength = 2 Or tgl_wafat.TextLength = 5 Then
    tgl_wafat.Text = tgl_wafat.Text + "-"
End If
End Sub

Private Sub UserForm_Initialize()
With statuskawin
.AddItem "Kawin"
.AddItem "Belum Kawin"
.AddItem "Janda"
.AddItem "Duda"
End With
End Sub
