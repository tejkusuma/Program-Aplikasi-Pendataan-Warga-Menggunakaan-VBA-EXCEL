VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} EditData 
   Caption         =   "EDIT DATA KEPENDUDUKAN"
   ClientHeight    =   8250.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14790
   OleObjectBlob   =   "EditData.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "EditData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cari_data_Click()

Dim cari
Dim CellTujuan As Range
cari = cari_nama.Text
Set CellTujuan = Range("K11:K500", "L12:L500").Find(What:=cari)

If Not CellTujuan Is Nothing Then
nama_lengkap.Text = Cells(CellTujuan.Row, 11)
nik.Text = Cells(CellTujuan.Row, 12)
no_kk.Text = Cells(CellTujuan.Row, 13)
tgl_lahir.Text = Cells(CellTujuan.Row, 14)
pindah_alamat.Text = Cells(CellTujuan.Row, 15)
tgl_wafat.Text = Cells(CellTujuan.Row, 16)
wafat_usia.Text = Cells(CellTujuan.Row, 17)
jdw_pilkada.Text = Cells(CellTujuan.Row, 18)
statuskawin.Text = Cells(CellTujuan.Row, 23)
agama.Text = Cells(CellTujuan.Row, 24)
pend_terakhir.Text = Cells(CellTujuan.Row, 25)
pekerjaan.Text = Cells(CellTujuan.Row, 26)
kedudukan.Text = Cells(CellTujuan.Row, 27)
nama_ayah.Text = Cells(CellTujuan.Row, 28)
nama_ibu.Text = Cells(CellTujuan.Row, 29)
tgl_kk.Text = Cells(CellTujuan.Row, 30)
no_hp.Text = Cells(CellTujuan.Row, 31)

If Cells(CellTujuan.Row, 22) = "Laki-laki" Then
OptionButton1.Value = True
ElseIf Cells(CellTujuan.Row, 22) = "Perempuan" Then
OptionButton2.Value = True
End If

Else
MsgBox "Maaf pencarian tidak ditemukan !"
End If
    
End Sub



Private Sub edit_Click()

Dim CellTujuan As Range
cari = cari_nama.Text
Set CellTujuan = Range("K:K").Find(What:=cari)

Cells(CellTujuan.Row, 11) = nama_lengkap
Cells(CellTujuan.Row, 12) = nik
Cells(CellTujuan.Row, 13) = no_kk
Cells(CellTujuan.Row, 14) = tgl_lahir
Cells(CellTujuan.Row, 15) = pindah_alamat
Cells(CellTujuan.Row, 16) = tgl_wafat
Cells(CellTujuan.Row, 17) = wafat_usia
Cells(CellTujuan.Row, 18) = jdw_pilkada
Cells(CellTujuan.Row, 23) = statuskawin
Cells(CellTujuan.Row, 24) = agama
Cells(CellTujuan.Row, 25) = pend_terakhir
Cells(CellTujuan.Row, 26) = pekerjaan
Cells(CellTujuan.Row, 27) = kedudukan
Cells(CellTujuan.Row, 28) = nama_ayah
Cells(CellTujuan.Row, 29) = nama_ibu
Cells(CellTujuan.Row, 30) = tgl_kk
Cells(CellTujuan.Row, 31) = no_hp

'coba
If OptionButton1.Value = True Then
Cells(CellTujuan.Row, 22) = "Laki-laki"
Else
Cells(CellTujuan.Row, 22) = "Perempuan"
End If

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

MsgBox "Data Berhasil Diubah"

End Sub

Private Sub Hapus_Click()
Dim cari
Dim CellTujuan As Range
cari = cari_nama.Text
Set CellTujuan = Range("K:K").Find(What:=cari)

Rows(CellTujuan.Row).Delete Shift:=xlUp
End Sub

Private Sub keluar_Click()
Unload Me
End Sub

Private Sub UserForm_Initialize()
With statuskawin
.AddItem "Kawin"
.AddItem "Belum Kawin"
.AddItem "Janda"
.AddItem "Duda"
End With

cari_nama.SetFocus

End Sub

