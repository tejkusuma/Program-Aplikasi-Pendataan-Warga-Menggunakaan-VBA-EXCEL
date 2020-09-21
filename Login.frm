VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Login 
   Caption         =   "Login"
   ClientHeight    =   2850
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "Login.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_masuk_Click()
Set sh = sheets("Login")
If username.Value = "" Then
    MsgBox "Silahkan Masukkan User Name", _
    vbExclamation + vbOKOnly, "Blank User Name"
    username.SetFocus
    Exit Sub
ElseIf password.Value = "" Then
    MsgBox "Silahkan Masukkan Password", _
    vbExclamation + vbOKOnly, "Blank Password"
    password.SetFocus
    Exit Sub
ElseIf username.Value <> sh.Range("Y12").Value Then
    MsgBox "User Name Salah/Tidak Terdaftar", _
    vbCritical + vbOKOnly, "Error User Name"
    username.SetFocus
    Exit Sub
ElseIf password.Value <> sh.Range("Z12").Value Then
    MsgBox "Password Salah, Silahkan ulangi lagi", _
    vbCritical + vbOKOnly, "Error Password"
    password.SetFocus
    Exit Sub
End If
MsgBox "Selamat Anda berhasil Login", _
    vbInformation + vbOKOnly, "Login Sukses"
Unload Me
sheets("Buku_Penduduk").Activate
End Sub

Private Sub CommandButton1_Click()
Unload Me
End Sub
