VERSION 5.00
Begin VB.MDIForm Main_frm 
   BackColor       =   &H8000000C&
   Caption         =   "Penilaian Guru"
   ClientHeight    =   7710
   ClientLeft      =   120
   ClientTop       =   765
   ClientWidth     =   13875
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   Begin VB.Menu user_mnu 
      Caption         =   "Data User"
   End
   Begin VB.Menu guru_mnu 
      Caption         =   "Data Guru"
   End
   Begin VB.Menu kategori_mnu 
      Caption         =   "Kriteria Nilai"
   End
   Begin VB.Menu Periode_mnu 
      Caption         =   "Periode"
   End
   Begin VB.Menu tbNilai_mnu 
      Caption         =   "Tabel Penilaian"
   End
   Begin VB.Menu InputNilai_mnu 
      Caption         =   "Input Nilai"
   End
   Begin VB.Menu Lap_mnu 
      Caption         =   "Laporan"
   End
   Begin VB.Menu logout_mnu 
      Caption         =   "Logout"
   End
   Begin VB.Menu win_mnu 
      Caption         =   "Windows"
      WindowList      =   -1  'True
   End
   Begin VB.Menu exit_mnu 
      Caption         =   "Keluar"
   End
End
Attribute VB_Name = "Main_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub exit_mnu_Click()
End
End Sub

Private Sub guru_mnu_Click()
With Guru_frm
    .Top = 0
    .Left = 0
    .Show
End With
End Sub

Private Sub InputNilai_mnu_Click()
With Penilaian_frm
    .Top = 0
    .Left = 0
    .Show
End With
End Sub

Private Sub kategori_mnu_Click()
With kategori_frm
    .Top = 0
    .Left = 0
    .Show
End With
End Sub

Private Sub Lap_mnu_Click()
With lap_frm
    .Top = 0
    .Left = 0
    .Show
End With
End Sub

Private Sub logout_mnu_Click()
Login_frm.Show
Unload Me
End Sub

Private Sub Periode_mnu_Click()
With Periode_frm
    .Top = 0
    .Left = 0
    .Show
End With
End Sub

Private Sub tbNilai_mnu_Click()
With TabelNilai_frm
    .Top = 0
    .Left = 0
    .Show
End With
End Sub

Private Sub user_mnu_Click()
With User_frm
    .Top = 0
    .Left = 0
    .Show
End With
End Sub
