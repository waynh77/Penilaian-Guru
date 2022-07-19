VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form TabelNilai_frm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tabel Penilaian"
   ClientHeight    =   4995
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   8160
   Begin VB.Data Data2 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   120
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   5280
      TabIndex        =   9
      Top             =   960
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1080
      TabIndex        =   8
      Text            =   "Combo1"
      Top             =   480
      Width           =   3255
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3480
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton cmd3 
      Caption         =   "Hapus"
      Height          =   495
      Left            =   6840
      TabIndex        =   6
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton cmd2 
      Caption         =   "Edit"
      Height          =   495
      Left            =   5640
      TabIndex        =   5
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton cmd1 
      Caption         =   "Tambah"
      Height          =   495
      Left            =   4440
      TabIndex        =   4
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   960
      Width           =   3255
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "TabelNilai_frm.frx":0000
      Height          =   3255
      Left            =   120
      OleObjectBlob   =   "TabelNilai_frm.frx":0014
      TabIndex        =   7
      Top             =   1560
      Width           =   7815
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Bobot Nilai"
      Height          =   195
      Index           =   3
      Left            =   4440
      TabIndex        =   10
      Top             =   960
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Keterangan"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   825
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Kriteria"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   480
   End
   Begin VB.Label lblId 
      AutoSize        =   -1  'True
      Caption         =   "Id"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   135
   End
End
Attribute VB_Name = "TabelNilai_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd1_Click()
BtnSimpan
Kosong
End Sub

Sub isiCbo()
Combo1.Clear
Data2.Refresh
With Data2.Recordset
    If Not .BOF Then
        .MoveFirst
        Do While Not .EOF
            Combo1.AddItem !kategori
            .MoveNext
        Loop
        Combo1.ListIndex = 0
    End If
End With
End Sub

Private Sub cmd2_Click()
If cmd2.Caption = "Edit" Then
    If Data1.Recordset.RecordCount > 0 Then
        BtnSimpan
    Else
        MsgBox "Data Kosong", vbInformation, "Gagal"
    End If
Else
    If Text1 = "" Then
        MsgBox "Kategori tidak boleh kosong", vbInformation, "Validasi"
    Else
        With Data1.Recordset
            If lblId.Caption = "" Then
                .AddNew
            Else
                .Edit
            End If
            !keterangan = Text1
            !kategori = Combo1.Text
            !bobot = Text3
            .Update
        End With
        BtnAwal
        Data1.Refresh
    End If
End If
End Sub

Private Sub cmd3_Click()
If cmd3.Caption = "Hapus" Then
    If Data1.Recordset.BOF Then
        MsgBox "Data kosong", vbInformation, "Validasi"
    Else
        Dim tny As String
        tny = MsgBox("Apakah anda yakin hapus?", vbYesNo, "Hapus")
        If tny = vbYes Then
            Data1.Recordset.Delete
            Data1.Refresh
        End If
    End If
Else
    BtnAwal
End If
End Sub

Private Sub Data1_Reposition()
If DBGrid1.Enabled = True Then
    Isi
End If
End Sub

Private Sub Form_Load()
Data2.DatabaseName = App.Path & "\dbnilai.mdb"
Data2.RecordSource = "Mskategori"
Kosong
Data1.DatabaseName = App.Path & "\dbnilai.mdb"
Data1.RecordSource = "tbnilai"
Data1.Refresh
BtnAwal
End Sub

Sub Kosong()
lblId.Caption = ""
Text1 = ""
Text3 = ""
isiCbo
End Sub

Sub Isi()
With Data1.Recordset
    If Not .BOF Then
        lblId.Caption = !id
        Combo1 = !kategori
        Text1 = !keterangan
        Text3 = !bobot
    Else
        Kosong
    End If
End With
End Sub

Sub BtnAwal()
cmd1.Visible = True
cmd2.Caption = "Edit"
cmd3.Caption = "Hapus"
DBGrid1.Enabled = True
Text1.Enabled = False
Combo1.Enabled = False
Text3.Enabled = False
End Sub

Sub BtnSimpan()
cmd1.Visible = False
cmd2.Caption = "Simpan"
cmd3.Caption = "Batal"
DBGrid1.Enabled = False
Text1.Enabled = True
Combo1.Enabled = True
Text3.Enabled = True
End Sub



