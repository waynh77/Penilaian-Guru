VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form kategori_frm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Kriteria Nilai"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   4935
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1440
      TabIndex        =   10
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1440
      TabIndex        =   9
      Top             =   600
      Width           =   975
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2040
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton cmd3 
      Caption         =   "Hapus"
      Height          =   495
      Left            =   3720
      TabIndex        =   5
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton cmd2 
      Caption         =   "Edit"
      Height          =   495
      Left            =   2520
      TabIndex        =   4
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton cmd1 
      Caption         =   "Tambah"
      Height          =   495
      Left            =   1320
      TabIndex        =   3
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   1080
      Width           =   3255
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "kategori_frm.frx":0000
      Height          =   2535
      Left            =   240
      OleObjectBlob   =   "kategori_frm.frx":0014
      TabIndex        =   6
      Top             =   2640
      Width           =   4455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Bobot"
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   8
      Top             =   1560
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Kode "
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   7
      Top             =   600
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Kriteria Nilai"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   825
   End
   Begin VB.Label lblId 
      AutoSize        =   -1  'True
      Caption         =   "Id"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   135
   End
End
Attribute VB_Name = "kategori_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd1_Click()
BtnSimpan
Kosong
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
            !kode = Text2
            !kategori = Text1
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
Kosong
Data1.DatabaseName = App.Path & "\dbnilai.mdb"
Data1.RecordSource = "Mskategori"
Data1.Refresh
BtnAwal
End Sub

Sub Kosong()
lblId.Caption = ""
Text1 = ""
Text2 = ""
Text3 = ""
End Sub

Sub Isi()
With Data1.Recordset
    If Not .BOF Then
        lblId.Caption = !Id
        Text1 = !kategori
        Text2 = !kode
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
Text2.Enabled = False
Text3.Enabled = False
End Sub

Sub BtnSimpan()
cmd1.Visible = False
cmd2.Caption = "Simpan"
cmd3.Caption = "Batal"
DBGrid1.Enabled = False
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
End Sub


