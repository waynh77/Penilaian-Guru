VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form User_frm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Data USer"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   4830
   Begin VB.TextBox txtnomor 
      Height          =   375
      Left            =   1320
      TabIndex        =   5
      Top             =   600
      Width           =   3255
   End
   Begin VB.TextBox txtnama 
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   1080
      Width           =   3255
   End
   Begin VB.CommandButton cmd1 
      Caption         =   "Tambah"
      Height          =   495
      Left            =   1200
      TabIndex        =   3
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton cmd2 
      Caption         =   "Edit"
      Height          =   495
      Left            =   2400
      TabIndex        =   2
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton cmd3 
      Caption         =   "Hapus"
      Height          =   495
      Left            =   3600
      TabIndex        =   1
      Top             =   1680
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
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1680
      Visible         =   0   'False
      Width           =   1065
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "User_frm.frx":0000
      Height          =   3495
      Left            =   120
      OleObjectBlob   =   "User_frm.frx":0014
      TabIndex        =   0
      Top             =   2280
      Width           =   4455
   End
   Begin VB.Label lblId 
      AutoSize        =   -1  'True
      Caption         =   "Id"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "User Name"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Password"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   690
   End
End
Attribute VB_Name = "User_frm"
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
    If txtnomor = "" Or txtnama = "" Then
        MsgBox "Nomor Induk dan Nama tidak boleh kosong", vbInformation, "Validasi"
    Else
        With Data1.Recordset
            If lblId.Caption = "" Then
                .AddNew
            Else
                .Edit
            End If
            !UserName = txtnomor
            !Password = txtnama
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
Data1.RecordSource = "msuser"
Data1.Refresh
BtnAwal
End Sub

Sub Kosong()
lblId.Caption = ""
txtnomor = ""
txtnama = ""
End Sub

Sub Isi()
With Data1.Recordset
    If Not .BOF Then
        lblId.Caption = !Id
        txtnomor = !UserName
        txtnama = !Password
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
txtnomor.Enabled = False
txtnama.Enabled = False
End Sub

Sub BtnSimpan()
cmd1.Visible = False
cmd2.Caption = "Simpan"
cmd3.Caption = "Batal"
DBGrid1.Enabled = False
txtnomor.Enabled = True
txtnama.Enabled = True
End Sub


