VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Penilaian_frm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Penilaian"
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   7695
   Begin VB.Data Data2 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   420
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   1140
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "Penilaian_frm.frx":0000
      Height          =   5175
      Left            =   240
      OleObjectBlob   =   "Penilaian_frm.frx":0014
      TabIndex        =   5
      Top             =   840
      Width           =   7335
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   420
      Left            =   7320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   240
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Hapus"
      Height          =   495
      Left            =   6000
      TabIndex        =   4
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Edit"
      Height          =   495
      Left            =   4800
      TabIndex        =   3
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Tambah"
      Height          =   495
      Left            =   3600
      TabIndex        =   2
      Top             =   240
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   960
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label lblidperiode 
      Caption         =   "Label2"
      Height          =   255
      Left            =   8880
      TabIndex        =   6
      Top             =   240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Periode"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   540
   End
End
Attribute VB_Name = "Penilaian_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
Data2.RecordSource = "select id from msperiode where periode='" & Combo1.Text & "'"
Data2.Refresh
If Data2.Recordset.BOF Then
    lblidperiode.Caption = ""
Else
    Me.lblidperiode.Caption = Data2.Recordset!id
    isiGrid
End If
End Sub

Private Sub Command1_Click()
With trNilai_frm
    .Top = 0
    .Left = 0
    .lblidperiode.Caption = Me.lblidperiode.Caption
    .Label1.Caption = Combo1.Text
    .lblidTrans.Caption = ""
    .Data1.DatabaseName = App.Path & "\dbnilai.mdb"
    .isiGuru
    .isiKategori
    .Show
End With
End Sub

Private Sub Command2_Click()
With trNilai_frm
    .Top = 0
    .Left = 0
    .lblidperiode.Caption = Me.lblidperiode.Caption
    .Label1.Caption = Combo1.Text
    .lblidTrans.Caption = Data1.Recordset!id
    .Data1.DatabaseName = App.Path & "\dbnilai.mdb"
    .isiGuru
    .isiKategori
    .cboguru.Text = Data1.Recordset!namaguru
    .cbokategori.Text = Data1.Recordset!kategori
    .Text1 = Data1.Recordset!nilai
    .Show
End With
End Sub

Private Sub Command3_Click()
    If Data1.Recordset.BOF Then
        MsgBox "Data kosong", vbInformation, "Validasi"
    Else
        Dim tny As String
        tny = MsgBox("Apakah anda yakin hapus?", vbYesNo, "Hapus")
        If tny = vbYes Then
            Data2.RecordSource = "select * from trnilai where id=" & Data1.Recordset!id
            Data2.Refresh
            Data2.Recordset.Delete
            Data1.Refresh
        End If
    End If
End Sub

Private Sub Form_Activate()
isiGrid
End Sub

Private Sub Form_Load()
Data1.DatabaseName = App.Path & "\dbnilai.mdb"
Data1.RecordSource = "trnilai"
Data1.Refresh
Data2.DatabaseName = App.Path & "\dbnilai.mdb"
Data2.RecordSource = "msperiode"
Data2.Refresh
isiCbo
End Sub

Sub isiCbo()
Combo1.Clear
With Data2.Recordset
    If Not .BOF Then
        .MoveFirst
        Do While Not .EOF
            Combo1.AddItem !periode
            .MoveNext
        Loop
        Combo1.ListIndex = 0
    End If
End With
End Sub

Sub isiGrid()
    Dim sql As String
    'sql = "select trnilai.id,msperiode.periode,msguru.namaguru,mskategori.kategori,nilai from trnilai,msperiode,msguru,mskategori where trnilai.periode=" & Me.lblidperiode.Caption & " and trnilai.periode=msperiode.id and trnilai.guru=msguru.id and trnilai.kriteria=mskategori.id order by trnilai.id"
    sql = "select * from view_nilai where period='" & Me.Combo1.Text & "'"
    Data1.RecordSource = sql
    Data1.Refresh
End Sub
